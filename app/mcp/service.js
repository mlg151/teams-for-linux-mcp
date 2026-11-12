'use strict';

const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { Readable } = require('node:stream');
const { marked } = require('marked');
const { SessionGraph, canUseBrowser } = require('./graph');
const { TeamsBrowser } = require('./browser');
const { McpOperationError } = require('./errors');

const id = encodeURIComponent;
const literal = text => encodeURIComponent(text.replaceAll("'", "''"));

function renderMarkdown(content) {
  return marked.parse(content, { async: false });
}

// Drive items below the root use the colon path syntax: root:/{path}, where
// each path segment is percent-encoded and separators stay literal slashes.
// See "How to address resources" in the Microsoft Graph docs.
function drivePathKey(clean) {
  const segments = clean.split('/');
  if (segments.some(segment => !segment)) throw new McpOperationError('invalid_path', 'The path must not contain empty segments.');
  return segments.map(encodeURIComponent).join('/');
}

function driveChildrenEndpoint(driveId, folderPath, limit) {
  const clean = (folderPath || 'root').replace(/^\//, '');
  const where = clean === 'root' ? 'root/children' : `root:/${drivePathKey(clean)}:/children`;
  return `/drives/${id(driveId)}/${where}?$top=${limit}`;
}

function driveContentEndpoint(driveId, filePath) {
  const clean = (filePath || '').replace(/^\//, '');
  if (!clean || clean === 'root') throw new McpOperationError('invalid_path', 'The path must identify a file below the drive root, not the root itself.');
  return `/drives/${id(driveId)}/root:/${drivePathKey(clean)}:/content`;
}

function maxDownloadBytes() {
  const configured = Number(process.env.TEAMS_MCP_MAX_DOWNLOAD_MB);
  const megabytes = Number.isInteger(configured) && configured > 0 ? configured : 4096;
  return megabytes * 1024 * 1024;
}

// The destination must be reachable without following a symlink, and the
// directory we write into must be owned by this user, so another OS account
// can neither read the download nor plant a file in its place. Higher-level
// components may be root-owned (/, /tmp, /home, ...). The service runs under
// umask 077, so created files are 0600 and owned by this user.
function secureDownloadDirectory(directory) {
  const resolved = path.resolve(directory);
  const segments = resolved.split(path.sep).filter(Boolean);
  let cursor = path.sep;
  for (const [index, segment] of segments.entries()) {
    cursor = path.join(cursor, segment);
    const stat = fs.lstatSync(cursor);
    if (stat.isSymbolicLink()) throw new McpOperationError('invalid_path', 'The destination path must not contain symlinks.');
    if (!stat.isDirectory()) throw new McpOperationError('invalid_path', 'The destination directory does not exist.');
    if (index === segments.length - 1 && stat.uid !== process.getuid()) {
      throw new McpOperationError('invalid_path', 'The destination directory must be owned by this user.');
    }
  }
  return resolved;
}

function defaultDownloadsDirectory() {
  return path.join(os.homedir(), '.local', 'share', 'teams-mcp', 'downloads');
}

function uniqueDestination(directory, filePath) {
  const raw = path.basename((filePath || '').replace(/\\/g, '/'));
  const base = raw.replace(/[^\w.@+-]/g, '_').slice(0, 180) || 'download';
  let destination = path.join(directory, base);
  for (let index = 0; index < 100 && fs.existsSync(destination); index++) {
    const dot = base.lastIndexOf('.');
    const stem = dot > 0 ? base.slice(0, dot) : base;
    const ext = dot > 0 ? base.slice(dot) : '';
    destination = path.join(directory, `${stem}-${Date.now().toString(36)}${index ? `-${index}` : ''}${ext}`);
  }
  return destination;
}

class Microsoft365Service {
  constructor(runtime, { graph = new SessionGraph(runtime), browser = new TeamsBrowser(runtime) } = {}) {
    this.runtime = runtime;
    this.graph = graph;
    this.browser = browser;
  }

  async withBrowser(backend, apiOperation, browserOperation) {
    if (backend === 'browser') return browserOperation();
    try { return await apiOperation(); }
    catch (error) {
      if (backend === 'graph' || !canUseBrowser(error)) throw error;
      return browserOperation();
    }
  }

  async searchPeople({ query, limit = 20 }) {
    const params = new URLSearchParams({ $search: JSON.stringify(query), $top: String(limit) });
    const data = await this.graph.request(`/me/people?${params}`);
    if (!Array.isArray(data?.value)) throw new McpOperationError('invalid_response', 'Microsoft Graph did not return a people collection.');
    return { source: 'graph', items: data.value.map(person => ({
      userId: person.id, displayName: person.displayName, userPrincipalName: person.userPrincipalName,
      emails: (person.scoredEmailAddresses || []).map(email => email.address).filter(Boolean),
    })) };
  }

  async resolveChat({ userId }) {
    await this.runtime.requireSession();
    const client = await this.graph.getTokenClient();
    const result = await client.resolveConversation(userId);
    if (!result?.success || !result.conversationId) {
      throw new McpOperationError('chat_resolution_failed', 'Could not resolve a verified one-to-one chat with this person. No message was sent.');
    }
    return { source: 'teams_session', chatId: result.conversationId, participantUserId: userId };
  }

  async searchMessages({ query, limit = 20, offset = 0, backend = 'auto' }) {
    return this.withBrowser(backend, async () => {
      const data = await this.graph.request('/search/query', { method: 'POST', readOnly: true, body: {
        requests: [{ entityTypes: ['chatMessage'], query: { queryString: query }, from: offset, size: limit }],
      } });
      const containers = (data?.value || []).flatMap(value => value.hitsContainers || []);
      const moreResultsAvailable = containers.some(container => container.moreResultsAvailable);
      return { source: 'graph', query, items: containers.flatMap(container => container.hits || []),
        moreResultsAvailable, ...(moreResultsAvailable ? { nextOffset: offset + limit } : {}) };
    }, () => {
      if (offset) throw new McpOperationError('unsupported_pagination', 'Browser search returns the visible first page. Graph pagination is unavailable for this session.');
      return this.browser.search('messages', query, limit);
    });
  }

  async searchDirectory(category, { query, limit = 20, backend = 'auto' }) {
    return this.withBrowser(backend, async () => {
      const endpoint = category === 'chats' ? '/me/chats?$expand=members&$top=50' : '/me/joinedTeams';
      const matches = [];
      let cursor;
      let pages = 0;
      const needle = query.toLocaleLowerCase();
      do {
        const result = await this.graph.collection(endpoint, { cursor });
        for (const item of result.items) {
          const haystack = [item.topic, item.displayName, item.description,
            ...(item.members || []).flatMap(member => [member.displayName, member.email])].filter(Boolean).join(' ').toLocaleLowerCase();
          if (haystack.includes(needle)) matches.push(item);
        }
        cursor = result.nextCursor;
        pages++;
      } while (cursor && pages < 10 && matches.length < limit);
      return { source: 'graph', query, items: matches.slice(0, limit), truncated: !!cursor || matches.length > limit,
        coverage: 'Names, descriptions, and available chat member names/emails; scans up to ten API pages. Narrow the query if truncated.' };
    }, () => this.browser.search(category, query, limit));
  }

  async searchOneDrive({ query, limit = 20, backend = 'auto', cursor }) {
    return this.withBrowser(backend,
      () => this.graph.collection(`/me/drive/root/search(q='${literal(query)}')?$top=${limit}`, { cursor }),
      () => {
        if (cursor) throw new McpOperationError('unsupported_pagination', 'The browser cannot continue a Graph cursor. Start a new browser search without a cursor.');
        return this.browser.searchOneDrive(query, limit);
      });
  }

  listOneDrive({ folderId = 'root', limit = 20, cursor }) {
    const folder = folderId === 'root' ? 'root' : `items/${id(folderId)}`;
    return this.graph.collection(`/me/drive/${folder}/children?$top=${limit}`, { cursor });
  }

  async searchSites({ query, limit = 20 }) {
    const data = await this.graph.collection(`/sites?search=${encodeURIComponent(query)}`);
    const items = data.items.slice(0, limit).map(site => ({
      siteId: site.id, displayName: site.displayName, description: site.description, webUrl: site.webUrl,
    }));
    return { source: 'graph', query, items, truncated: data.items.length > limit || !!data.nextCursor,
      coverage: 'SharePoint sites this account can access that match the query. Use the returned siteId with list_drive_items or upload_file.' };
  }

  async listDrives({ query, limit = 20 }) {
    const data = await this.graph.request('/me/drives');
    if (!Array.isArray(data?.value)) throw new McpOperationError('invalid_response', 'Microsoft Graph did not return a drives collection.');
    const needle = (query || '').toLocaleLowerCase();
    const matches = data.value.filter(drive => !needle ||
      [drive.name, drive.description].some(text => (text || '').toLocaleLowerCase().includes(needle)));
    return { source: 'graph', query: query || null,
      items: matches.slice(0, limit).map(drive => ({ driveId: drive.id, name: drive.name, description: drive.description, url: drive.url, driveType: drive.driveType })),
      truncated: matches.length > limit,
      coverage: 'OneDrive and SharePoint site drives this account can access. Use the returned driveId with list_drive_items, download_file, or upload_file.' };
  }

  async listDriveItems({ driveId, folderPath = 'root', limit = 20, cursor }) {
    const data = await this.graph.collection(driveChildrenEndpoint(driveId, folderPath, limit), { cursor });
    return { source: 'graph', folder: (folderPath || 'root').replace(/^\//, ''),
      items: data.items.map(item => ({
        id: item.id, name: item.name, kind: item.folder ? 'folder' : 'file',
        ...(item.file ? { size: item.file.size, mimeType: item.file.mimeType } : {}),
        createdDateTime: item.createdDateTime, lastModifiedDateTime: item.lastModifiedDateTime, webUrl: item.webUrl,
      })), ...(data.nextCursor ? { nextCursor: data.nextCursor } : {}) };
  }

  async downloadFile({ driveId, filePath, destDir }) {
    const endpoint = driveContentEndpoint(driveId, filePath);
    const response = await this.graph.requestStream(endpoint);
    const releaseBody = () => { response.body?.cancel().catch(() => {}); };
    let directory;
    let fd;
    try {
      if (!response.body) throw new McpOperationError('invalid_response', 'Microsoft Graph returned no file content.');
      directory = destDir ? path.resolve(destDir) : defaultDownloadsDirectory();
      if (!fs.existsSync(directory)) fs.mkdirSync(directory, { recursive: true, mode: 0o700 });
      directory = secureDownloadDirectory(directory);
      const destination = uniqueDestination(directory, filePath);
      const maxBytes = maxDownloadBytes();
      // O_NOFOLLOW closes the race where another local user plants a symlink
      // at the destination name in a shared directory before we open it.
      fd = fs.openSync(destination, fs.constants.O_CREAT | fs.constants.O_WRONLY | fs.constants.O_TRUNC | fs.constants.O_NOFOLLOW, 0o600);
      const writer = fs.createWriteStream(destination, { fd });
      const stream = Readable.fromWeb(response.body, { signal: this.runtime.signal });
      let bytes = 0;
      let failed;
      try {
        for await (const chunk of stream) {
          bytes += chunk.length;
          if (bytes > maxBytes) {
            failed = new McpOperationError('file_too_large', `The file exceeds the ${Math.floor(maxBytes / 1024 / 1024)} MB download limit. Set TEAMS_MCP_MAX_DOWNLOAD_MB in service.env to allow larger files.`);
            break;
          }
          if (!writer.write(chunk)) await new Promise(resolve => writer.once('drain', resolve));
        }
        await new Promise((resolve, reject) => { writer.end(); writer.once('finish', resolve); writer.once('error', reject); });
      } catch (error) {
        failed = failed || error;
      } finally {
        stream.destroy();
        writer.destroy();
      }
      if (failed) {
        await fs.promises.unlink(destination).catch(() => {});
        if (failed.name === 'AbortError' || this.runtime.signal?.aborted) throw new McpOperationError('cancelled', 'The download was cancelled before it finished. No file was kept.');
        if (failed instanceof McpOperationError) throw failed;
        throw new McpOperationError('download_failed', 'The file download failed; no partial file was kept. The remote file may still be available.');
      }
      return { source: 'graph', downloaded: true, path: destination, name: path.basename(destination),
        mimeType: (response.headers.get('content-type') || 'application/octet-stream').split(';')[0], size: bytes };
    } catch (error) {
      releaseBody();
      if (fd !== undefined) {
        try { fs.closeSync(fd); } catch { /* the writer already closed the descriptor */ }
      }
      if (error instanceof McpOperationError) throw error;
      if (error.code === 'ELOOP' || error.code === 'EACCES' || error.code === 'EEXIST') {
        throw new McpOperationError('invalid_path', 'The destination file could not be created safely.');
      }
      throw new McpOperationError('download_failed', 'The file download failed; no partial file was kept. The remote file may still be available.');
    }
  }

  listChannels({ teamId, cursor }) {
    return this.graph.collection(`/teams/${id(teamId)}/channels?$select=id,displayName,description,membershipType,webUrl`, { cursor });
  }

  readMessages({ chatId, teamId, channelId, limit = 20, cursor }) {
    const endpoint = chatId ? `/chats/${id(chatId)}/messages` : `/teams/${id(teamId)}/channels/${id(channelId)}/messages`;
    return this.graph.collection(`${endpoint}?$top=${limit}`, { cursor });
  }

  async sendMessage({ chatId, teamId, channelId, replyToId, content, format = 'text' }) {
    let endpoint = chatId ? `/chats/${id(chatId)}/messages` : `/teams/${id(teamId)}/channels/${id(channelId)}/messages`;
    if (replyToId) endpoint += `/${id(replyToId)}/replies`;
    const body = format === 'markdown' ? { contentType: 'html', content: renderMarkdown(content) } :
      format === 'html' ? { contentType: 'html', content } : { contentType: 'text', content };
    const data = await this.graph.request(endpoint, { method: 'POST', body: { body } });
    if (!data?.id) throw new McpOperationError('outcome_unknown', 'Microsoft did not return a message ID. Check Teams before retrying the send.');
    return { source: 'graph', sent: true, format: body.contentType, message: data };
  }

  listCalendarEvents({ startDateTime, endDateTime, calendarId, limit = 20, cursor }) {
    const calendar = calendarId ? `/me/calendars/${id(calendarId)}` : '/me/calendar';
    const params = new URLSearchParams({ startDateTime, endDateTime, $top: limit, $orderby: 'start/dateTime' });
    return this.graph.collection(`${calendar}/calendarView?${params}`, { cursor });
  }

  async createCalendarEvent({ calendarId, event }) {
    const calendar = calendarId ? `/me/calendars/${id(calendarId)}` : '/me/calendar';
    return { source: 'graph', event: await this.graph.request(`${calendar}/events`, { method: 'POST', body: event }) };
  }

  async updateCalendarEvent({ eventId, updates }) {
    return { source: 'graph', event: await this.graph.request(`/me/events/${id(eventId)}`, { method: 'PATCH', body: updates }) };
  }
}

module.exports = { Microsoft365Service, renderMarkdown, driveContentEndpoint, secureDownloadDirectory, defaultDownloadsDirectory };
