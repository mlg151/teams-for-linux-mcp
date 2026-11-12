'use strict';

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

module.exports = { Microsoft365Service, renderMarkdown };
