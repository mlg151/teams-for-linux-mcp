'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { Microsoft365Service, SMALL_UPLOAD_LIMIT, UPLOAD_CHUNK_SIZE } = require('../../app/mcp/service');
const { McpOperationError } = require('../../app/mcp/errors');

test('people lookup returns exact recipient identifiers and safely quotes the query', async () => {
  let endpoint;
  const service = new Microsoft365Service({}, { graph: { request: async url => {
    endpoint = url;
    return { value: [{ id: 'person-id', displayName: 'Test Recipient', userPrincipalName: 'recipient@example.test', scoredEmailAddresses: [{ address: 'recipient@example.test' }] }] };
  } }, browser: {} });
  const result = await service.searchPeople({ query: 'Test "Recipient"', limit: 5 });
  assert.equal(new URL(`https://graph.microsoft.com${endpoint}`).searchParams.get('$search'), '"Test \\"Recipient\\""');
  assert.equal(result.items[0].userId, 'person-id');
  assert.deepEqual(result.items[0].emails, ['recipient@example.test']);
});

test('chat resolution requires a verified result and never sends a message', async () => {
  let verified = false;
  const service = new Microsoft365Service({ requireSession: async () => {} }, {
    graph: { getTokenClient: async () => ({ resolveConversation: async () => verified ? { success: true, conversationId: 'verified-chat' } : { success: false } }) }, browser: {},
  });
  await assert.rejects(service.resolveChat({ userId: 'person' }), { code: 'chat_resolution_failed' });
  verified = true;
  assert.deepEqual(await service.resolveChat({ userId: 'person' }), { source: 'teams_session', chatId: 'verified-chat', participantUserId: 'person' });
});

test('permission denied falls back to the requested browser category; failures do not', async () => {
  const calls = [];
  const service = new Microsoft365Service({}, {
    graph: { request: async () => { throw new McpOperationError('permission_denied', 'Denied'); } },
    browser: { search: async (...args) => { calls.push(args); return { source: 'teams_browser' }; } },
  });
  assert.equal((await service.searchMessages({ query: 'quarterly', limit: 5 })).source, 'teams_browser');
  assert.deepEqual(calls, [['messages', 'quarterly', 5]]);
  await assert.rejects(service.searchMessages({ query: 'quarterly', backend: 'graph' }), { code: 'permission_denied' });
  service.graph.request = async () => { throw new McpOperationError('rate_limited', 'Wait'); };
  await assert.rejects(service.searchMessages({ query: 'quarterly' }), { code: 'rate_limited' });
  assert.equal(calls.length, 1);
});

test('browser search is selectable and never silently ignores pagination', async () => {
  const service = new Microsoft365Service({}, {
    graph: { request: async () => assert.fail('Graph should not be called') },
    browser: { search: async () => ({ source: 'teams_browser' }) },
  });
  assert.equal((await service.searchMessages({ query: 'q', backend: 'browser' })).source, 'teams_browser');
  await assert.rejects(service.searchMessages({ query: 'q', backend: 'browser', offset: 20 }), { code: 'unsupported_pagination' });
});

test('OneDrive OData search encodes apostrophes and URL control characters', async () => {
  let endpoint;
  const service = new Microsoft365Service({}, {
    graph: { collection: async url => { endpoint = url; return { items: [] }; } }, browser: {},
  });
  await service.searchOneDrive({ query: "O'Brien & budget?#", limit: 5 });
  assert.equal(endpoint, "/me/drive/root/search(q='O''Brien%20%26%20budget%3F%23')?$top=5");
});

test('messages target exactly the encoded chat or channel; a send returns the created message', async () => {
  const calls = [];
  const service = new Microsoft365Service({}, {
    graph: { request: async (endpoint, options) => { calls.push({ endpoint, options }); return { id: 'created' }; } }, browser: {},
  });
  const result = await service.sendMessage({ chatId: '19:chat@thread.v2', content: 'hello <team>' });
  assert.equal(result.message.id, 'created');
  assert.equal(calls[0].endpoint, '/chats/19%3Achat%40thread.v2/messages');
  assert.deepEqual(calls[0].options.body.body, { contentType: 'text', content: 'hello <team>' });
  await service.sendMessage({ teamId: 'team', channelId: 'channel', replyToId: 'parent', content: 'reply' });
  assert.equal(calls[1].endpoint, '/teams/team/channels/channel/messages/parent/replies');
});

test('chat search scans subsequent pages and matches participant names', async () => {
  let calls = 0;
  const service = new Microsoft365Service({}, {
    graph: { collection: async () => ++calls === 1 ? { items: [{ topic: 'Other' }], nextCursor: 'next' } :
      { items: [{ id: 'match', members: [{ displayName: 'Maria Garcia' }] }] } }, browser: {},
  });
  const result = await service.searchDirectory('chats', { query: 'garcia' });
  assert.equal(result.items[0].id, 'match');
  assert.equal(calls, 2);
  assert.equal(result.truncated, false);
});

test('calendar view sends an offset-bearing range and a calendar-specific cursor scope', async () => {
  let endpoint;
  const service = new Microsoft365Service({}, {
    graph: { collection: async url => { endpoint = url; return { items: [] }; } }, browser: {},
  });
  await service.listCalendarEvents({ startDateTime: '2026-10-29T09:00:00-04:00', endDateTime: '2026-10-30T09:00:00-04:00', calendarId: 'cal/1' });
  const url = new URL(`https://graph.microsoft.com/v1.0${endpoint}`);
  assert.equal(url.pathname, '/v1.0/me/calendars/cal%2F1/calendarView');
  assert.equal(url.searchParams.get('startDateTime'), '2026-10-29T09:00:00-04:00');
});

test('send_message renders markdown as HTML, passes HTML through, and keeps text plain', async () => {
  const bodies = [];
  const service = new Microsoft365Service({}, {
    graph: { request: async (endpoint, options) => { bodies.push(options.body.body); return { id: 'created' }; } }, browser: {},
  });
  const plain = await service.sendMessage({ chatId: '19:chat@thread.v2', content: 'hello <team>' });
  assert.deepEqual(bodies[0], { contentType: 'text', content: 'hello <team>' });
  assert.equal(plain.format, 'text');
  const markdown = await service.sendMessage({ chatId: '19:chat@thread.v2', content: '# Title\n\n- **bold**\n- `code`', format: 'markdown' });
  assert.equal(markdown.format, 'html');
  assert.equal(bodies[1].contentType, 'html');
  assert.match(bodies[1].content, /<h1>Title<\/h1>/);
  assert.match(bodies[1].content, /<strong>bold<\/strong>/);
  assert.match(bodies[1].content, /<code>code<\/code>/);
  await service.sendMessage({ chatId: '19:chat@thread.v2', content: '<b>raw</b>', format: 'html' });
  assert.deepEqual(bodies[2], { contentType: 'html', content: '<b>raw</b>' });
});

test('site search queries the sites collection with the raw search term', async () => {
  let endpoint;
  const service = new Microsoft365Service({}, {
    graph: { collection: async url => { endpoint = url; return { items: [{ id: 'site-1', displayName: 'Project A', description: 'Docs', webUrl: 'https://a.sharepoint.com/sites/ProjectA' }] }; } }, browser: {},
  });
  const result = await service.searchSites({ query: 'project a', limit: 5 });
  assert.equal(endpoint, '/sites?search=project%20a');
  assert.equal(result.items[0].siteId, 'site-1');
  assert.equal(result.truncated, false);
});

test('drive listing filters OneDrive and SharePoint drives client-side', async () => {
  const service = new Microsoft365Service({}, {
    graph: { request: async () => ({ value: [
      { id: 'drive-1', name: 'OneDrive - Contoso', description: null, url: 'https://contoso-my.sharepoint.com/personal/me', driveType: 'business' },
      { id: 'drive-2', name: 'Project A', description: 'site', url: 'https://contoso.sharepoint.com/sites/ProjectA', driveType: 'business' },
      { id: 'drive-3', name: 'Other', description: '', url: 'https://contoso.sharepoint.com/sites/Other', driveType: 'business' },
    ] }) }, browser: {},
  });
  const result = await service.listDrives({ query: 'project' });
  assert.deepEqual(result.items.map(item => item.driveId), ['drive-2']);
  const all = await service.listDrives({});
  assert.equal(all.items.length, 3);
  assert.equal(all.items[0].name, 'OneDrive - Contoso');
});

test('drive folder listing addresses root and relative paths with the colon path syntax', async () => {
  const endpoints = [];
  const service = new Microsoft365Service({}, {
    graph: { collection: async url => { endpoints.push(url); return { items: [{ id: 'item-1', name: 'Notes', folder: {} }, { id: 'item-2', name: 'Notes.txt', file: { size: 42, mimeType: 'text/plain' } }] }; } }, browser: {},
  });
  const result = await service.listDriveItems({ driveId: 'drive-1' });
  assert.equal(endpoints[0], '/drives/drive-1/root/children?$top=20');
  assert.equal(result.items[0].kind, 'folder');
  assert.equal(result.items[1].kind, 'file');
  assert.equal(result.items[1].size, 42);
  await service.listDriveItems({ driveId: 'drive-1', folderPath: 'Reports/2026', limit: 5 });
  assert.equal(endpoints[1], '/drives/drive-1/root:/Reports/2026:/children?$top=5');
  await service.listDriveItems({ driveId: 'drive-1', folderPath: "O'Brien/2026 report", limit: 3 });
  assert.equal(endpoints[2], "/drives/drive-1/root:/O'Brien/2026%20report:/children?$top=3");
  await assert.rejects(service.downloadFile({ driveId: 'drive-1', filePath: 'root' }), { code: 'invalid_path' });
});

test('download_file streams bytes to an owned directory and reports local metadata', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-download-'));
  const bytes = Buffer.alloc(300 * 1024, 7);
  const service = new Microsoft365Service({ signal: undefined }, {
    graph: { requestStream: async endpoint => {
      assert.equal(endpoint, '/drives/drive-1/root:/docs/file.txt:/content');
      return new Response(bytes, { headers: { 'Content-Type': 'text/plain; charset=utf-8' } });
    } }, browser: {},
  });
  const result = await service.downloadFile({ driveId: 'drive-1', filePath: 'docs/file.txt', destDir: path.join(dir, 'out') });
  assert.equal(result.downloaded, true);
  assert.equal(result.path, path.join(dir, 'out', 'file.txt'));
  assert.ok(fs.readFileSync(result.path).equals(bytes));
  assert.equal(result.mimeType, 'text/plain');
  assert.equal(result.size, bytes.length);
  fs.rmSync(dir, { recursive: true, force: true });
});

test('download_file refuses symlinked destinations and enforces the size cap', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-download2-'));
  const real = path.join(dir, 'real');
  fs.mkdirSync(real);
  const link = path.join(dir, 'link');
  fs.symlinkSync(real, link);
  const service = new Microsoft365Service({ signal: undefined }, {
    graph: { requestStream: async () => new Response(new Uint8Array(2 * 1024 * 1024)) }, browser: {},
  });
  await assert.rejects(service.downloadFile({ driveId: 'd', filePath: 'a.txt', destDir: link }), { code: 'invalid_path' });
  const original = process.env.TEAMS_MCP_MAX_DOWNLOAD_MB;
  process.env.TEAMS_MCP_MAX_DOWNLOAD_MB = '1';
  try {
    await assert.rejects(service.downloadFile({ driveId: 'd', filePath: 'a.txt', destDir: real }), { code: 'file_too_large' });
    assert.equal(fs.readdirSync(real).length, 0);
  } finally {
    if (original === undefined) delete process.env.TEAMS_MCP_MAX_DOWNLOAD_MB; else process.env.TEAMS_MCP_MAX_DOWNLOAD_MB = original;
  }
  fs.rmSync(dir, { recursive: true, force: true });
});

test('small uploads PUT content once and refuse existing names without overwrite', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-upload-'));
  const source = path.join(dir, 'notes.txt');
  fs.writeFileSync(source, 'hello world\n');
  const calls = [];
  // Destroyed request-body streams still settle their pending open; wait for
  // them to close before the fixture directory is removed.
  const streams = [];
  const service = new Microsoft365Service({}, {
    graph: { requestStream: async (endpoint, options) => {
      calls.push({ endpoint, method: options.method, headers: options.headers, length: options.options.contentLength });
      const body = options.options.body;
      if (typeof body?.destroy === 'function') streams.push(body);
      body?.destroy?.();
      if (endpoint.endsWith('/content')) return new Response(JSON.stringify({ id: 'item-1', name: 'notes.txt' }), { status: 200 });
      return new Response('{}', { status: 200 });
    } }, browser: {},
  });
  const result = await service.uploadFile({ driveId: 'drive-1', filePath: 'docs/notes.txt', sourcePath: source });
  assert.equal(result.method, 'put');
  assert.equal(result.item.id, 'item-1');
  assert.equal(calls[0].endpoint, '/drives/drive-1/root:/docs/notes.txt:/content');
  assert.equal(calls[0].method, 'PUT');
  assert.equal(calls[0].headers['if-match'], '0');
  assert.equal(calls[0].length, 12);
  await service.uploadFile({ driveId: 'drive-1', filePath: 'docs/notes.txt', sourcePath: source, overwrite: true });
  assert.equal(calls[1].headers['if-match'], undefined);
  const existing = new Microsoft365Service({}, {
    graph: { requestStream: async (endpoint, options) => {
      const body = options.options.body;
      if (typeof body?.destroy === 'function') streams.push(body);
      body?.destroy?.();
      return new Response(JSON.stringify({ error: { message: 'precondition failed' } }), { status: 412 });
    } }, browser: {},
  });
  await assert.rejects(existing.uploadFile({ driveId: 'drive-1', filePath: 'docs/notes.txt', sourcePath: source }), { code: 'file_exists' });
  await assert.rejects(service.uploadFile({ driveId: 'drive-1', filePath: 'docs/missing.txt', sourcePath: dir }), { code: 'invalid_path' });
  await Promise.all(streams.map(body => new Promise(resolve => body.once('close', resolve))));
  fs.rmSync(dir, { recursive: true, force: true });
});

test('large uploads use a chunked upload session with byte ranges', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-upload2-'));
  const source = path.join(dir, 'large.bin');
  const size = 2 * SMALL_UPLOAD_LIMIT + 12345;
  const fd = fs.openSync(source, 'w');
  fs.ftruncateSync(fd, size);
  fs.closeSync(fd);
  const calls = [];
  const streams = [];
  const service = new Microsoft365Service({}, {
    graph: { requestStream: async (endpoint, options) => {
      const body = options.options.body;
      calls.push({ endpoint, method: options.method, url: options.url, headers: options.headers, body });
      if (typeof body?.destroy === 'function') streams.push(body);
      body?.destroy?.();
      if (endpoint) return new Response(JSON.stringify({ uploadUrl: 'https://graph.sharepointonline.com/v2.1/_api/upload/session' }), { status: 200 });
      return new Response(JSON.stringify({ id: 'item-1', name: 'large.bin' }), { status: 200 });
    } }, browser: {},
  });
  const result = await service.uploadFile({ driveId: 'drive-1', filePath: 'big/large.bin', sourcePath: source });
  assert.equal(result.method, 'upload_session');
  assert.equal(result.size, size);
  assert.equal(calls[0].endpoint, '/drives/drive-1/root:/big/large.bin:/createUploadSession');
  assert.equal(calls[0].method, 'POST');
  assert.deepEqual(JSON.parse(calls[0].body), { item: { fields: { name: 'large.bin' } } });
  const chunks = calls.slice(1);
  assert.equal(chunks.length, Math.ceil(size / UPLOAD_CHUNK_SIZE));
  for (let index = 0; index < chunks.length; index++) {
    const start = index * UPLOAD_CHUNK_SIZE;
    const end = Math.min(size, start + UPLOAD_CHUNK_SIZE) - 1;
    assert.equal(chunks[index].url, 'https://graph.sharepointonline.com/v2.1/_api/upload/session');
    assert.equal(chunks[index].method, 'PUT');
    assert.equal(chunks[index].headers['Content-Range'], `bytes ${start}-${end}/${size}`);
    assert.equal(chunks[index].headers['Content-Length'], String(end - start + 1));
  }
  await Promise.all(streams.map(body => new Promise(resolve => body.once('close', resolve))));
  fs.rmSync(dir, { recursive: true, force: true });
});
