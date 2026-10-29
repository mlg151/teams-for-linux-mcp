'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const { InMemoryTransport } = require('@modelcontextprotocol/server');
const { createServer } = require('../../app/mcp/tools');
const { McpOperationError } = require('../../app/mcp/errors');

async function connect(t, service = {}, runtime = {}) {
  const server = createServer({ run: fn => fn(), status: () => ({ authenticated: false }), ...runtime }, service);
  const [serverTransport, clientTransport] = InMemoryTransport.createLinkedPair();
  let requestId = 0;
  const pending = new Map();
  clientTransport.onmessage = message => {
    if (pending.has(message.id)) { pending.get(message.id)(message); pending.delete(message.id); }
  };
  await server.connect(serverTransport);
  await clientTransport.start();
  t.after(() => server.close());
  async function request(method, params) {
    const id = ++requestId;
    const result = new Promise(resolve => pending.set(id, resolve));
    await clientTransport.send({ jsonrpc: '2.0', id, method, ...(params ? { params } : {}) });
    return result;
  }
  await request('initialize', { protocolVersion: '2025-11-25', capabilities: {}, clientInfo: { name: 'test', version: '1' } });
  await clientTransport.send({ jsonrpc: '2.0', method: 'notifications/initialized' });
  return request;
}

test('MCP discovers all requested tools and exposes no call controls', async t => {
  const request = await connect(t);
  const { result } = await request('tools/list');
  const names = result.tools.map(tool => tool.name);
  for (const name of ['session_status', 'login', 'search_people', 'resolve_chat', 'send_message', 'search_messages', 'search_chats', 'search_teams', 'search_onedrive', 'list_calendar_events']) {
    assert.ok(names.includes(name), name);
  }
  assert.equal(names.some(name => /call|mute|camera|presence/.test(name)), false);
  assert.equal(result.tools.find(tool => tool.name === 'send_message').annotations.readOnlyHint, false);
});

test('MCP resolves a selected person before sending to the returned chat', async t => {
  const userId = '11111111-1111-4111-8111-111111111111';
  const operations = [];
  const request = await connect(t, {
    searchPeople: async () => ({ items: [{ userId, displayName: 'Test Recipient', emails: ['recipient@example.test'] }] }),
    resolveChat: async args => { operations.push(['resolve', args.userId]); return { chatId: '19:verified@thread.v2', participantUserId: args.userId }; },
    sendMessage: async args => { operations.push(['send', args.chatId]); return { sent: true }; },
  });
  const people = await request('tools/call', { name: 'search_people', arguments: { query: 'Test Recipient' } });
  const person = people.result.structuredContent.items[0];
  const chat = await request('tools/call', { name: 'resolve_chat', arguments: { userId: person.userId } });
  assert.deepEqual(operations, [['resolve', userId]]);
  const sent = await request('tools/call', { name: 'send_message', arguments: { chatId: chat.result.structuredContent.chatId, content: 'Test message' } });
  assert.equal(sent.result.structuredContent.sent, true);
  assert.deepEqual(operations, [['resolve', userId], ['send', '19:verified@thread.v2']]);
});

test('MCP input validation prevents ambiguous sends before executing them', async t => {
  let calls = 0;
  const request = await connect(t, { sendMessage: async () => { calls++; return { sent: true }; } });
  for (const args of [{ content: 'hello' }, { chatId: 'a', teamId: 'b', channelId: 'c', content: 'hello' },
    { chatId: 'a', content: ' ' }, { chatId: 'a', replyToId: 'b', content: 'hello' }]) {
    const response = await request('tools/call', { name: 'send_message', arguments: args });
    assert.ok(response.error || response.result?.isError);
  }
  assert.equal(calls, 0);
  assert.equal((await request('tools/call', { name: 'send_message', arguments: { chatId: 'a', content: 'hello' } })).result.structuredContent.sent, true);
  assert.equal(calls, 1);
});

test('MCP validates date ranges and returns actionable permission errors', async t => {
  const request = await connect(t, { searchOneDrive: async () => { throw new McpOperationError('permission_denied', 'Denied'); } });
  const invalid = await request('tools/call', { name: 'list_calendar_events', arguments: { startDateTime: '2026-10-30T00:00:00Z', endDateTime: '2026-10-29T00:00:00Z' } });
  assert.ok(invalid.error || invalid.result?.isError);
  const denied = await request('tools/call', { name: 'search_onedrive', arguments: { query: 'budget' } });
  assert.equal(denied.result.isError, true);
  assert.equal(denied.result.structuredContent.error.code, 'permission_denied');
});

test('MCP login returns a private browser link without requesting credentials', async t => {
  const link = { state: 'awaiting_login', url: 'http://127.0.0.1:8787/#private-flow', expiresAt: '2026-10-29T00:10:00Z' };
  const request = await connect(t, {}, { beginLogin: async () => link });
  const login = await request('tools/call', { name: 'login', arguments: {} });
  assert.deepEqual(login.result.structuredContent, link);
  const listing = await request('tools/list');
  const schema = listing.result.tools.find(tool => tool.name === 'login').inputSchema;
  assert.deepEqual(schema.properties, {});
});
