'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const { SessionGraph, graphUrl } = require('../../app/mcp/graph');

const token = 'private-session-token';
function graphWith(fetchImpl) {
  return new SessionGraph({ requireSession: async () => {} }, {
    tokenClient: { acquireToken: async () => ({ success: true, token }) }, fetchImpl,
  });
}

test('Graph URL validation prevents credentials leaving Microsoft Graph', () => {
  assert.equal(graphUrl('/me'), 'https://graph.microsoft.com/v1.0/me');
  for (const endpoint of ['https://attacker.test/v1.0/me', 'https://graph.microsoft.com.attacker.test/v1.0/me',
    'https://user:pass@graph.microsoft.com/v1.0/me', '/../../oauth2/token', 'https://graph.microsoft.com/beta/me']) {
    assert.throws(() => graphUrl(endpoint));
  }
});

test('Graph pagination uses query-bound opaque cursors and validates nextLink', async () => {
  const urls = [];
  const graph = graphWith(async (url, options) => {
    urls.push(url);
    assert.equal(options.headers.Authorization, `Bearer ${token}`);
    assert.equal(options.redirect, 'error');
    return Response.json(urls.length === 1 ? {
      value: [{ id: 'first' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/chats?$skiptoken=second',
    } : { value: [{ id: 'second' }] });
  });
  const first = await graph.collection('/me/chats');
  assert.equal(JSON.stringify(first).includes(token), false);
  assert.equal(first.items[0].id, 'first');
  await assert.rejects(graph.collection('/me/drive/root/children', { cursor: first.nextCursor }), { code: 'invalid_cursor' });
  assert.equal((await graph.collection('/me/chats', { cursor: first.nextCursor })).items[0].id, 'second');
  assert.match(urls[1], /skiptoken=second/);
  const malicious = graphWith(async () => Response.json({ value: [], '@odata.nextLink': 'https://attacker.test/v1.0/me' }));
  await assert.rejects(malicious.collection('/me/chats'), { code: 'invalid_endpoint' });
});

test('Graph never retries an uncertain send and never exposes token/error response details', async () => {
  let calls = 0;
  const graph = graphWith(async () => { calls++; throw new Error(`Network error with ${token}`); });
  await assert.rejects(graph.request('/chats/id/messages', { method: 'POST', body: { body: { content: 'hello' } } }), error => {
    assert.equal(error.code, 'outcome_unknown');
    assert.equal(error.message.includes(token), false);
    return true;
  });
  assert.equal(calls, 1);
});

test('read-only POST searches report ordinary network failures', async () => {
  const graph = graphWith(async () => { throw new Error('offline'); });
  await assert.rejects(graph.request('/search/query', { method: 'POST', readOnly: true }), { code: 'request_failed' });
});

test('permission errors are explicit and rate limits retain retry-after', async () => {
  const denied = graphWith(async () => Response.json({ error: { message: token } }, { status: 403 }));
  await assert.rejects(denied.request('/me/chats'), error => error.code === 'permission_denied' && !error.message.includes(token));
  const limited = graphWith(async () => Response.json({}, { status: 429, headers: { 'retry-after': '15' } }));
  await assert.rejects(limited.request('/me/chats'), error => error.code === 'rate_limited' && error.details.retryAfter === '15');
});

test('cancellation during token acquisition prevents the network write', async () => {
  const controller = new AbortController();
  let calls = 0;
  const graph = new SessionGraph({ requireSession: async () => {}, signal: controller.signal }, {
    tokenClient: { acquireToken: async () => { controller.abort(); return { success: true, token }; } },
    fetchImpl: async () => { calls++; return Response.json({ id: 'should-not-send' }); },
  });
  await assert.rejects(graph.request('/chats/id/messages', { method: 'POST' }), { code: 'cancelled' });
  assert.equal(calls, 0);
});

test('invalid API collections are errors rather than false empty search results', async () => {
  const graph = graphWith(async () => Response.json({ unexpected: true }));
  await assert.rejects(graph.collection('/me/chats'), { code: 'invalid_response' });
  const malformed = graphWith(async () => new Response('<html>not Graph data</html>', { status: 200 }));
  await assert.rejects(malformed.request('/me/chats'), { code: 'invalid_response' });
});
