'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const { EventEmitter } = require('node:events');
const http = require('node:http');
const { LoginPortal, portalOptions } = require('../../app/mcp/loginPortal');

async function fixture(t, options = {}) {
  const page = new EventEmitter();
  const actions = [];
  page.url = () => 'https://login.microsoftonline.com/tenant/login?private_auth_code=never-return';
  page.mainFrame = () => 'main';
  page.evaluate = async () => ({ width: 800, height: 600 });
  page.screenshot = async () => Buffer.from('private-screen-image');
  page.frames = () => [{ evaluate: async () => true }];
  page.mouse = { click: async (x, y) => actions.push(['click', x, y]), wheel: async (x, y) => actions.push(['scroll', x, y]) };
  page.keyboard = { insertText: async text => actions.push(['type', text]), press: async key => actions.push(['key', key]) };
  const runtime = { run: operation => operation(), status: async () => ({ authenticated: false }), getLoginPage: async () => page,
    captureLoginScreen: async () => Buffer.from('private-screen-image').toString('base64'), focusLoginPage: async () => {} };
  const portal = new LoginPortal(runtime, options);
  t.after(() => portal.close());
  const login = await portal.start();
  const url = new URL(login.url);
  async function request(route, body, overrides = {}) {
    return fetch(`${url.origin}${route}`, {
      method: body === undefined ? 'GET' : 'POST',
      headers: { Authorization: `Bearer ${url.hash.slice(1)}`, ...(body === undefined ? {} : { Origin: url.origin, 'Content-Type': 'application/json' }), ...overrides },
      ...(body === undefined ? {} : { body: JSON.stringify(body) }),
    });
  }
  return { portal, login, runtime, page, actions, request, url };
}

test('login link serves a private sign-in view, while session images require authorization', async t => {
  const { request, url, login } = await fixture(t);
  assert.equal(login.state, 'awaiting_login');
  assert.equal(login.url.includes('never-return'), false);
  const shell = await fetch(url.origin);
  assert.equal(shell.status, 200);
  assert.match(shell.headers.get('content-security-policy'), /frame-ancestors 'none'/);
  assert.equal(shell.headers.get('cache-control'), 'no-store');
  assert.equal((await fetch(`${url.origin}/api/state`)).status, 401);
  const response = await request('/api/state');
  assert.equal(response.status, 200);
  const state = await response.json();
  assert.equal(state.origin, 'https://login.microsoftonline.com');
  assert.equal(state.image, Buffer.from('private-screen-image').toString('base64'));
  assert.equal(JSON.stringify(state).includes('never-return'), false);
});

test('browser actions stay on the current sign-in page and never echo typed credentials', async t => {
  const { request, page, actions } = await fixture(t);
  const { viewId } = await (await request('/api/state')).json();
  const password = 'fixture-password-do-not-echo';
  const response = await request('/api/action', { viewId, action: { type: 'type', text: password } });
  assert.equal(response.status, 200);
  assert.equal((await response.text()).includes(password), false);
  assert.deepEqual(actions, [['type', password]]);
  page.emit('framenavigated', 'main');
  assert.equal((await request('/api/action', { viewId, action: { type: 'type', text: 'do-not-type' } })).status, 409);
  assert.equal(actions.length, 1);
});

test('typing without a selected input is rejected instead of sending text to the page', async t => {
  const { request, page, actions } = await fixture(t);
  const { viewId } = await (await request('/api/state')).json();
  page.frames = () => [{ evaluate: async () => false }];
  assert.equal((await request('/api/action', { viewId, action: { type: 'type', text: 'do-not-type' } })).status, 409);
  assert.deepEqual(actions, []);
});

test('login portal rejects cross-origin actions, invalid hosts, arbitrary keys, and out-of-bounds clicks', async t => {
  const { request, url } = await fixture(t);
  const { viewId } = await (await request('/api/state')).json();
  const click = { viewId, action: { type: 'click', x: 10, y: 10 } };
  assert.equal((await request('/api/action', click, { Origin: 'https://evil.test' })).status, 403);
  const invalidHost = await new Promise((resolve, reject) => {
    const request = http.get(`${url.origin}/api/state`, { headers: { Host: 'evil.test' } }, response => {
      response.resume();
      resolve(response.statusCode);
    });
    request.once('error', reject);
  });
  assert.equal(invalidHost, 403);
  assert.equal((await request('/api/action', { viewId, action: { type: 'key', key: 'Control+Shift+I' } })).status, 400);
  assert.equal((await request('/api/action', { viewId, action: { type: 'click', x: 900, y: 10 } })).status, 409);
});

test('completed login revokes the link before returning Teams content or accepting another action', async t => {
  const { request, runtime, portal, actions } = await fixture(t);
  const { viewId } = await (await request('/api/state')).json();
  runtime.status = async () => ({ authenticated: true });
  const result = await (await request('/api/action', { viewId, action: { type: 'type', text: 'do-not-type' } })).json();
  assert.deepEqual(result, { state: 'complete', authenticated: true });
  assert.equal(portal.closed, true);
  assert.equal(portal.token, null);
  assert.deepEqual(actions, []);
});

test('expired and cancelled links lose access to the session', async t => {
  const expired = await fixture(t);
  expired.portal.expiresAt = Date.now() - 1;
  assert.equal((await expired.request('/api/state')).status, 401);
  const cancelled = await fixture(t);
  assert.equal((await cancelled.request('/api/cancel', {})).status, 200);
  assert.equal(cancelled.portal.closed, true);
  assert.equal(cancelled.portal.reason, 'cancelled');
});

test('remote login configuration requires HTTPS and an explicit local forwarding port', () => {
  assert.deepEqual(portalOptions({}), { port: 0, publicOrigin: undefined });
  assert.deepEqual(portalOptions({ TEAMS_MCP_LOGIN_PORT: '8787', TEAMS_MCP_LOGIN_PUBLIC_URL: 'https://login.example.test' }), { port: 8787, publicOrigin: 'https://login.example.test' });
  for (const env of [{ TEAMS_MCP_LOGIN_PORT: 'abc' }, { TEAMS_MCP_LOGIN_PORT: '-1' },
    { TEAMS_MCP_LOGIN_PUBLIC_URL: 'https://login.example.test' },
    { TEAMS_MCP_LOGIN_PORT: '8787', TEAMS_MCP_LOGIN_PUBLIC_URL: 'http://login.example.test' },
    { TEAMS_MCP_LOGIN_PORT: '8787', TEAMS_MCP_LOGIN_PUBLIC_URL: 'https://user:password@login.example.test' }]) {
    assert.throws(() => portalOptions(env), { code: 'login_configuration' });
  }
});

test('HTTP login requires explicit opt-in and a Tailscale IPv4 address', () => {
  const env = { TEAMS_MCP_LOGIN_PORT: '8787', TEAMS_MCP_LOGIN_ALLOW_TAILNET_HTTP: '1' };
  assert.deepEqual(portalOptions({ ...env, TEAMS_MCP_LOGIN_PUBLIC_URL: 'http://100.100.10.20:8443' }), { port: 8787, publicOrigin: 'http://100.100.10.20:8443' });
  for (const origin of ['http://example.test', 'http://192.168.1.1', 'http://100.63.0.1', 'http://100.128.0.1', 'http://100.100.10.20.evil.test']) {
    assert.throws(() => portalOptions({ ...env, TEAMS_MCP_LOGIN_PUBLIC_URL: origin }), { code: 'login_configuration' });
  }
  assert.throws(() => portalOptions({ TEAMS_MCP_LOGIN_PORT: '8787', TEAMS_MCP_LOGIN_PUBLIC_URL: 'http://100.100.10.20:8443' }), { code: 'login_configuration' });
});
