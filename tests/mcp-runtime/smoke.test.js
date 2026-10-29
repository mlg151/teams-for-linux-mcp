'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const { spawn } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');

test('real stdio server starts Electron under Xvfb and exits when the client disconnects', { timeout: 90000 }, async t => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-mcp-smoke-'));
  const child = spawn(process.execPath, ['scripts/mcp.js', '--data-dir', dir], {
    cwd: path.resolve(__dirname, '../..'),
    env: { ...process.env, DISPLAY: '', DEBUG: '' }, stdio: ['pipe', 'pipe', 'pipe'],
  });
  let buffer = '';
  let stderr = '';
  let requestId = 0;
  const pending = new Map();
  const invalidLines = [];
  const exit = new Promise((resolve, reject) => {
    child.once('error', reject);
    child.once('exit', code => resolve(code));
  });
  child.stderr.on('data', data => { stderr += data.toString(); });
  child.stdout.on('data', data => {
    buffer += data.toString();
    let newline;
    while ((newline = buffer.indexOf('\n')) >= 0) {
      const line = buffer.slice(0, newline);
      buffer = buffer.slice(newline + 1);
      let message;
      try { message = JSON.parse(line); } catch { invalidLines.push(line); continue; }
      if (pending.has(message.id)) { pending.get(message.id)(message); pending.delete(message.id); }
    }
  });
  t.after(async () => {
    child.stdin.end();
    const timer = setTimeout(() => child.kill('SIGTERM'), 10000);
    try { await exit; } finally { clearTimeout(timer); fs.rmSync(dir, { recursive: true, force: true }); }
  });
  const send = message => child.stdin.write(`${JSON.stringify(message)}\n`);
  async function request(method, params) {
    const id = ++requestId;
    const response = new Promise(resolve => pending.set(id, resolve));
    send({ jsonrpc: '2.0', id, method, ...(params ? { params } : {}) });
    return Promise.race([response, exit.then(code => { throw new Error(`MCP exited early (${code}): ${stderr}`); })]);
  }
  const initialized = await request('initialize', { protocolVersion: '2025-11-25', capabilities: {}, clientInfo: { name: 'smoke', version: '1' } });
  assert.equal(initialized.result.serverInfo.name, 'teams-for-linux-mcp');
  send({ jsonrpc: '2.0', method: 'notifications/initialized' });
  const listed = await request('tools/list');
  assert.equal(listed.result.tools.length, 15);
  const status = await request('tools/call', { name: 'session_status', arguments: {} });
  assert.equal(status.result.isError, undefined, JSON.stringify(status.result));
  assert.equal(status.result.structuredContent.authenticated, false);
  assert.equal(status.result.structuredContent.mode, 'headless');
  const login = await request('tools/call', { name: 'login', arguments: {} });
  assert.equal(login.result.isError, undefined, JSON.stringify(login.result));
  const loginUrl = new URL(login.result.structuredContent.url);
  assert.equal(loginUrl.hostname, '127.0.0.1');
  const screen = await fetch(`${loginUrl.origin}/api/state`, { headers: { Authorization: `Bearer ${loginUrl.hash.slice(1)}` } });
  assert.equal(screen.status, 200);
  const loginState = await screen.json();
  assert.equal(loginState.state, 'pending');
  assert.match(loginState.image, /^iVBOR/);
  child.stdin.end();
  assert.equal(await exit, 0, stderr);
  assert.deepEqual(invalidLines, []);
});
