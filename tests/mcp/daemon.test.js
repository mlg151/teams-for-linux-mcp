'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawn } = require('node:child_process');
const { setTimeout: delay } = require('node:timers/promises');
const { TeamsDaemon } = require('../../app/mcp/daemon');
const { TeamsRuntime } = require('../../app/mcp/runtime');
const { connectClient, connectStream } = require('../../app/mcp/localClient');
const { assertPrivateNetwork } = require('../../app/mcp/privateNetwork');

test('three relays share one signed-in runtime, serialize sends and survive client EOF', { timeout: 15000 }, async t => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-daemon-'));
  const endpoint = path.join(dir, 'mcp.sock');
  const runtime = new TeamsRuntime();
  let closes = 0;
  const close = runtime.close.bind(runtime);
  runtime.close = async () => { closes++; await close(); };
  runtime.status = async () => ({ authenticated: true, state: 'ready', serviceId: runtime.serviceId });
  const writes = [];
  let active = 0;
  let maximum = 0;
  const service = { sendMessage: async args => {
    maximum = Math.max(maximum, ++active);
    await delay(30);
    writes.push(args);
    active--;
    return { id: `message-${writes.length}`, ...args };
  } };
  const daemon = await new TeamsDaemon(runtime, { socketPath: endpoint, service }).listen();
  const clients = [];
  const children = [];
  t.after(async () => {
    for (const client of clients) client.close();
    for (const child of children) child.kill();
    await daemon.close();
    fs.rmSync(dir, { recursive: true, force: true });
  });
  assert.equal(fs.statSync(endpoint).mode & 0o777, 0o600);
  assert.equal(fs.statSync(dir).mode & 0o777, 0o700);
  for (let i = 0; i < 3; i++) {
    const child = spawn(process.execPath, ['scripts/mcp.js', '--socket', endpoint], { cwd: path.resolve(__dirname, '../..'), stdio: ['pipe', 'pipe', 'pipe'] });
    child.exitPromise = new Promise(resolve => child.once('exit', resolve));
    children.push(child);
    clients.push(await connectStream(child.stdout, child.stdin, () => child.stdin.end()));
  }
  const states = await Promise.all(clients.map(client => client.call('session_status')));
  for (const state of states) assert.deepEqual(state.structuredContent, { authenticated: true, state: 'ready', serviceId: daemon.id });
  const sent = await Promise.all([
    clients[0].call('send_message', { chatId: 'self-chat', content: 'one' }),
    clients[1].call('send_message', { chatId: 'self-chat', content: 'two' }),
  ]);
  assert.ok(sent.every(result => !result.isError));
  assert.equal(maximum, 1);
  assert.deepEqual(writes.map(write => write.content).sort(), ['one', 'two']);
  clients[0].close();
  assert.equal(await children[0].exitPromise, 0);
  assert.equal(closes, 0);
  assert.equal((await clients[2].call('session_status')).structuredContent.serviceId, daemon.id);
  const replacement = await connectClient(endpoint);
  clients.push(replacement);
  assert.equal((await replacement.call('session_status')).structuredContent.authenticated, true);
  // A second listener must never replace the live endpoint.
  await assert.rejects(new TeamsDaemon(new TeamsRuntime(), { socketPath: endpoint, service }).listen(), /already running/);
});

test('relay refuses an endpoint with group/world access and never launches a browser', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-relay-'));
  const endpoint = path.join(dir, 'mcp.sock');
  const daemon = await new TeamsDaemon(new TeamsRuntime(), { socketPath: endpoint, service: {} }).listen();
  try {
    fs.chmodSync(endpoint, 0o666);
    const child = spawn(process.execPath, ['scripts/mcp.js', '--socket', endpoint], { cwd: path.resolve(__dirname, '../..') });
    let stdout = '';
    child.stdout.on('data', chunk => { stdout += chunk; });
    child.stderr.resume();
    assert.equal(await new Promise(resolve => child.once('exit', resolve)), 1);
    assert.equal(stdout, '');
    assert.equal(daemon.connections.size, 0);
  } finally { await daemon.close(); fs.rmSync(dir, { recursive: true, force: true }); }
});

test('a marker cannot bypass namespace isolation on the host network', () => {
  const env = { TEAMS_MCP_PRIVATE_NETWORK: '1' };
  for (const kind of ['net', 'user', 'mnt']) env[`TEAMS_MCP_HOST_${kind.toUpperCase()}NS`] = fs.readlinkSync(`/proc/self/ns/${kind}`);
  assert.throws(() => assertPrivateNetwork(env), /Refusing to launch/);
});
