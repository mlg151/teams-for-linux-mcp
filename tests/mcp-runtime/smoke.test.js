'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { spawn } = require('node:child_process');
const fs = require('node:fs');
const net = require('node:net');
const path = require('node:path');
const { setTimeout: delay } = require('node:timers/promises');
const { connectStream } = require('../../app/mcp/localClient');

async function unusedPort() {
  const server = net.createServer();
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const port = server.address().port;
  await new Promise(resolve => server.close(resolve));
  return port;
}
async function reachable(port) {
  return new Promise(resolve => {
    const socket = net.createConnection({ host: '127.0.0.1', port });
    socket.setTimeout(1000, () => { socket.destroy(); resolve(false); });
    socket.once('connect', () => { socket.destroy(); resolve(true); });
    socket.once('error', () => resolve(false));
  });
}

test('isolated real service shares one browser across three relays and only forwards bearer-protected login', { timeout: 120000 }, async t => {
  const base = process.env.XDG_RUNTIME_DIR || `/run/user/${process.getuid()}`;
  const dir = fs.mkdtempSync(path.join(base, 'teams-mcp-test-'));
  const endpoint = path.join(dir, 'teams-mcp/mcp.sock');
  const profile = path.join(dir, 'profile');
  const loginPort = await unusedPort();
  const root = path.resolve(__dirname, '../..');
  const child = spawn(process.execPath, ['scripts/mcp-service.js', '--data-dir', profile], {
    cwd: root, detached: true,
    env: { ...process.env, XDG_RUNTIME_DIR: dir, TEAMS_MCP_LOGIN_PORT: String(loginPort), DISPLAY: '', DEBUG: '' },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  let logs = '';
  let exited = false;
  const exit = new Promise(resolve => child.once('exit', code => { exited = true; resolve(code); }));
  child.stdout.on('data', chunk => { logs += chunk; });
  child.stderr.on('data', chunk => { logs += chunk; });
  const relays = [];
  const clients = [];
  t.after(async () => {
    for (const client of clients) client.close();
    for (const relay of relays) relay.kill();
    try { process.kill(-child.pid, 'SIGTERM'); } catch { /* Already stopped. */ }
    const timer = setTimeout(() => { try { process.kill(-child.pid, 'SIGKILL'); } catch { /* Already stopped. */ } }, 8000);
    await exit;
    clearTimeout(timer);
    try { process.kill(-child.pid, 'SIGKILL'); } catch { /* Reap descendants of the manual test supervisor. */ }
    fs.rmSync(dir, { recursive: true, force: true });
  });
  const deadline = Date.now() + 75000;
  while (!logs.includes('isolated browser ready')) {
    if (exited || Date.now() > deadline) throw new Error(`Service failed to start: ${logs}`);
    await delay(100);
  }
  for (let i = 0; i < 3; i++) {
    const relay = spawn(process.execPath, ['scripts/mcp.js', '--socket', endpoint], { cwd: root, stdio: ['pipe', 'pipe', 'pipe'] });
    relay.exitPromise = new Promise(resolve => relay.once('exit', resolve));
    relay.stderr.resume();
    relays.push(relay);
    clients.push(await connectStream(relay.stdout, relay.stdin, () => relay.stdin.end()));
  }
  const states = await Promise.all(clients.map(client => client.call('session_status')));
  for (const state of states) {
    assert.equal(state.isError, undefined);
    assert.equal(state.structuredContent.authenticated, false);
    assert.equal(state.structuredContent.serviceId, states[0].structuredContent.serviceId);
  }
  const pid = states[0].structuredContent.servicePid;
  const workerCgroup = fs.readFileSync(`/proc/${pid}/cgroup`, 'utf8');
  const childPids = fs.readFileSync(`/proc/${pid}/task/${pid}/children`, 'utf8').trim().split(/\s+/).map(Number);
  const electronPid = childPids.find(childPid => fs.readFileSync(`/proc/${childPid}/comm`, 'utf8').trim() === 'electron');
  assert.ok(electronPid);
  assert.equal(fs.readFileSync(`/proc/${electronPid}/cgroup`, 'utf8'), workerCgroup, 'Electron must remain in its supervisor cgroup.');
  for (const kind of ['net', 'user', 'mnt']) {
    assert.notEqual(fs.readlinkSync(`/proc/${pid}/ns/${kind}`), fs.readlinkSync(`/proc/self/ns/${kind}`));
  }
  const privatePorts = fs.readFileSync(`/proc/${pid}/net/tcp`, 'utf8').split('\n').slice(1)
    .map(line => line.trim().split(/\s+/)).filter(fields => fields[3] === '0A' && fields[1]?.startsWith('0100007F:'))
    .map(fields => parseInt(fields[1].split(':')[1], 16));
  assert.ok(privatePorts.length >= 2, 'Both Electron debugging listeners exist in the private network.');
  for (const port of privatePorts) assert.equal(await reachable(port), false, `Debugging port ${port} must be unreachable from the host.`);
  const login = await clients[1].call('login');
  assert.equal(login.isError, undefined, JSON.stringify(login));
  const url = new URL(login.structuredContent.url);
  assert.equal(url.port, String(loginPort));
  assert.equal((await fetch(`${url.origin}/api/state`)).status, 401);
  const screen = await fetch(`${url.origin}/api/state`, { headers: { Authorization: `Bearer ${url.hash.slice(1)}`, Host: 'localhost:52123' } });
  assert.equal(screen.status, 200);
  assert.match((await screen.json()).image, /^iVBOR/);
  clients[0].close();
  assert.equal(await relays[0].exitPromise, 0);
  assert.equal(exited, false);
  assert.equal((await clients[2].call('session_status')).structuredContent.serviceId, states[0].structuredContent.serviceId);
  const duplicate = spawn(process.execPath, ['scripts/mcp-service.js', '--data-dir', profile], { cwd: root, env: { ...process.env, XDG_RUNTIME_DIR: dir }, stdio: 'ignore' });
  assert.equal(await new Promise(resolve => duplicate.once('exit', resolve)), 73);
});
