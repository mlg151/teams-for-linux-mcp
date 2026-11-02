'use strict';
const fs = require('node:fs');
const path = require('node:path');
const net = require('node:net');
const { spawn, execFileSync } = require('node:child_process');
const { setTimeout: delay } = require('node:timers/promises');
const namespace = kind => fs.readlinkSync(`/proc/self/ns/${kind}`);
function assertPrivateNetwork(env = process.env) {
  if (process.platform !== 'linux' || env.TEAMS_MCP_PRIVATE_NETWORK !== '1' ||
      ['net', 'user', 'mnt'].some(kind => !env[`TEAMS_MCP_HOST_${kind.toUpperCase()}NS`] || namespace(kind) === env[`TEAMS_MCP_HOST_${kind.toUpperCase()}NS`])) {
    throw new Error('Refusing to launch Electron outside private user, network and mount namespaces. Start teams-mcp.service.');
  }
}
function unshareBinary() {
  return fs.existsSync('/usr/local/libexec/teams-mcp-unshare') ? '/usr/local/libexec/teams-mcp-unshare' : '/usr/bin/unshare';
}
async function namespaceWorkerSetup(directory) {
  assertPrivateNetwork();
  const mount = path.join(directory, 'mount');
  execFileSync(mount, ['--make-rprivate', '/']);
  execFileSync(mount, ['-t', 'tmpfs', '-o', 'mode=1777,size=512m', 'tmpfs', '/tmp']);
  execFileSync(mount, ['-t', 'tmpfs', '-o', 'mode=1777,size=256m', 'tmpfs', '/dev/shm']);
  execFileSync(mount, ['--bind', path.join(directory, 'resolv.conf'), '/etc/resolv.conf']);
  fs.writeSync(3, `${JSON.stringify({ pid: process.pid })}\n`);
  fs.closeSync(3);
  const deadline = Date.now() + 20000;
  while (!fs.existsSync(path.join(directory, 'network.ready'))) {
    if (Date.now() > deadline) throw new Error('Private network setup timed out.');
    await delay(50);
  }
}
function slirpRequest(socketPath, command) {
  return new Promise((resolve, reject) => {
    const socket = net.createConnection(socketPath);
    let buffer = '';
    socket.setTimeout(5000, () => socket.destroy(new Error('slirp4netns API timed out.')));
    socket.on('connect', () => socket.end(`${JSON.stringify(command)}\n`));
    socket.on('data', chunk => { buffer += chunk; });
    socket.on('error', reject);
    socket.on('end', () => {
      try { const response = JSON.parse(buffer); if (response.error) reject(new Error('slirp4netns could not forward the login port.')); else resolve(response); }
      catch (error) { reject(error); }
    });
  });
}
async function startPrivateNetwork(entry, args, directory, loginPort) {
  const env = { ...process.env, TEAMS_MCP_PRIVATE_NETWORK: '1', TEAMS_MCP_VIRTUAL_DISPLAY: '1' };
  for (const kind of ['net', 'user', 'mnt']) env[`TEAMS_MCP_HOST_${kind.toUpperCase()}NS`] = namespace(kind);
  delete env.NODE_OPTIONS;
  delete env.ELECTRON_RUN_AS_NODE;
  env.XDG_RUNTIME_DIR ||= `/run/user/${process.getuid()}`;
  const busSocket = path.join(directory, 'session-bus.sock');
  env.DBUS_SESSION_BUS_ADDRESS = `unix:path=${busSocket}`;
  // A non-setuid mount executable preserves the namespace's ambient caps and
  // avoids switching to an unmapped uid 0. No host capabilities are granted.
  fs.copyFileSync('/usr/bin/mount', path.join(directory, 'mount'));
  fs.chmodSync(path.join(directory, 'mount'), 0o700);
  for (const file of ['network.ready', 'slirp.sock', 'session-bus.sock']) { try { fs.unlinkSync(path.join(directory, file)); } catch { /* First launch. */ } }
  fs.writeFileSync(path.join(directory, 'resolv.conf'), 'nameserver 10.0.2.3\n', { mode: 0o600 });
  const child = spawn(unshareBinary(), ['--user', '--map-current-user', '--keep-caps', '--net', '--mount', '--fork', '--kill-child=KILL', process.execPath, entry, '--namespace-worker', ...args],
    { env, stdio: ['ignore', 'inherit', 'inherit', 'pipe'] });
  let slirp, busProxy;
  const exited = new Promise(resolve => child.once('exit', (code, signal) => resolve({ code, signal })));
  const stop = () => { child.kill('SIGTERM'); slirp?.kill('SIGTERM'); busProxy?.kill('SIGTERM'); };
  try {
    const pid = await new Promise((resolve, reject) => {
      let buffer = '';
      const timer = setTimeout(() => reject(new Error('Private namespace startup timed out.')), 10000);
      child.once('error', reject);
      child.once('exit', () => { clearTimeout(timer); reject(new Error('Namespace creation failed. On Ubuntu, install the scoped AppArmor helper; do not disable isolation.')); });
      child.stdio[3].on('data', chunk => { buffer += chunk; if (buffer.includes('\n')) { clearTimeout(timer); try { resolve(JSON.parse(buffer.split('\n')[0]).pid); } catch (error) { reject(error); } } });
    });
    slirp = spawn('slirp4netns', ['--configure', '--disable-host-loopback', '--ready-fd=3', `--api-socket=${path.join(directory, 'slirp.sock')}`,
      `--userns-path=/proc/${pid}/ns/user`, String(pid), 'tap0'], { stdio: ['ignore', 'inherit', 'inherit', 'pipe'] });
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('slirp4netns startup timed out.')), 10000);
      slirp.once('error', reject);
      slirp.once('exit', () => { clearTimeout(timer); reject(new Error('slirp4netns exited during startup.')); });
      slirp.stdio[3].once('data', () => { clearTimeout(timer); resolve(); });
    });
    await slirpRequest(path.join(directory, 'slirp.sock'), { execute: 'add_hostfwd', arguments: {
      proto: 'tcp', host_addr: '127.0.0.1', host_port: loginPort, guest_addr: '10.0.2.100', guest_port: loginPort,
    } });
    // Chromium otherwise moves its browser PID into a transient systemd scope,
    // outside this unit's KillMode=control-group. Keep the original keyring bus
    // reachable, but hide systemd and desktop integrations from the browser.
    const busAddress = process.env.DBUS_SESSION_BUS_ADDRESS || `unix:path=/run/user/${process.getuid()}/bus`;
    busProxy = spawn('xdg-dbus-proxy', ['--fd=3', busAddress, busSocket, '--filter',
      '--talk=org.freedesktop.secrets', '--talk=org.kde.kwalletd', '--talk=org.kde.kwalletd5', '--talk=org.kde.kwalletd6'],
    { stdio: ['ignore', 'inherit', 'inherit', 'pipe'] });
    await new Promise((resolve, reject) => {
      const timer = setTimeout(() => reject(new Error('D-Bus proxy startup timed out.')), 10000);
      busProxy.once('error', error => { clearTimeout(timer); reject(error); });
      busProxy.once('exit', () => { clearTimeout(timer); reject(new Error('D-Bus proxy exited during startup.')); });
      busProxy.stdio[3].once('data', () => { clearTimeout(timer); resolve(); });
    });
    fs.writeFileSync(path.join(directory, 'network.ready'), '', { mode: 0o600 });
    slirp.once('exit', () => child.kill('SIGTERM'));
    busProxy.once('exit', () => child.kill('SIGTERM'));
    return { child, stop, exited: exited.finally(() => { slirp?.kill('SIGTERM'); busProxy?.kill('SIGTERM'); }) };
  } catch (error) { stop(); throw error; }
}
module.exports = { assertPrivateNetwork, namespaceWorkerSetup, startPrivateNetwork, unshareBinary, slirpRequest };
