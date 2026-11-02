#!/usr/bin/env node
'use strict';
const { Console } = require('node:console');
const fs = require('node:fs');
const path = require('node:path');
const { parseArgs } = require('node:util');
const { spawn } = require('node:child_process');
const { runtimeDirectory, socketPath, profilePath, secureDirectory } = require('../app/mcp/paths');
Object.assign(console, new Console({ stdout: process.stderr, stderr: process.stderr }));
process.umask(0o077);
async function main() {
  const { values } = parseArgs({ options: { locked: { type: 'boolean' }, 'namespace-worker': { type: 'boolean' }, 'data-dir': { type: 'string' }, socket: { type: 'string' } } });
  const directory = runtimeDirectory();
  secureDirectory(directory);
  if (!values.locked && !values['namespace-worker']) {
    const lock = spawn('flock', ['--exclusive', '--nonblock', '--conflict-exit-code', '73', '--no-fork', path.join(directory, 'service.lock'), process.execPath, __filename, '--locked', ...process.argv.slice(2)], { stdio: 'inherit' });
    process.once('SIGTERM', () => lock.kill('SIGTERM'));
    process.once('SIGINT', () => lock.kill('SIGINT'));
    process.exitCode = await new Promise((resolve, reject) => { lock.once('error', reject); lock.once('exit', code => resolve(code ?? 1)); });
    return;
  }
  const loginPort = Number(process.env.TEAMS_MCP_LOGIN_PORT || 8788);
  if (!Number.isInteger(loginPort) || loginPort < 1 || loginPort > 65535) throw new Error('Set a fixed TEAMS_MCP_LOGIN_PORT between 1 and 65535.');
  process.env.TEAMS_MCP_LOGIN_PORT = String(loginPort);
  const profile = path.resolve(values['data-dir'] || profilePath());
  if (profile === '/tmp' || profile.startsWith('/tmp/')) throw new Error('Keep the persistent profile outside /tmp; the service has a private temporary filesystem.');
  secureDirectory(profile);
  if (!values['namespace-worker']) {
    const { startPrivateNetwork } = require('../app/mcp/privateNetwork');
    const args = ['--data-dir', profile, '--socket', values.socket || socketPath()];
    const isolated = await startPrivateNetwork(__filename, args, directory, loginPort);
    process.once('SIGTERM', isolated.stop);
    process.once('SIGINT', isolated.stop);
    const { code, signal } = await isolated.exited;
    process.exitCode = code || (signal ? 1 : 0);
    return;
  }
  const { namespaceWorkerSetup } = require('../app/mcp/privateNetwork');
  await namespaceWorkerSetup(directory);
  const xvfb = spawn('Xvfb', [':99', '-screen', '0', '1440x1000x24', '-nolisten', 'tcp'], { stdio: ['ignore', 'inherit', 'inherit'] });
  const { setTimeout: delay } = require('node:timers/promises');
  const deadline = Date.now() + 10000;
  while (!fs.existsSync('/tmp/.X11-unix/X99')) { if (Date.now() > deadline) throw new Error('Xvfb startup timed out.'); await delay(50); }
  process.env.DISPLAY = ':99';
  const { TeamsRuntime } = require('../app/mcp/runtime');
  const { TeamsDaemon } = require('../app/mcp/daemon');
  const runtime = new TeamsRuntime({ mode: 'headless', dataDir: values['data-dir'], virtualDisplay: true });
  const daemon = new TeamsDaemon(runtime, { socketPath: values.socket });
  let closing;
  const close = () => closing ||= (async () => { await daemon.close(); xvfb.kill('SIGTERM'); })();
  process.once('SIGTERM', close);
  process.once('SIGINT', close);
  await daemon.listen();
  console.error('Teams MCP service socket ready.');
  await runtime.start();
  runtime.application.once('close', () => { if (!closing) { process.exitCode = 1; close(); } });
  xvfb.once('exit', () => { if (!closing) { process.exitCode = 1; close(); } });
  console.error('Teams MCP isolated browser ready.');
}
main().catch(error => { console.error(error.message); process.exit(1); });
