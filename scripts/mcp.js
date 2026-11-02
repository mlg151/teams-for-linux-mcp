#!/usr/bin/env node
'use strict';
// This relay imports no browser or SDK code, and never starts the service.
const fs = require('node:fs');
const net = require('node:net');
const { parseArgs } = require('node:util');
const { socketPath } = require('../app/mcp/paths');
function main() {
  const { values } = parseArgs({ options: { socket: { type: 'string' }, help: { type: 'boolean' } } });
  if (values.help) { console.error('Usage: node scripts/mcp.js [--socket /absolute/path/mcp.sock]\nConnects stdio to teams-mcp.service. Install/start the service first.'); return; }
  const endpoint = values.socket || socketPath();
  const stat = fs.lstatSync(endpoint);
  if (!stat.isSocket() || stat.uid !== process.getuid() || (stat.mode & 0o077)) throw new Error('The socket must belong to this user with mode 0600.');
  const socket = net.createConnection(endpoint);
  socket.on('connect', () => { process.stdin.pipe(socket); socket.pipe(process.stdout); });
  socket.on('error', () => { console.error('Teams service connection failed. Run systemctl --user restart teams-mcp.service. Requests are never retried.'); process.exitCode = 1; });
  socket.on('close', () => { process.stdin.unpipe(socket); process.stdin.pause(); });
  process.stdin.on('end', () => socket.end());
  process.stdin.on('error', () => socket.destroy());
  process.stdout.on('error', () => socket.destroy());
  const stop = () => { socket.destroy(); process.stdin.pause(); };
  process.once('SIGINT', stop);
  process.once('SIGTERM', stop);
}
try { main(); } catch { console.error('Cannot connect to the owner-only Teams socket. Start teams-mcp.service and check the socket path. This relay never starts Electron.'); process.exitCode = 1; }
