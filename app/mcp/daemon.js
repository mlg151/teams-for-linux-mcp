'use strict';
const fs = require('node:fs');
const net = require('node:net');
const path = require('node:path');
const { randomUUID } = require('node:crypto');
const { serveStdio, StdioServerTransport } = require('@modelcontextprotocol/server/stdio');
const { createServer } = require('./tools');
const { Microsoft365Service } = require('./service');
const { secureDirectory } = require('./paths');

// Each connection owns protocol state; every connection shares one runtime,
// credential cache, service and execution queue. No request is replayed.
class TeamsDaemon {
  constructor(runtime, { socketPath, service = new Microsoft365Service(runtime) }) {
    this.runtime = runtime;
    this.service = service;
    this.path = socketPath;
    this.connections = new Set();
    this.id = randomUUID();
    runtime.serviceId = this.id;
    this.closed = false;
  }
  async listen() {
    secureDirectory(path.dirname(this.path));
    if (Buffer.byteLength(this.path) > 103) throw new Error('The Unix socket path is too long.');
    this.server = net.createServer(socket => {
      const transport = new StdioServerTransport(socket, socket, { maxBufferSize: 1024 * 1024 });
      const handle = serveStdio(() => createServer(this.runtime, this.service), { transport, onerror: () => {} });
      this.connections.add(socket);
      socket.on('error', () => {});
      socket.once('end', () => socket.destroy());
      socket.once('close', () => { this.connections.delete(socket); handle.close().catch(() => {}); });
    });
    // The supervisor holds flock before namespace setup. Only an owned stale
    // endpoint can be removed; a live endpoint is never replaced.
    if (fs.existsSync(this.path)) {
      const stat = fs.lstatSync(this.path);
      if (!stat.isSocket() || stat.uid !== process.getuid()) throw new Error('Refusing to replace an unowned or non-socket endpoint.');
      if (await isListening(this.path)) throw new Error('The Teams service is already running.');
      fs.unlinkSync(this.path);
    }
    const previousMask = process.umask(0o077);
    try { await new Promise((resolve, reject) => { this.server.once('error', reject); this.server.listen(this.path, resolve); }); }
    finally { process.umask(previousMask); }
    fs.chmodSync(this.path, 0o600);
    this.inode = fs.lstatSync(this.path).ino;
    return this;
  }
  async close() {
    if (this.closed) return;
    this.closed = true;
    for (const socket of this.connections) socket.destroy();
    if (this.server?.listening) await new Promise(resolve => this.server.close(resolve));
    try { if (fs.lstatSync(this.path).ino === this.inode) fs.unlinkSync(this.path); } catch { /* Already removed. */ }
    await this.runtime.close();
  }
}
function isListening(socketPath) {
  return new Promise((resolve, reject) => {
    const socket = net.createConnection(socketPath);
    socket.once('connect', () => { socket.destroy(); resolve(true); });
    socket.once('error', error => { if (['ECONNREFUSED', 'ENOENT'].includes(error.code)) resolve(false); else reject(error); });
  });
}
module.exports = { TeamsDaemon };
