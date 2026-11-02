'use strict';
const net = require('node:net');

// Operator/test client only; the session relay remains an opaque byte stream.
async function connectClient(socketPath) {
  const socket = net.createConnection(socketPath);
  await new Promise((resolve, reject) => { socket.once('connect', resolve); socket.once('error', reject); });
  return connectStream(socket, socket, () => socket.destroy());
}
async function connectStream(input, output, close) {
  let nextId = 0;
  let buffer = '';
  const pending = new Map();
  const fail = () => {
    for (const entry of pending.values()) { clearTimeout(entry.timer); entry.reject(new Error('Teams service disconnected; requests were not replayed.')); }
    pending.clear();
  };
  input.on('error', fail);
  output.on('error', fail);
  input.on('close', fail);
  input.on('data', chunk => {
    buffer += chunk;
    let newline;
    while ((newline = buffer.indexOf('\n')) !== -1) {
      const line = buffer.slice(0, newline);
      buffer = buffer.slice(newline + 1);
      let message;
      try { message = JSON.parse(line); } catch { fail(); close(); return; }
      const entry = pending.get(message.id);
      if (entry) {
        pending.delete(message.id);
        clearTimeout(entry.timer);
        if (message.error) entry.reject(new Error('The MCP request failed.'));
        else entry.resolve(message.result);
      }
    }
  });
  const send = value => output.write(`${JSON.stringify(value)}\n`);
  const request = (method, params) => new Promise((resolve, reject) => {
    const id = ++nextId;
    const timer = setTimeout(() => { pending.delete(id); reject(new Error('Teams request timed out; it was not retried.')); }, 120000);
    pending.set(id, { resolve, reject, timer });
    send({ jsonrpc: '2.0', id, method, params });
  });
  try {
    await request('initialize', { protocolVersion: '2025-11-25', capabilities: {}, clientInfo: { name: 'teams-local-client', version: '1' } });
    send({ jsonrpc: '2.0', method: 'notifications/initialized' });
  } catch (error) { close(); throw error; }
  return { request, call: (name, args = {}) => request('tools/call', { name, arguments: args }), close };
}
module.exports = { connectClient, connectStream };
