#!/usr/bin/env node
'use strict';
const { parseArgs } = require('node:util');
const { connectClient } = require('../app/mcp/localClient');
const { socketPath } = require('../app/mcp/paths');
(async () => {
  const { values, positionals } = parseArgs({ allowPositionals: true, options: { socket: { type: 'string' }, arguments: { type: 'string', default: '{}' } } });
  const client = await connectClient(values.socket || socketPath());
  try {
    const result = await client.call(positionals[0] || 'session_status', JSON.parse(values.arguments));
    console.log(JSON.stringify(result.structuredContent || result, null, 2));
    if (result.isError) process.exitCode = 1;
  } finally { client.close(); }
})().catch(error => { console.error(error.message); process.exitCode = 1; });
