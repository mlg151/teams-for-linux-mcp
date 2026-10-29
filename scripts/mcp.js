#!/usr/bin/env node
'use strict';

const { Console } = require('node:console');
const { spawn } = require('node:child_process');
const { parseArgs } = require('node:util');
const path = require('node:path');

// Only the SDK writes to stdout. Electron runs in a separate process and its
// application logs are captured by Playwright rather than the MCP pipe.
Object.assign(console, new Console({ stdout: process.stderr, stderr: process.stderr }));

async function main() {
  const { values } = parseArgs({ options: {
    login: { type: 'boolean', default: false },
    'data-dir': { type: 'string' },
    help: { type: 'boolean', default: false },
  } });
  if (values.help) {
    console.error('Usage: node scripts/mcp.js [--login] [--data-dir /absolute/profile/path]\nWithout --login, serves MCP over stdio and uses private xvfb-run on Linux. Call the MCP login tool for a browser sign-in link. --login opens the optional desktop app.');
    return;
  }
  if (process.platform === 'linux' && !values.login && process.env.TEAMS_MCP_VIRTUAL_DISPLAY !== '1') {
    const child = spawn('xvfb-run', ['-a', '-s', '-screen 0 1440x1000x24 -nolisten tcp',
      process.execPath, path.resolve(__filename), ...process.argv.slice(2)], {
      stdio: 'inherit', detached: true, env: { ...process.env, TEAMS_MCP_VIRTUAL_DISPLAY: '1' },
    });
    const stop = () => { try { process.kill(-child.pid, 'SIGTERM'); } catch { /* Already exited. */ } };
    process.once('SIGTERM', stop);
    process.once('SIGINT', stop);
    await new Promise((resolve, reject) => {
      child.once('error', error => reject(new Error(`Could not start xvfb-run: ${error.code}. Install xvfb and xauth.`)));
      child.once('exit', code => { process.exitCode = code ?? 1; resolve(); });
    });
    return;
  }
  if (process.platform === 'linux' && values.login && !process.env.DISPLAY) {
    throw new Error('--login opens the optional desktop UI. For headless sign-in, start the MCP server without --login and call its login tool to get a private browser link.');
  }
  const { TeamsRuntime } = require('../app/mcp/runtime');
  const runtime = new TeamsRuntime({ mode: values.login ? 'login' : 'headless',
    dataDir: values['data-dir'] || process.env.TEAMS_MCP_USER_DATA_DIR });
  let handle;
  let closing;
  const close = () => {
    if (!closing) closing = (async () => { await runtime.close(); await handle?.close(); })();
    return closing;
  };
  process.once('SIGINT', close);
  process.once('SIGTERM', close);
  if (values.login) {
    const application = await runtime.start();
    console.error('Complete Microsoft sign-in in Teams, then close the app. The MCP server will reuse the same session.');
    await new Promise(resolve => application.once('close', resolve));
    await close();
  } else {
    const { serveStdio } = require('@modelcontextprotocol/server/stdio');
    const { createServer } = require('../app/mcp/tools');
    handle = serveStdio(() => createServer(runtime), { onerror: () => console.error('MCP transport error.') });
    process.stdin.once('end', close);
    process.stdin.once('error', close);
    process.stdin.once('close', close);
    process.stdout.once('error', close);
  }
}

main().catch(error => {
  console.error(error.message);
  process.exitCode = 1;
});
