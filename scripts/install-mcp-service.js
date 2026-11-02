#!/usr/bin/env node
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { execFileSync } = require('node:child_process');
const { parseArgs } = require('node:util');
const { profilePath } = require('../app/mcp/paths');

const quote = value => `"${value.replaceAll('\\', '\\\\').replaceAll('"', '\\"').replaceAll('%', '%%').replaceAll('$', '$$')}"`;
const unitPath = value => {
  if (/[\r\n\0]/.test(value)) throw new Error('Paths must not contain control characters.');
  return value.replaceAll('%', '%%');
};
function main() {
  const { values } = parseArgs({ options: { 'data-dir': { type: 'string' }, 'login-port': { type: 'string', default: '8788' }, 'no-start': { type: 'boolean' } } });
  if (process.platform !== 'linux' || process.getuid() === 0) throw new Error('Install as the owning non-root Linux user.');
  const port = Number(values['login-port']);
  if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error('Invalid login port.');
  const root = path.resolve(__dirname, '..');
  const configRoot = process.env.XDG_CONFIG_HOME || path.join(os.homedir(), '.config');
  const directory = path.join(configRoot, 'teams-mcp');
  const units = path.join(configRoot, 'systemd/user');
  fs.mkdirSync(directory, { recursive: true, mode: 0o700 });
  fs.mkdirSync(units, { recursive: true });
  const environment = path.join(directory, 'service.env');
  if (!fs.existsSync(environment)) {
    const profile = path.resolve(values['data-dir'] || profilePath());
    fs.writeFileSync(environment, `TEAMS_MCP_USER_DATA_DIR=${JSON.stringify(profile)}\nTEAMS_MCP_LOGIN_PORT=${port}\n`, { mode: 0o600 });
  }
  const unit = `[Unit]
Description=Private per-user Teams MCP service
After=network-online.target
StartLimitIntervalSec=60
StartLimitBurst=5

[Service]
Type=simple
RuntimeDirectory=teams-mcp
RuntimeDirectoryMode=0700
RuntimeDirectoryPreserve=yes
UMask=0077
EnvironmentFile=${unitPath(environment)}
UnsetEnvironment=NODE_OPTIONS ELECTRON_RUN_AS_NODE
WorkingDirectory=${unitPath(root)}
ExecStart=/usr/bin/flock --exclusive --nonblock --conflict-exit-code 73 --no-fork %t/teams-mcp/service.lock ${quote(process.execPath)} ${quote(path.join(root, 'scripts/mcp-service.js'))} --locked
Restart=always
RestartSec=3
RestartPreventExitStatus=73
KillMode=control-group
TimeoutStopSec=20

[Install]
WantedBy=default.target
`;
  fs.writeFileSync(path.join(units, 'teams-mcp.service'), unit, { mode: 0o644 });
  console.log(`Installed teams-mcp.service; configuration: ${environment}`);
  if (!values['no-start']) {
    const env = { ...process.env, XDG_RUNTIME_DIR: process.env.XDG_RUNTIME_DIR || `/run/user/${process.getuid()}` };
    env.DBUS_SESSION_BUS_ADDRESS ||= `unix:path=${env.XDG_RUNTIME_DIR}/bus`;
    execFileSync('systemctl', ['--user', 'daemon-reload'], { env, stdio: 'inherit' });
    execFileSync('systemctl', ['--user', 'enable', '--now', 'teams-mcp.service'], { env, stdio: 'inherit' });
  }
}
try { main(); } catch (error) { console.error(error.message); process.exitCode = 1; }
