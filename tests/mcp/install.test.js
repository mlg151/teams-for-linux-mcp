'use strict';
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { execFileSync } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');

test('installer generates a valid user unit and preserves existing service settings', { skip: process.platform !== 'linux' || process.getuid() === 0 }, () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-install-'));
  const config = path.join(dir, 'config with spaces');
  const env = { ...process.env, XDG_CONFIG_HOME: config, XDG_RUNTIME_DIR: dir };
  const run = args => execFileSync(process.execPath, ['scripts/install-mcp-service.js', '--no-start', ...args], { cwd: path.resolve(__dirname, '../..'), env });
  try {
    run(['--data-dir', path.join(dir, 'profile'), '--login-port', '18788']);
    const settings = path.join(config, 'teams-mcp/service.env');
    const unit = path.join(config, 'systemd/user/teams-mcp.service');
    assert.equal(fs.statSync(settings).mode & 0o777, 0o600);
    const original = fs.readFileSync(settings, 'utf8');
    assert.match(original, /TEAMS_MCP_LOGIN_PORT=18788/);
    const definition = fs.readFileSync(unit, 'utf8');
    assert.match(definition, /Restart=always/);
    assert.match(definition, /KillMode=control-group/);
    assert.match(definition, /--nonblock --conflict-exit-code 73/);
    // Validates actual systemd syntax, including path quoting and ExecStart.
    if (fs.existsSync('/usr/bin/systemd-analyze')) execFileSync('systemd-analyze', ['--user', 'verify', unit], { env, stdio: 'pipe' });
    run(['--login-port', '18789']);
    assert.equal(fs.readFileSync(settings, 'utf8'), original);
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});
