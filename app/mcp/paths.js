'use strict';
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const runtimeDirectory = (env = process.env) => path.join(env.XDG_RUNTIME_DIR || `/run/user/${process.getuid()}`, 'teams-mcp');
const socketPath = (env = process.env) => env.TEAMS_MCP_SOCKET || path.join(runtimeDirectory(env), 'mcp.sock');
const profilePath = (env = process.env) => env.TEAMS_MCP_USER_DATA_DIR || path.join(env.XDG_DATA_HOME || path.join(os.homedir(), '.local/share'), 'teams-mcp/profile');
function secureDirectory(directory) {
  fs.mkdirSync(directory, { recursive: true, mode: 0o700 });
  const stat = fs.lstatSync(directory);
  if (!stat.isDirectory() || stat.isSymbolicLink() || stat.uid !== process.getuid() || (stat.mode & 0o077)) throw new Error(`Expected an owner-only directory: ${directory}`);
}
module.exports = { runtimeDirectory, socketPath, profilePath, secureDirectory };
