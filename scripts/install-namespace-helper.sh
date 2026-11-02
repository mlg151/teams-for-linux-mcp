#!/bin/sh
set -eu
if [ "$(id -u)" != 0 ]; then
  echo 'Run with sudo: sh scripts/install-namespace-helper.sh' >&2
  exit 1
fi
# No setuid bit or host capability is granted. This root-owned copy has a
# narrowly attached AppArmor userns policy, rather than changing a global sysctl
# or granting every Node process permission to create user namespaces.
install -D -m 0755 /usr/bin/unshare /usr/local/libexec/teams-mcp-unshare
if command -v apparmor_parser >/dev/null 2>&1 && [ -d /sys/kernel/security/apparmor ]; then
  cat > /etc/apparmor.d/teams-mcp-unshare <<'POLICY'
abi <abi/4.0>,
include <tunables/global>
profile teams-mcp-unshare /usr/local/libexec/teams-mcp-unshare flags=(unconfined) {
  userns,
}
POLICY
  chmod 0644 /etc/apparmor.d/teams-mcp-unshare
  apparmor_parser -r /etc/apparmor.d/teams-mcp-unshare
fi
echo 'Installed the Teams MCP rootless namespace launcher.'
