# Shared headless Microsoft 365 MCP service

One systemd **user** service owns one Teams Electron instance, one persistent profile, one sign-in, and one execution queue. Every Claude session connects through a small stdio relay. Disconnecting a client does not close Teams. No call-control tools are exposed.

```text
Claude session A ─┐                       private user/network/mount namespaces
Claude session B ─┼─ stdio relays ─ Unix socket ─ shared MCP runtime ─ Electron
Claude session C ─┘                 mode 0600       one queue          + Xvfb
```

This is experimental. Automated checks exercise simultaneous clients, queueing, client disconnects, private networking, the real signed-out Electron app, browser login controls, and Graph request construction. Tenant-specific permissions and a real user's sign-in still need checking on that user's account.

## Install on Linux

Use Node 24 (`.nvmrc`) and an ordinary, non-root OS account. On Debian/Ubuntu, install the namespace/network/display tools plus the normal Electron libraries listed in the repository installation documentation:

```bash
sudo apt-get install slirp4netns xdg-dbus-proxy util-linux xvfb xauth
npm ci
```

The Electron download must exist. If npm install scripts were disabled, run `node node_modules/electron/install.js` explicitly.

The service requires unprivileged user namespaces. Ubuntu's AppArmor policy may block them. In that case install this repository's scoped launcher once on the host:

```bash
sudo sh scripts/install-namespace-helper.sh
```

This installs a root-owned, **non-setuid** copy of `unshare` at `/usr/local/libexec/teams-mcp-unshare` with an AppArmor rule permitting user namespaces for that executable. It does not disable the host-wide user-namespace restriction or grant host capabilities. Hosts which already allow user namespaces can use `/usr/bin/unshare` directly. Missing isolation or `slirp4netns` makes startup fail; there is no fallback to exposed debugging ports.

Install and start the service **as the account that will own the Teams sign-in**:

```bash
npm run mcp:install -- --data-dir "$HOME/.local/share/teams-mcp/profile" --login-port 8788
systemctl --user status teams-mcp.service
```

The installer writes `~/.config/systemd/user/teams-mcp.service` and a mode-0600 `~/.config/teams-mcp/service.env`, then enables and starts the unit. It uses absolute paths to this checkout and Node. Re-run the installer after moving either one. Existing `service.env` settings are preserved; edit that file to change an existing installation:

```ini
TEAMS_MCP_USER_DATA_DIR="/home/you/.local/share/teams-mcp/profile"
TEAMS_MCP_LOGIN_PORT=8788
```

Keep the profile outside `/tmp`, owned by this user, with mode 0700. The service uses private temporary filesystems for `/tmp` and `/dev/shm`. A desktop Teams instance must use another profile. Every OS account needs its own profile, socket, and a distinct host login port if multiple users run this service. The port is reserved by slirp while the service runs, including between login links.

To keep the user manager running after the last login ends and start it at boot:

```bash
sudo loginctl enable-linger "$USER"
```

An administrator may need to enable lingering under the host's policy. Check with `loginctl show-user "$USER" -p Linger`. The unit uses `Restart=always`, restarts after three seconds, and stops the entire service cgroup before a replacement starts. Repeated startup failures hit a rate limit. An exclusive `flock` prevents a second service from starting against the same runtime directory.

If a remote terminal has no user-bus environment, set `XDG_RUNTIME_DIR=/run/user/$(id -u)` and `DBUS_SESSION_BUS_ADDRESS=unix:path=$XDG_RUNTIME_DIR/bus` before using `systemctl --user`.

## Configure every Claude session

In **`~/.claude.json`**, replace only the `mcpServers.teams` entry, preserving other settings:

```json
{
  "mcpServers": {
    "teams": {
      "type": "stdio",
      "command": "/usr/bin/node",
      "args": ["/home/you/teams-for-linux-mcp/scripts/mcp.js"]
    }
  }
}
```

Use your actual absolute Node and checkout paths; `~` does not expand in JSON. Remove the old `--data-dir` argument and per-client `TEAMS_MCP_LOGIN_PORT` setting. Profile and login settings now belong to the service. The default socket is `$XDG_RUNTIME_DIR/teams-mcp/mcp.sock`, falling back to `/run/user/UID/teams-mcp/mcp.sock`. An explicit `--socket /absolute/path/mcp.sock` relay argument is available when a client has a different runtime environment.

Close/restart old Claude sessions after changing their configuration. Before starting the service against an existing profile, stop their old MCP launchers and their Electron/Xvfb descendants. Inspect `ps -u "$USER" -o pid,ppid,args` to identify processes for this checkout; do not stop unrelated Node or Electron applications. Old sessions retain the old configuration until restarted.

A relay only forwards MCP bytes. It imports no browser code, never starts the service, and never reconnects or replays a request. If the service is unavailable, start it and reconnect the MCP client. Use the direct Node command in Claude, since npm prints a non-protocol banner to stdout.

## Sign in through VS Code Remote Tunnels

1. In the VS Code window connected to the Linux host, open the **Ports** tab. Choose **Forward a Port**, enter remote port **8788**, and keep the forward private. Do this **before** requesting a login link.
2. From Claude, call **`login`** with no arguments. Open the returned private link on your Mac. It looks like `http://127.0.0.1:8788/#...`.
3. If VS Code selected another local port, replace the link's host/port with the **Forwarded Address** from the Ports tab, retaining the entire `/#...` fragment. Both `localhost` and loopback IP addresses on any local port are accepted. Do not make the port public.
4. The page displays the actual headless sign-in window. Click a field in the image, type using the private input above it, and complete Microsoft sign-in and MFA. Passwords and codes travel outside the MCP conversation.
5. The page confirms completion and revokes the link. Call **`session_status`**. All Claude sessions now use the same signed-in service. Keep the service/profile; ending Claude does not sign you out.

The link expires after ten minutes. If the tunnel was missing or the link expired, forward the port and call `login` again. You do not need to restart Teams. Closing a login link does not clear the Microsoft session. Do not share the fragment: it is a temporary bearer credential for the sign-in screen. Hardware keys/passkeys depend on devices available to the remote browser; use another permitted sign-in method when appropriate.

For SSH instead of Remote Tunnels, forward `ssh -N -L 8788:127.0.0.1:8788 your-server` before opening the link. An existing HTTPS proxy can use `TEAMS_MCP_LOGIN_PUBLIC_URL=https://login.example.test` in `service.env`; preserve the browser Origin header and public Host (or loopback upstream Host). The service itself still forwards only a loopback host listener.

Operator commands use the same socket and tools:

```bash
npm run mcp:status
npm run mcp:login
node scripts/mcp-call.js search_chats --arguments '{"query":"Your name"}'
```

After login, choose your own chat deliberately from `search_chats` results and verify its participants before calling `send_message` with its explicit `chatId`. Send one test message, then use `read_messages` on that chat to check it. For a multi-client send test use two distinct message texts, and verify each appears once. The test suite uses synthetic sends; it never sends real Microsoft messages.

## Start, restart and troubleshoot

```bash
systemctl --user start teams-mcp.service
systemctl --user restart teams-mcp.service
systemctl --user stop teams-mcp.service
journalctl --user -u teams-mcp.service -n 60 --no-pager
```

After editing `service.env`, restart the unit. After editing the unit, run `systemctl --user daemon-reload` first. After resolving a repeated startup failure, use `systemctl --user reset-failed teams-mcp.service` and start it again. To remove autostart, use `systemctl --user disable --now teams-mcp.service`; this preserves the profile. A service restart disconnects existing clients; reconnect their MCP transport. In-flight writes are never replayed.

Use `session_status.serviceId` and `.servicePid` to confirm clients share the same service incarnation. These change on a full service restart. A persistent profile reuses Microsoft's saved session when still valid; Microsoft may require sign-in again. Keep the same OS account/keyring environment. Profile migration by copying encrypted cookies is not supported.

## Security on a shared host

Electron's inspector and remote-debugging ports bind **inside private network namespaces**, along with the service and Xvfb. They have no host forwarding. The namespace also has private `/tmp` and `/dev/shm`, protecting X11 sockets. `slirp4netns --disable-host-loopback` supplies outbound networking and forwards **only** the chosen login port to host `127.0.0.1`. The login API requires a random, expiring bearer credential; unauthenticated access reveals no sign-in screen or input.

`xdg-dbus-proxy` retains access to the owning user's existing Secret Service/KWallet keyring while hiding the systemd bus service. This prevents Chromium's desktop-portal integration from moving Electron outside the service cgroup, so a supervisor crash cannot leave it holding the profile lock. It does not switch cookie encryption to plaintext or replace the user's keyring. Re-run the scoped helper installer after updating `util-linux` to refresh its copy of `unshare`.

Clients use a filesystem Unix socket, mode **0600**, in a mode-0700 user-owned runtime directory. The relay rejects unowned or group/world-accessible sockets. Other unprivileged OS users cannot connect to the MCP socket or reach the private debugging listeners through host TCP. Root and other processes running as the owning OS account remain trusted. This protects cross-user access; it is not a sandbox against the owner or administrator.

## Tools

| Tool | Purpose |
| --- | --- |
| `session_status` | Start the runtime and check login readiness |
| `login` | Return a private browser link for headless Microsoft sign-in and MFA |
| `search_people` | Find people by name or email and return their exact user IDs and email addresses |
| `resolve_chat` | Resolve a verified one-to-one chat for an exact person without sending a message |
| `search_messages` | Search Teams chat/channel message content |
| `search_chats` | Search chat names and available participant names/emails |
| `search_teams` | Search joined Teams by name/description |
| `search_onedrive` | Search the user's OneDrive files |
| `list_onedrive_files` | List a OneDrive folder's metadata |
| `list_channels` | List channels in a known Team |
| `read_messages` | Read messages from a known chat or Team/channel |
| `send_message` | Send plain text to a known chat or Team/channel; reply to a channel thread |
| `list_calendar_events` | List events and recurring occurrences in a date range |
| `create_calendar_event` | Create an event, optionally with attendees and a Teams meeting link |
| `update_calendar_event` | Update an existing event |

Message tools accept either `chatId` or both `teamId` and `channelId`. Sending requires an explicit target; a search term is not a destination. `replyToId` is supported for channel threads. IDs can come from Graph results or Teams links. Browser result items include IDs only when the rendered DOM or links expose them.

To message a person by name, call `search_people`, select the intended account from its email addresses, call `resolve_chat` with that exact `userId`, then call `send_message` with the returned `chatId`. Chat resolution uses the app's existing conversation helper and verifies that the one-to-one chat has exactly two members including the selected user. Resolving a recipient does not itself send a message.

Search tools accept `backend: "auto" | "graph" | "browser"`. `auto` attempts Graph and falls back to the signed-in UI only for permission errors or an unavailable Graph token. Network errors, throttling, expired login, and invalid input do not cause fallback.

`search_messages` supports `offset` for Graph results. `search_onedrive` and list/read tools return an opaque `nextCursor` when more Graph results exist; pass it with the same query, target, and limit. Cursors are held in memory and expire on a full Teams navigation or server restart. Chat/Team name searches scan at most ten API pages and report `truncated` when appropriate.

## How session access works

`SessionGraph` reuses this repository's `GraphApiClient.acquireToken()` through the signed-in Teams renderer. Graph requests remain on `https://graph.microsoft.com/v1.0/`, validate continuation URLs, and reject redirects. Raw session tokens never appear in MCP results. All clients share the same token client/cache as well as the runtime.

Using the Teams session does not request a separate Entra application consent. It retains the session token's existing permissions: it does not bypass tenant policies or grant scopes that account lacks. Chat-history reads, channels, OneDrive, and calendar operations can return explicit permission errors. Auto search falls back to signed-in browser controls only for permission/token availability failures.

Browser fallback targets English UI labels and known Teams/OneDrive layouts and request patterns. It is fixture-tested and needs tenant validation. Results identify their source and cover **visible results only**; they do not provide exhaustive history or Graph pagination. A UI mismatch returns an error instead of returning stale results.

All clients' tools serialize through one shared queue. A cancelled queued operation is discarded before dispatch. Calendar writes and sends are never automatically retried. If a request times out or loses its connection after dispatch, delivery may be uncertain: inspect Teams/calendar before attempting it again. Queueing prevents concurrent UI interference and duplicate automatic dispatch; it cannot guarantee exactly-once delivery across a crash or a manual retry after an unknown outcome.

## Validation

```bash
npm run test:mcp
npm run test:mcp:browser
npm run test:mcp:runtime
npm run lint
```

Unit/protocol tests connect three real relay processes to a signed-in fixture runtime, exercise concurrent sends through one queue, reject insecure socket permissions, and check that client EOF leaves the service available. Browser tests use synthetic sign-in/search pages, including password/MFA controls; they never use an account.

The Linux runtime test requires the installation prerequisites and working namespace support. It launches real Electron under one private Xvfb, connects three relays, compares their shared signed-out status, verifies namespace separation and host-inaccessible debug listeners, exercises the bearer-protected login screen through slirp, and rejects a duplicate service before browser startup. The profile lives in a temporary user runtime directory outside `/tmp` and is removed afterward.

For live acceptance, sign in once as the intended OS user, compare `session_status` in three Claude sessions, end one session and check the others, then intentionally send to your own chat. Test Graph/browser search paths separately to identify tenant permissions. Verify service crash recovery with systemd and reconnect clients afterward; saved sign-in persistence depends on Microsoft's session validity.

## References

- [Existing session-token implementation](../graphApi/README.md)
- [slirp4netns](https://github.com/rootless-containers/slirp4netns)
- [Playwright Electron support](https://playwright.dev/docs/api/class-electron)
- [Microsoft Graph search](https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview?view=graph-rest-1.0)
