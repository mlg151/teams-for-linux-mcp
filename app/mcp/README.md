# Headless Microsoft 365 MCP server

The server launches the existing Electron app with its existing login implementation and persistent session partition. On Linux, server mode runs inside a private virtual display using `xvfb-run`, even when a desktop display is available. Sign-in windows can render and receive input there without appearing on your real desktop. No call-control tools are registered.

This is an experimental implementation. Automated checks cover the MCP protocol, Graph request construction, permission fallback, browser interactions against controlled fixtures, and real signed-out headless startup. Live sign-in, people lookup, verified one-to-one chat resolution, and message sending through MCP have also been exercised with a Microsoft account. Other search/read and calendar operations still require validation against your tenant.

## Install

Use Node 24, as specified by `.nvmrc`, and install the repository dependencies:

```bash
npm ci
```

The Electron package must have downloaded its runtime. If your npm configuration disables install scripts, run this explicitly:

```bash
node node_modules/electron/install.js
```

On Debian/Ubuntu, install `xvfb` and `xauth` for server mode, along with the normal Electron system libraries listed in the repository's installation documentation. The MCP launcher reports a missing `xvfb-run` command. It does not install system packages or change sandbox/keyring settings automatically.

## Login through MCP

1. Configure the MCP client as shown below and connect it.
2. Call **`login`** with no arguments. It starts the headless app and returns a private browser link.
3. Open that link. It displays a live image of the app's sign-in window. Click the Microsoft field you want to use, type through the private input above the screen, and use **Enter / Next** or click the Microsoft buttons in the screen. Complete any MFA prompt there or in your authenticator app.
4. The page confirms completion and revokes the link. Call **`session_status`** in MCP; the same running session is ready to use. No restart or VNC desktop is required.

Passwords and MFA entries travel directly between this private page and the server's browser, outside the MCP conversation. Do not paste them into chat or tool arguments. The link itself grants temporary access to the sign-in screen; do not share it. It expires after ten minutes and is revoked on completion or when you choose **Close login link**. Closing the link does not log out of Microsoft or clear an existing saved session.

Authenticator approval, SMS, and typed verification codes can be completed through this view. Hardware security keys and passkeys depend on the server's device/browser environment; use another Microsoft sign-in method if the required device is only connected to your local computer. The server does not bypass MFA or tenant sign-in policies.

Keep the same persistent profile path, OS user, and keyring environment for subsequent launches. The profile contains your Microsoft session; keep it private and outside version control. Copying encrypted cookies to another machine is not a supported login method. The ordinary desktop app and the MCP runtime cannot share a profile simultaneously.

The optional `npm run mcp:login -- --data-dir /absolute/profile/path` command still opens the full desktop app when a display is available. It is not needed for the MCP browser-link flow.

## Connect an MCP client

Configure a local stdio MCP server. Use absolute paths; `~` does not expand inside JSON arguments:

```json
{
  "mcpServers": {
    "teams": {
      "command": "node",
      "args": [
        "/absolute/path/teams-for-linux-mcp/scripts/mcp.js",
        "--data-dir",
        "/home/you/.local/share/teams-mcp/profile"
      ]
    }
  }
}
```

Use `node scripts/mcp.js` directly in the client configuration. `npm run mcp` is convenient in a terminal but npm's banner is not part of the MCP protocol. The direct command reserves stdout for MCP; diagnostic messages go to stderr.

Electron starts lazily on the first tool call. Call `login` for a new session, or `session_status` to check an existing session. If sign-in has expired, call `login` again and use the new browser link; the server keeps running.

The profile defaults to the ordinary app's user-data directory when neither `--data-dir` nor `TEAMS_MCP_USER_DATA_DIR` is set. Supplying a dedicated directory makes profile selection explicit. Headless options are applied in memory; the server does not rewrite the desktop configuration.

## Browser link on a remote server

The login listener binds only to `127.0.0.1`. When the MCP server runs on your computer, open the returned link directly. When it runs on another machine, `127.0.0.1` in your browser refers to your computer, so forward the login port over SSH.

Set a fixed port in the MCP server's environment:

```json
"env": { "TEAMS_MCP_LOGIN_PORT": "8787" }
```

Then, on the computer running your browser:

```bash
ssh -N -L 8787:127.0.0.1:8787 your-server
```

Call `login` and open its returned `http://127.0.0.1:8787/#...` link. Microsoft sign-in still runs inside the app on the server; the tunnel carries the private screen and input. The tunnel must be active when opening the link.

Alternatively, an existing HTTPS reverse proxy can forward a dedicated origin to that local port. Set `TEAMS_MCP_LOGIN_PUBLIC_URL=https://your-login-host.example` and `TEAMS_MCP_LOGIN_PORT=8787`. The public URL must be an HTTPS origin with no path, query string, or credentials. The proxy must preserve the browser's Origin header and either the public Host or the local upstream Host. The server returns links at that HTTPS origin and still binds only to loopback. It does not configure or publish a reverse proxy itself.

For an IP-address link, Tailscale Serve can forward TCP privately within your tailnet. Set `TEAMS_MCP_LOGIN_PUBLIC_URL=http://YOUR_TAILSCALE_IP:8443`, `TEAMS_MCP_LOGIN_PORT=8787`, and explicitly set `TEAMS_MCP_LOGIN_ALLOW_TAILNET_HTTP=1`, then run `tailscale serve --bg --tcp=8443 tcp://127.0.0.1:8787`. TCP forwarding preserves IP-based Host headers; Tailscale's HTTP virtual-host proxy expects its DNS name. This exception accepts only Tailscale IPv4 addresses in `100.64.0.0/10`; it is intended for traffic carried by Tailscale, not an unencrypted public or LAN listener. The app continues to listen only on loopback.

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

```text
MCP stdio client
  -> Node MCP server
    -> Playwright-controlled existing Electron app
      -> Teams session authentication / visible web-app controls
```

`SessionGraph` reuses this repository's `GraphApiClient.acquireToken()` through the signed-in Teams renderer. Requests stay on `https://graph.microsoft.com/v1.0/`; continuation URLs are validated, redirects are rejected, and raw session tokens are never included in tool results.

This retains the existing session-token permission limits. Calendar and chat sending build on the repository's existing Graph functionality. Chat-history reads, channel access, and OneDrive API operations may require scopes not available in a particular Teams session. These tools return explicit permission errors; adding MCP does not grant extra access.

The browser search adapters use Teams' search box/category filters and the embedded OneDrive search view. They wait for a Microsoft search response matching the submitted query, then read the rendered result pane. This supports searches through the app's existing access without asking for broader Graph permissions.

Browser fallback currently targets English UI labels and known Teams/OneDrive layouts and search request patterns. It is fixture-tested, not yet tenant-tested. If those patterns do not match, it returns an explicit UI error instead of returning the previous screen as search results. Browser responses identify their source and state that they cover **visible results only**. They do not provide exhaustive history or Graph pagination; the `viewText` field is rendered text, not a structured message collection. The Teams/Channels search filter can also cover a broader scope than Graph's joined-Teams list.

All tools serialize operations against the single browser session. A cancelled queued operation is discarded before it starts. Sends and calendar writes are not automatically retried. If a request times out after dispatch, the result reports `outcome_unknown`: inspect Teams/calendar before retrying to avoid duplicates.

## Validation

```bash
npm run test:mcp
npm run test:mcp:browser
npm run test:mcp:runtime
npm run lint
```

The browser tests use a real Electron renderer with controlled search fixtures. The runtime smoke test starts the real stdio server with an empty temporary profile, exercises discovery/session status under Xvfb, and checks clean exit on client disconnect. Neither test sends Microsoft messages or creates calendar events.

The login tests also exercise private link access, expiry/revocation, stale screens, and a full username/password/MFA fixture through the live browser controls at desktop and mobile widths. The real runtime smoke test obtains a login link through MCP and captures the headless app's sign-in screen through its authenticated HTTP endpoint. These checks do not replace a live tenant sign-in test.

After first login, validate `session_status`, then each search with a known query. Try `backend: "graph"` and `backend: "browser"` separately to identify which paths your tenant permits. Validate sends and calendar writes only with an intentional test destination/event.

## References

- [Existing session-token implementation](../graphApi/README.md)
- [Electron on headless Linux](https://www.electronjs.org/docs/latest/tutorial/testing-on-headless-ci)
- [Playwright Electron support](https://playwright.dev/docs/api/class-electron)
- [Microsoft Graph search](https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview?view=graph-rest-1.0)
