'use strict';

const path = require('node:path');
const { setTimeout: delay } = require('node:timers/promises');
const { McpOperationError } = require('./errors');

const ROOT = path.resolve(__dirname, '../..');
const TEAMS_HOSTS = new Set(['teams.microsoft.com', 'teams.cloud.microsoft', 'teams.live.com']);

function isTeamsUrl(value) {
  try {
    const url = new URL(value);
    return url.protocol === 'https:' && TEAMS_HOSTS.has(url.hostname);
  } catch { return false; }
}

// Browser search remains usable when Teams changes its private auth-service
// tree. The signed-in search UI is an independent readiness signal.
function sessionReady() {
  try {
    if (window.teamsForLinuxReactHandler?._getTeams2CoreServices?.()?.authenticationService) return true;
  } catch { /* The browser UI can work independently of the private auth tree. */ }
  return !!(document.querySelector('[data-tid="search-f6-navigation-region"], [role="search"]') && document.querySelector('[role="main"]'));
}

class TeamsRuntime {
  constructor({ mode = 'headless', dataDir, launch, loginOptions, virtualDisplay = process.env.TEAMS_MCP_VIRTUAL_DISPLAY === '1' } = {}) {
    this.mode = mode;
    this.dataDir = dataDir;
    this.launch = launch;
    this.loginOptions = loginOptions;
    this.virtualDisplay = virtualDisplay;
    this.application = null;
    this.starting = null;
    this.closing = false;
    this.queue = Promise.resolve();
  }

  // All tools share one browser session. Serialize navigation and sends so a
  // concurrent search cannot change the conversation underneath another tool.
  run(operation, { signal } = {}) {
    const next = this.queue.then(async () => {
      if (this.closing) throw new McpOperationError('shutting_down', 'The Teams session is closing.');
      if (signal?.aborted) throw new McpOperationError('cancelled', 'The operation was cancelled before it started.');
      this.signal = signal;
      try { return await operation(); }
      finally { this.signal = undefined; }
    });
    this.queue = next.catch(() => {});
    return next;
  }

  async start() {
    if (this.closing) throw new McpOperationError('shutting_down', 'The Teams session is closing.');
    if (this.application) return this.application;
    if (!this.starting) this.starting = this.#launch().finally(() => { this.starting = null; });
    return this.starting;
  }

  async #launch() {
    const launch = this.launch || require('playwright-core')._electron.launch.bind(require('playwright-core')._electron);
    const env = { ...process.env, TEAMS_MCP_MODE: this.mode };
    delete env.ELECTRON_RUN_AS_NODE;
    if (this.dataDir) env.TEAMS_MCP_USER_DATA_DIR = path.resolve(this.dataDir);
    const args = [path.join(ROOT, 'app/index.js')];
    if (process.platform === 'linux' && process.env.DISPLAY) args.push('--ozone-platform=x11');
    try {
      this.application = await launch({
        executablePath: require('electron'), args, env, cwd: ROOT, timeout: 45000,
      });
      this.application.on('close', () => { this.application = null; this.page = null; });
      await this.application.firstWindow({ timeout: 45000 });
      return this.application;
    } catch {
      await this.closeApplication();
      throw new McpOperationError('startup_failed',
        'Could not start Teams. Close other instances using this profile, check Electron system libraries, and ensure a display or Xvfb is available.');
    }
  }

  async getPage() {
    const application = await this.start();
    if (this.page && !this.page.isClosed()) return this.page;
    const deadline = Date.now() + 30000;
    while (Date.now() < deadline) {
      const mainId = await application.evaluate((_electron, modulePath) =>
        process.getBuiltinModule('module').createRequire(modulePath)(modulePath).getWindow()?.id,
      path.join(ROOT, 'app/mainAppWindow/index.js'));
      for (const page of application.windows()) {
        if (page.isClosed()) continue;
        const handle = await application.browserWindow(page);
        try {
          if (await handle.evaluate(window => window.id) === mainId) {
            this.page = page;
            page.setDefaultTimeout(15000);
            return page;
          }
        } finally { await handle.dispose(); }
      }
      await delay(200);
    }
    throw new McpOperationError('startup_failed', 'The main Teams window did not become available.');
  }

  async status() {
    const page = await this.getPage();
    const onTeams = isTeamsUrl(page.url());
    const authenticated = onTeams && await page.evaluate(sessionReady).catch(() => false);
    return {
      authenticated, mode: this.mode,
      state: authenticated ? 'ready' : onTeams ? 'loading_or_login_required' : 'login_required',
      // Never return URLs: login redirect query strings may contain auth codes.
      origin: new URL(page.url()).origin,
      ...(authenticated ? {} : { action: 'Call the MCP login tool and open its private browser link to complete Microsoft sign-in and MFA.' }),
    };
  }

  async requireSession() {
    const page = await this.getPage();
    if (!isTeamsUrl(page.url())) {
      throw new McpOperationError('login_required', 'Call the MCP login tool and complete sign-in using its private browser link.');
    }
    await page.waitForFunction(sessionReady, null, { timeout: 30000 }).catch(() => {
      throw new McpOperationError('login_required', 'Teams is still loading or requires sign-in. Check session_status or call the MCP login tool.');
    });
    return page;
  }

  async beginLogin() {
    const status = await this.status();
    if (status.authenticated) return status;
    if (!this.loginPortal || this.loginPortal.closed) {
      const { LoginPortal, portalOptions } = require('./loginPortal');
      this.loginPortal = new LoginPortal(this, this.loginOptions || portalOptions());
    }
    return this.loginPortal.start();
  }

  async getLoginPage() {
    const main = await this.getPage();
    // Microsoft/organization SSO can open a popup. Ignore the app's file/data
    // notification windows and drive the most recent HTTP(S) sign-in popup.
    const popups = this.application.windows().filter(page => page !== main && !page.isClosed() && /^https?:\/\//.test(page.url()));
    return popups.at(-1) || main;
  }

  async captureLoginScreen(page) {
    const application = await this.start();
    const window = await application.browserWindow(page);
    try {
      // Native capture does not wait on remote web fonts or Playwright's
      // screenshot animations. Showing inside private Xvfb permits painting
      // without opening a window on the user's real desktop.
      return await window.evaluate(async (window, virtualDisplay) => {
        if (virtualDisplay && !window.isVisible()) window.showInactive();
        const image = await window.webContents.capturePage(undefined, { stayHidden: !virtualDisplay, stayAwake: true });
        if (image.isEmpty()) throw new Error('The sign-in screen has not painted yet.');
        return image.toPNG().toString('base64');
      }, this.virtualDisplay);
    } finally { await window.dispose(); }
  }

  async focusLoginPage(page) {
    const application = await this.start();
    const window = await application.browserWindow(page);
    try {
      await window.evaluate((window, virtualDisplay) => {
        if (virtualDisplay) { window.show(); window.focus(); }
        window.webContents.focus();
      }, this.virtualDisplay);
    } finally { await window.dispose(); }
  }

  async closeApplication() {
    const application = this.application;
    this.application = null;
    this.page = null;
    if (!application) return;
    let timer;
    try {
      await Promise.race([
        application.close(),
        new Promise((_, reject) => { timer = setTimeout(() => reject(new Error('Close timed out')), 5000); }),
      ]);
    } catch { application.process()?.kill('SIGKILL'); }
    finally { clearTimeout(timer); }
  }

  async close() {
    this.closing = true;
    this.loginPortal?.close();
    await this.starting?.catch(() => {});
    await this.closeApplication();
  }
}

module.exports = { TeamsRuntime, isTeamsUrl };
