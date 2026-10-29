'use strict';

const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
const { isIP } = require('node:net');
const { randomBytes, timingSafeEqual, randomUUID } = require('node:crypto');
const { z } = require('zod/v4');
const { McpOperationError } = require('./errors');

const KEYS = ['Enter', 'Tab', 'Shift+Tab', 'Backspace', 'Delete', 'Escape', 'ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End', 'Control+A', 'Meta+A', 'Space'];
const actionSchema = z.discriminatedUnion('type', [
  z.object({ type: z.literal('click'), x: z.number().finite().nonnegative(), y: z.number().finite().nonnegative() }),
  z.object({ type: z.literal('type'), text: z.string().min(1).max(4096) }),
  z.object({ type: z.literal('key'), key: z.enum(KEYS) }),
  z.object({ type: z.literal('scroll'), deltaY: z.number().finite().min(-3000).max(3000) }),
]);
const requestSchema = z.object({ viewId: z.string().uuid(), action: actionSchema });
const ASSETS = new Map([
  ['/', ['login.html', 'text/html; charset=utf-8']],
  ['/login.js', ['login.js', 'text/javascript; charset=utf-8']],
  ['/login.css', ['login.css', 'text/css; charset=utf-8']],
]);

function portalOptions(env = process.env) {
  const port = env.TEAMS_MCP_LOGIN_PORT === undefined ? 0 : Number(env.TEAMS_MCP_LOGIN_PORT);
  if (!Number.isInteger(port) || port < 0 || port > 65535) {
    throw new McpOperationError('login_configuration', 'TEAMS_MCP_LOGIN_PORT must be a port number between 0 and 65535.');
  }
  let publicOrigin;
  if (env.TEAMS_MCP_LOGIN_PUBLIC_URL) {
    let url;
    try { url = new URL(env.TEAMS_MCP_LOGIN_PUBLIC_URL); } catch { /* Validated below. */ }
    const octets = url?.hostname.split('.').map(Number);
    const tailnetHttp = env.TEAMS_MCP_LOGIN_ALLOW_TAILNET_HTTP === '1' && url?.protocol === 'http:' &&
      isIP(url.hostname) === 4 && octets[0] === 100 && octets[1] >= 64 && octets[1] <= 127;
    if (!url || (url.protocol !== 'https:' && !tailnetHttp) || url.username || url.password || url.pathname !== '/' || url.search || url.hash || !port) {
      throw new McpOperationError('login_configuration', 'Use an HTTPS origin and a fixed TEAMS_MCP_LOGIN_PORT. HTTP on a Tailscale IPv4 address requires explicit TEAMS_MCP_LOGIN_ALLOW_TAILNET_HTTP=1. Paths, queries, and credentials are not allowed.');
    }
    publicOrigin = url.origin;
  }
  return { port, publicOrigin };
}

async function readJson(request) {
  if (request.headers['content-type']?.split(';')[0] !== 'application/json') throw new Error('Invalid content type');
  let size = 0;
  const chunks = [];
  for await (const chunk of request) {
    size += chunk.length;
    if (size > 32768) throw new Error('Request too large');
    chunks.push(chunk);
  }
  return JSON.parse(Buffer.concat(chunks).toString('utf8'));
}

class LoginPortal {
  constructor(runtime, { port = 0, publicOrigin, ttlMs = 10 * 60 * 1000 } = {}) {
    this.runtime = runtime;
    this.port = port;
    this.publicOrigin = publicOrigin;
    this.ttlMs = ttlMs;
    this.token = randomBytes(32).toString('base64url');
    this.closed = false;
    this.reason = 'pending';
    this.pages = new WeakMap();
  }

  async start() {
    if (!this.server) {
      this.server = http.createServer((request, response) => {
        this.handle(request, response).catch(() => {
          if (!response.headersSent) this.json(response, 502, { error: 'The login screen is changing or unavailable. Refresh it and try again.' });
          else response.end();
        });
      });
      this.server.requestTimeout = 15000;
      this.server.headersTimeout = 10000;
      this.server.keepAliveTimeout = 1000;
      await new Promise((resolve, reject) => {
        this.server.once('error', reject);
        this.server.listen(this.port, '127.0.0.1', resolve);
      }).catch(() => {
        this.close();
        throw new McpOperationError('login_unavailable', 'Could not open the browser login link. Check whether TEAMS_MCP_LOGIN_PORT is already in use.');
      });
      this.localOrigin = `http://127.0.0.1:${this.server.address().port}`;
      this.origin = this.publicOrigin || this.localOrigin;
      this.expiresAt = Date.now() + this.ttlMs;
      this.timer = setTimeout(() => this.finish('expired'), this.ttlMs);
      this.timer.unref();
    }
    return {
      authenticated: false, state: 'awaiting_login', url: `${this.origin}/#${this.token}`,
      expiresAt: new Date(this.expiresAt).toISOString(),
      action: 'Open this private link, complete Microsoft sign-in and MFA, then call session_status. No desktop/VNC is needed. Do not enter your password or MFA codes into MCP messages.',
    };
  }

  headers(response) {
    response.setHeader('Cache-Control', 'no-store');
    response.setHeader('Referrer-Policy', 'no-referrer');
    response.setHeader('X-Content-Type-Options', 'nosniff');
    response.setHeader('Content-Security-Policy', "default-src 'none'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'");
  }

  json(response, status, value) {
    this.headers(response);
    response.writeHead(status, { 'Content-Type': 'application/json' });
    response.end(JSON.stringify(value));
  }

  authorized(request) {
    const supplied = request.headers.authorization || '';
    const expected = `Bearer ${this.token}`;
    return !this.closed && Date.now() < this.expiresAt &&
      Buffer.byteLength(supplied) === Buffer.byteLength(expected) && timingSafeEqual(Buffer.from(supplied), Buffer.from(expected));
  }

  async handle(request, response) {
    this.headers(response);
    const allowedHosts = new Set([new URL(this.localOrigin).host, new URL(this.origin).host]);
    if (!allowedHosts.has(request.headers.host) ||
      (request.headers.origin && request.headers.origin !== this.origin)) {
      return this.json(response, 403, { error: 'Invalid request origin.' });
    }
    const url = new URL(request.url, this.localOrigin);
    if (request.method === 'GET' && ASSETS.has(url.pathname)) {
      const [file, contentType] = ASSETS.get(url.pathname);
      response.writeHead(200, { 'Content-Type': contentType });
      return response.end(fs.readFileSync(path.join(__dirname, 'login', file)));
    }
    if (!this.authorized(request)) return this.json(response, 401, { error: 'This login link is invalid, finished, or expired. Call the MCP login tool for a new link.' });
    if (request.method === 'GET' && url.pathname === '/api/state') {
      const state = await this.runtime.run(() => this.snapshot());
      return this.json(response, 200, state);
    }
    if (request.method === 'POST') {
      if (request.headers.origin !== this.origin) return this.json(response, 403, { error: 'Invalid request origin.' });
      if (url.pathname === '/api/cancel') {
        this.json(response, 200, { state: 'cancelled' });
        this.finish('cancelled');
        return;
      }
      if (url.pathname === '/api/action') {
        let input;
        try { input = requestSchema.parse(await readJson(request)); }
        catch { return this.json(response, 400, { error: 'Invalid login action.' }); }
        const result = await this.runtime.run(() => this.perform(input));
        return this.json(response, result.stale ? 409 : 200, result);
      }
    }
    return this.json(response, 404, { error: 'Not found.' });
  }

  async snapshot() {
    if (this.closed) return { state: this.reason };
    const status = await this.runtime.status();
    if (status.authenticated) {
      this.finish('complete');
      return { state: 'complete', authenticated: true };
    }
    const page = await this.runtime.getLoginPage();
    let view = this.pages.get(page);
    if (!view) {
      view = { id: randomUUID() };
      this.pages.set(page, view);
      page.on('framenavigated', frame => { if (frame === page.mainFrame()) view.id = randomUUID(); });
    }
    const viewId = view.id;
    const viewport = await page.evaluate(() => ({ width: innerWidth, height: innerHeight }));
    const image = await this.runtime.captureLoginScreen(page);
    if ((await this.runtime.status()).authenticated) {
      this.finish('complete');
      return { state: 'complete', authenticated: true };
    }
    if (view.id !== viewId) return { state: 'loading' };
    this.currentView = { page, id: viewId, ...viewport };
    return { state: 'pending', origin: new URL(page.url()).origin, viewId, ...viewport, image };
  }

  async perform({ viewId, action }) {
    if (this.closed) return { state: this.reason };
    if ((await this.runtime.status()).authenticated) {
      this.finish('complete');
      return { state: 'complete', authenticated: true };
    }
    const page = await this.runtime.getLoginPage();
    if (this.currentView?.page !== page || this.currentView.id !== viewId || this.pages.get(page)?.id !== viewId) {
      return { stale: true, error: 'The login screen changed. Wait for the updated image before typing again.' };
    }
    await this.runtime.focusLoginPage(page);
    if (this.closed) return { state: this.reason };
    if (this.pages.get(page)?.id !== viewId) return { stale: true, error: 'The login screen changed. Wait for the updated image before typing again.' };
    if (action.type === 'click') {
      if (action.x >= this.currentView.width || action.y >= this.currentView.height) return { stale: true, error: 'Click inside the login screen.' };
      await page.mouse.click(action.x, action.y);
    } else if (action.type === 'type') {
      const inputs = await Promise.all(page.frames().map(frame => frame.evaluate(() => {
        let element = document.activeElement;
        while (element?.shadowRoot?.activeElement) element = element.shadowRoot.activeElement;
        return document.hasFocus() && !!element?.matches('input:not([disabled]):not([readonly]), textarea:not([disabled]):not([readonly])');
      }).catch(() => false)));
      if (!inputs.some(Boolean)) return { stale: true, error: 'Click a sign-in input in the live screen before typing.' };
      await page.keyboard.insertText(action.text);
    }
    else if (action.type === 'key') await page.keyboard.press(action.key);
    else if (action.type === 'scroll') await page.mouse.wheel(0, action.deltaY);
    return { state: 'pending' };
  }

  finish(reason) {
    if (this.closed) return;
    this.closed = true;
    this.reason = reason;
    this.token = null;
    this.currentView = null;
    clearTimeout(this.timer);
    // Finish the current HTTP response before closing the listener. Never wait
    // on server.close from inside a request that itself must finish first.
    setImmediate(() => { this.server?.close(); this.server?.closeIdleConnections(); });
  }

  close() {
    this.finish('cancelled');
    this.server?.closeAllConnections();
  }
}

module.exports = { LoginPortal, portalOptions };
