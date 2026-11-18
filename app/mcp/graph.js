'use strict';

const { randomUUID } = require('node:crypto');
const { McpOperationError } = require('./errors');

const BASE = 'https://graph.microsoft.com/v1.0';

function graphUrl(endpoint) {
  const url = new URL(endpoint.startsWith('/') ? `${BASE}${endpoint}` : endpoint);
  if (url.origin !== 'https://graph.microsoft.com' || !url.pathname.startsWith('/v1.0/') || url.username || url.password || url.hash) {
    throw new McpOperationError('invalid_endpoint', 'Only Microsoft Graph v1.0 endpoints are allowed.');
  }
  return url.href;
}

// Reuse the repo's session-token acquisition; no client ID, client secret, or
// separate OAuth flow. Tokens stay inside this process and never enter results.
class SessionGraph {
  constructor(runtime, { tokenClient, fetchImpl = fetch } = {}) {
    this.runtime = runtime;
    this.fetch = fetchImpl;
    this.tokenClient = tokenClient;
    this.cursors = new Map();
  }

  async getTokenClient() {
    if (!this.tokenClient) {
      const GraphApiClient = require('../graphApi');
      this.tokenClient = new GraphApiClient({ graphApi: { enabled: true } });
      this.tokenClient.initialize({ webContents: {
        executeJavaScript: async script => (await this.runtime.requireSession()).evaluate(script),
      } });
      const page = await this.runtime.getPage();
      page.on('framenavigated', frame => {
        if (frame === page.mainFrame()) {
          this.tokenClient.currentToken = null;
          this.tokenClient.tokenExpiry = null;
          this.cursors.clear();
        }
      });
    }
    return this.tokenClient;
  }

  async request(endpoint, { method = 'GET', body, readOnly = method === 'GET' } = {}) {
    const url = graphUrl(endpoint);
    await this.runtime.requireSession();
    const client = await this.getTokenClient();
    const tokenResult = await client.acquireToken();
    if (!tokenResult?.success || !tokenResult.token) {
      throw new McpOperationError('session_api_unavailable', 'The Teams session could not provide a Graph token. Try the browser search tools or sign in again.');
    }
    if (this.runtime.signal?.aborted) throw new McpOperationError('cancelled', 'The operation was cancelled before sending the request.');
    let response;
    try {
      response = await this.fetch(url, {
        method, redirect: 'error', signal: AbortSignal.any([AbortSignal.timeout(30000), ...(this.runtime.signal ? [this.runtime.signal] : [])]),
        headers: { Authorization: `Bearer ${tokenResult.token}`, 'Content-Type': 'application/json' },
        ...(body === undefined ? {} : { body: JSON.stringify(body) }),
      });
    } catch {
      throw new McpOperationError(readOnly ? 'request_failed' : 'outcome_unknown',
        readOnly ? 'The Microsoft Graph request failed or timed out.' :
          'The request ended without a response. Check Teams before retrying; the write may have succeeded.');
    }
    const data = await response.json().catch(() => null);
    if (!response.ok) {
      if (response.status === 401) {
        client.currentToken = null;
        client.tokenExpiry = null;
      }
      const code = response.status === 403 ? 'permission_denied' : response.status === 401 ? 'login_required' :
        response.status === 429 ? 'rate_limited' : response.status >= 500 && !readOnly ? 'outcome_unknown' : 'graph_error';
      throw new McpOperationError(code,
        response.status === 403 ? 'This operation is not permitted by the Teams session Graph token.' :
          code === 'outcome_unknown' ? 'Microsoft returned a server error. Check Teams before retrying this write.' :
            `Microsoft Graph returned HTTP ${response.status}.`,
        { status: response.status, ...(response.headers.get('retry-after') ? { retryAfter: response.headers.get('retry-after') } : {}) });
    }
    if (readOnly && data === null) throw new McpOperationError('invalid_response', 'Microsoft Graph returned an unreadable response. No results were returned.');
    return data;
  }

  // Binary transfer variant: returns the parsed Response without consuming
  // its body so the caller can stream it to or from disk. Graph-issued
  // upload URLs carry their own credentials and never receive a token.
  async requestStream(endpoint, { method = 'GET', url, headers, options = {} } = {}) {
    const target = url || graphUrl(endpoint);
    const parsed = new URL(target);
    if (parsed.protocol !== 'https:' || parsed.username || parsed.password || parsed.hash) {
      throw new McpOperationError('invalid_endpoint', 'Only HTTPS Microsoft Graph or Graph-issued upload URLs are allowed.');
    }
    const requestHeaders = { ...headers };
    let client;
    if (!url) {
      await this.runtime.requireSession();
      client = await this.getTokenClient();
      const tokenResult = await client.acquireToken();
      if (!tokenResult?.success || !tokenResult.token) {
        throw new McpOperationError('session_api_unavailable', 'The Teams session could not provide a Graph token. Sign in again and retry.');
      }
      requestHeaders.Authorization = `Bearer ${tokenResult.token}`;
    }
    if (this.runtime.signal?.aborted) throw new McpOperationError('cancelled', 'The operation was cancelled before sending the request.');
    if (options.body !== undefined && !('Content-Type' in requestHeaders)) requestHeaders['Content-Type'] = options.contentType || 'application/json';
    if (options.contentLength !== undefined && !('Content-Length' in requestHeaders)) requestHeaders['Content-Length'] = String(options.contentLength);
    let response;
    try {
      // Transfers have no fixed 30 second ceiling; they honor the caller's
      // cancellation signal instead.
      response = await this.fetch(target, {
        method, redirect: method === 'GET' && !url ? 'manual' : 'error', headers: requestHeaders,
        ...(options.body === undefined ? {} : { body: options.body, duplex: 'half' }),
        ...(this.runtime.signal ? { signal: this.runtime.signal } : {}),
      });
    } catch {
      if (this.runtime.signal?.aborted) throw new McpOperationError('cancelled', 'The transfer was cancelled.');
      throw new McpOperationError(method === 'GET' ? 'request_failed' : 'outcome_unknown',
        method === 'GET' ? 'The Microsoft Graph request failed or timed out.' :
          'The request ended without a response. Check the drive before retrying; the write may have succeeded.');
    }
    // Graph content endpoints redirect to a short-lived, preauthenticated URL.
    if (method === 'GET' && !url && response.status === 302) {
      await response.body?.cancel();
      const location = response.headers.get('location');
      if (!location || !URL.canParse(location)) throw new McpOperationError('invalid_response', 'Microsoft Graph returned an invalid download location.');
      return this.requestStream(null, { url: location });
    }
    if (client && response.status === 401) {
      client.currentToken = null;
      client.tokenExpiry = null;
    }
    return response;
  }

  async collection(endpoint, { cursor, scope = endpoint } = {}) {
    if (cursor) {
      const entry = this.cursors.get(cursor);
      if (!entry || entry.scope !== scope) throw new McpOperationError('invalid_cursor', 'The cursor is expired or belongs to another query.');
      endpoint = entry.url;
    }
    const data = await this.request(endpoint);
    if (!Array.isArray(data?.value)) throw new McpOperationError('invalid_response', 'Microsoft Graph did not return a collection. No results were returned.');
    let nextCursor;
    if (data?.['@odata.nextLink']) {
      const url = graphUrl(data['@odata.nextLink']);
      nextCursor = randomUUID();
      if (this.cursors.size >= 256) this.cursors.delete(this.cursors.keys().next().value);
      this.cursors.set(nextCursor, { url, scope });
    }
    return { source: 'graph', items: data?.value || [], ...(nextCursor ? { nextCursor } : {}) };
  }
}

function canUseBrowser(error) {
  return ['permission_denied', 'session_api_unavailable'].includes(error.code);
}

module.exports = { SessionGraph, graphUrl, canUseBrowser };
