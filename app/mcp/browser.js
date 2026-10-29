'use strict';

const { McpOperationError } = require('./errors');
const { setTimeout: delay } = require('node:timers/promises');

const CATEGORIES = {
  messages: /^Messages$/i,
  chats: /^(Chats|Group chats)(?:Group chats)?$/i,
  teams: /^(Teams|Teams and channels|Channels)$/i,
};

function isSearchResponse(response, query) {
  try {
    const url = new URL(response.url());
    const trusted = url.hostname === 'graph.microsoft.com' ||
      /(^|\.)(teams\.cloud\.microsoft|teams\.microsoft\.com|office\.com|sharepoint\.com|onedrive\.cloud\.microsoft)$/.test(url.hostname);
    if (url.protocol !== 'https:' || !trusted || !/search|\/query(?:\/|$)/i.test(url.pathname)) return false;
    const requestText = `${decodeURIComponent(url.href)} ${response.request().postData() || ''}`;
    return requestText.includes(query) || requestText.includes(JSON.stringify(query).slice(1, -1));
  } catch { return false; }
}

async function firstVisible(locators, description) {
  const combined = locators.reduce((previous, current) => previous.or(current));
  await combined.filter({ visible: true }).first().waitFor({ state: 'visible', timeout: 15000 }).catch(() => {
    throw new McpOperationError('ui_unavailable', `Could not find ${description}. The Teams layout or language may differ from the supported selectors.`);
  });
  for (const locator of locators) {
    const candidates = await locator.all();
    for (const candidate of candidates) {
      if (await candidate.isVisible()) return candidate;
    }
  }
  throw new McpOperationError('ui_unavailable', `Could not find ${description}. The Teams layout or language may differ from the supported selectors.`);
}

function searchInputs(frame) {
  return [
    frame.locator('[data-tid="search-f6-navigation-region"] input'),
    frame.getByRole('combobox', { name: /search/i }),
    frame.getByRole('searchbox'),
    frame.getByPlaceholder(/search/i),
    frame.locator('[role="search"] input'),
  ];
}

// Return only the rendered result surface, not an invented API response or an
// exhaustive history. UI virtualization and pagination are explicit limitations.
async function readSurface(frame, limit, source, extra = {}) {
  const root = await firstVisible([
    frame.locator('[data-tid="search-results"]'),
    frame.locator('[data-tid="search-results-container"]'),
    frame.getByRole('main'),
  ], 'the search result area');
  const data = await root.evaluate((element, limit) => {
    const visible = node => !!(node.getClientRects().length && getComputedStyle(node).visibility !== 'hidden');
    const nodes = [...element.querySelectorAll('[data-tid*="search-result"], [data-tid*="search-hit"], [role="listitem"], [role="row"], article')];
    const seen = new Set();
    const items = [];
    for (const node of nodes) {
      if (!visible(node)) continue;
      const text = node.innerText?.trim();
      if (!text || seen.has(text)) continue;
      seen.add(text);
      const hrefs = [...node.querySelectorAll('a[href]')].map(link => link.href).filter(href => /^https:\/\//.test(href));
      const attributes = ['id', 'data-convid', 'data-chat-id', 'data-tid'].map(name => node.getAttribute(name) || '').join(' ');
      const decodedLinks = hrefs.map(href => { try { return decodeURIComponent(href); } catch { return href; } });
      const conversationId = `${attributes} ${decodedLinks.join(' ')}`.match(/19:[a-zA-Z0-9_-]+@(?:thread\.(?:v2|tacv2)|unq\.gbl\.spaces)/)?.[0];
      const target = conversationId ? conversationId.endsWith('@thread.tacv2') ? { channelId: conversationId } : { chatId: conversationId } : {};
      for (const href of hrefs) {
        const teamId = new URL(href).searchParams.get('groupId');
        if (teamId && /^[a-f0-9-]{36}$/i.test(teamId)) target.teamId = teamId;
      }
      items.push({ text: text.slice(0, 6000), ...(hrefs.length ? { links: hrefs.slice(0, 5) } : {}), ...target });
      if (items.length >= limit) break;
    }
    return { items, viewText: (element.innerText || '').slice(0, 24000) };
  }, limit);
  return {
    source, ...extra, ...data,
    coverage: 'Visible results only; the web app may load additional results when scrolled. viewText is the rendered result pane, not structured message history.',
  };
}

class TeamsBrowser {
  constructor(runtime) { this.runtime = runtime; }

  async search(category, query, limit = 20) {
    const page = await this.runtime.requireSession();
    const input = await firstVisible(searchInputs(page), 'the Teams global search box');
    await input.fill(query);
    await this.searchAction(page, query, () => input.press('Enter'));
    const label = CATEGORIES[category];
    const tab = page.getByRole('tab', { name: label }).or(page.getByRole('button', { name: label }));
    await tab.first().waitFor({ state: 'visible', timeout: 15000 }).catch(() => {
      throw new McpOperationError('ui_unavailable', `The ${category} search filter did not appear. This Teams layout or UI language needs updated selectors.`);
    });
    const selected = await tab.first().getAttribute('aria-selected') === 'true' ||
      await tab.first().getAttribute('aria-pressed') === 'true';
    if (!selected) await this.searchAction(page, query, () => tab.first().click());
    await this.waitForResults(page);
    return readSurface(page, limit, 'teams_browser', { query, category });
  }

  async searchAction(frame, query, action) {
    const page = typeof frame.page === 'function' ? frame.page() : frame;
    // Correlate completion with this query. A generic idle/loading check alone
    // can return the previous query's results before a new request begins.
    const responsePromise = page.waitForResponse(response => isSearchResponse(response, query), { timeout: 15000 }).catch(() => null);
    await action();
    const response = await responsePromise;
    if (!response) throw new McpOperationError('ui_search_unverified',
      'Could not verify completion of this search. No results were returned; the web app may have changed its search flow.');
    if (!response.ok()) throw new McpOperationError('ui_search_failed', `The web app search returned HTTP ${response.status()}.`);
    await response.finished();
  }

  async waitForResults(frame) {
    // Teams exposes loading with progressbar/aria-busy. Observe at least one
    // render turn before waiting, including for cached search results.
    await frame.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
    await frame.locator('[role="progressbar"]:visible, [aria-busy="true"]:visible').first()
      .waitFor({ state: 'hidden', timeout: 15000 }).catch(() => {
        throw new McpOperationError('ui_timeout', 'Search results are still loading. Try the search again.');
      });
  }

  async searchOneDrive(query, limit = 20) {
    const page = await this.runtime.requireSession();
    const nav = await firstVisible([
      page.getByRole('button', { name: /^OneDrive(?: \(.*\))?$/i }),
      page.getByRole('link', { name: /^OneDrive$/i }),
      page.getByRole('tab', { name: /^OneDrive$/i }),
    ], 'OneDrive in the Teams app rail (pin it in Teams if needed)');
    await nav.click();
    await page.waitForFunction(() => [...document.querySelectorAll('iframe')].some(frame => {
      try { return /(^|\.)(sharepoint\.com|onedrive\.com|onedrive\.cloud\.microsoft)$/.test(new URL(frame.src).hostname); }
      catch { return false; }
    }), null, { timeout: 15000 }).catch(() => {
      throw new McpOperationError('ui_unavailable', 'The embedded OneDrive view did not load. Open or pin OneDrive in Teams and try again.');
    });
    const oneDriveFrame = frame => {
      try { return /(^|\.)(sharepoint\.com|onedrive\.com|onedrive\.cloud\.microsoft)$/.test(new URL(frame.url()).hostname); }
      catch { return false; }
    };
    let frame;
    const deadline = Date.now() + 15000;
    while (!(frame = page.frames().find(oneDriveFrame)) && Date.now() < deadline) await delay(100);
    if (!frame) throw new McpOperationError('ui_unavailable', 'Could not locate the signed-in OneDrive frame.');
    await frame.waitForLoadState('domcontentloaded');
    const input = await firstVisible(searchInputs(frame), 'the OneDrive search box');
    await input.fill(query);
    await this.searchAction(frame, query, () => input.press('Enter'));
    await this.waitForResults(frame);
    return readSurface(frame, limit, 'onedrive_browser', { query });
  }
}

module.exports = { TeamsBrowser, firstVisible, readSurface, isSearchResponse };
