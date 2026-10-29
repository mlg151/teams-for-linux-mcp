'use strict';

// Browser plugin not available. Exercise the browser adapter in a real Electron
// renderer with controlled fixtures; no Microsoft credentials or messages.
const { before, after, test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { _electron } = require('playwright-core');
const { TeamsBrowser, isSearchResponse } = require('../../app/mcp/browser');

let application;
let page;
let fixtureDir;

const fixture = `<!doctype html><html><head><title>Search fixture</title></head><body>
<div role="search"><input type="search" aria-label="Search Teams"></div>
<div role="tablist"><button role="tab" aria-selected="true">Messages</button><button role="tab" aria-selected="false">Group chats</button><button role="tab" aria-selected="false">Teams and channels</button></div>
<button id="onedrive">OneDrive</button><main><p>Previous conversation must never be returned as search results.</p></main>
<script>
(() => {
let category = 'Messages';
async function search() {
  const main = document.querySelector('main');
  main.setAttribute('aria-busy', 'true');
  const query = document.querySelector('input').value;
  const response = await fetch('https://substrate.office.com/search/api/v1/query', {method:'POST',headers:{'Content-Type':'text/plain'},body:JSON.stringify({query,category})});
  await response.json();
  main.replaceChildren();
  for(let i=0;i<3;i++) {
    const item=document.createElement('article');item.textContent=category+': '+query+' '+i;
    if(i===0 && category==='Messages') { const link=document.createElement('a');link.href='https://teams.microsoft.com/l/chat/19%3Afixture%40thread.v2/conversations';link.textContent='Open chat';item.append(link); }
    main.append(item);
  }
  main.setAttribute('aria-busy', 'false');
}
document.querySelector('input').addEventListener('keydown',e=>{if(e.key==='Enter')search()});
document.querySelectorAll('[role=tab]').forEach(button=>button.addEventListener('click',()=>{
  document.querySelectorAll('[role=tab]').forEach(b=>b.setAttribute('aria-selected',String(b===button)));
  category=button.textContent;search();
}));
document.getElementById('onedrive').onclick=()=>{const frame=document.createElement('iframe');frame.src='https://fixture-my.sharepoint.com/onedrive';document.body.append(frame);};
})();
</script></body></html>`;

before(async () => {
  fixtureDir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-mcp-ui-'));
  const main = path.join(fixtureDir, 'main.js');
  fs.writeFileSync(main, `const {app,BrowserWindow}=require('electron');app.setPath('userData',${JSON.stringify(path.join(fixtureDir, 'profile'))});let window;app.whenReady().then(()=>{window=new BrowserWindow({show:false,webPreferences:{backgroundThrottling:false}});window.loadURL('about:blank')});app.on('window-all-closed',()=>app.quit());`);
  const env = { ...process.env };
  delete env.ELECTRON_RUN_AS_NODE;
  application = await _electron.launch({ executablePath: require('electron'), args: [main], env, timeout: 30000 });
  page = await application.firstWindow();
  await page.route('https://substrate.office.com/search/api/v1/query', async route => {
    await new Promise(resolve => setTimeout(resolve, 150));
    await route.fulfill({ status: 200, contentType: 'application/json', headers: { 'access-control-allow-origin': '*' }, body: '{}' });
  });
}, { timeout: 45000 });

after(async () => {
  await application?.close();
  if (fixtureDir) fs.rmSync(fixtureDir, { recursive: true, force: true });
});

test('message browser search waits for this query and limits returned items', async () => {
  await page.setContent(fixture);
  const browser = new TeamsBrowser({ requireSession: async () => page });
  const result = await browser.search('messages', 'quarterly budget', 2);
  assert.equal(await page.title(), 'Search fixture');
  assert.equal(result.items.length, 2);
  assert.match(result.items[0].text, /Messages: quarterly budget/);
  assert.equal(result.items[0].chatId, '19:fixture@thread.v2');
  assert.equal(result.viewText.includes('Previous conversation'), false);
  assert.match(result.coverage, /Visible results only/);
});

test('chat and team searches switch the actual filter instead of returning message hits', async () => {
  await page.setContent(fixture);
  const browser = new TeamsBrowser({ requireSession: async () => page });
  assert.match((await browser.search('chats', 'project')).items[0].text, /Group chats: project/);
  assert.match((await browser.search('teams', 'engineering')).items[0].text, /Teams and channels: engineering/);
});

test('OneDrive browser search enters the signed-in cross-origin frame', async () => {
  await page.setContent(fixture);
  await page.route('https://fixture-my.sharepoint.com/onedrive', route => route.fulfill({
    contentType: 'text/html', body: `<!doctype html><input type="search" aria-label="Search OneDrive"><main>Old files</main><script>document.querySelector('input').onkeydown=async e=>{if(e.key!=='Enter')return;const main=document.querySelector('main');main.setAttribute('aria-busy','true');const q=e.target.value;await fetch('/_api/search/query?q='+encodeURIComponent(q));main.innerHTML='<article>OneDrive: '+q+'</article>';main.setAttribute('aria-busy','false');};</script>`,
  }));
  await page.route('https://fixture-my.sharepoint.com/_api/search/**', route => route.fulfill({ contentType: 'application/json', body: '{}' }));
  const browser = new TeamsBrowser({ requireSession: async () => page });
  const result = await browser.searchOneDrive('forecast');
  assert.equal(result.source, 'onedrive_browser');
  assert.equal(result.items[0].text, 'OneDrive: forecast');
});

test('search completion rejects unrelated or untrusted responses', () => {
  const response = (url, body = '') => ({ url: () => url, request: () => ({ postData: () => body }) });
  assert.equal(isSearchResponse(response('https://substrate.office.com/search/api/v1/query', '{"query":"budget"}'), 'budget'), true);
  assert.equal(isSearchResponse(response('https://evil.test/search', 'budget'), 'budget'), false);
  assert.equal(isSearchResponse(response('https://substrate.office.com/search', 'other'), 'budget'), false);
  assert.equal(isSearchResponse(response('https://login.microsoftonline.com/search', 'budget'), 'budget'), false);
});
