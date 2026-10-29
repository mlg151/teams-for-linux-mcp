'use strict';

// Browser plugin not available. Real Electron renderers exercise the private
// browser link against an intercepted Microsoft-origin login/MFA fixture.
const { before, after, test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const { _electron } = require('playwright-core');
const { TeamsRuntime } = require('../../app/mcp/runtime');

let application;
let target;
let viewer;
let runtime;
let fixtureDir;
const errors = [];
const fixture = `<!doctype html><html><head><title>Microsoft sign-in fixture</title><style>
body{font:18px system-ui;background:#f8f9fc;padding:40px;color:#24324b}main{background:white;border:1px solid #ccd6e5;border-radius:8px;padding:24px;width:360px}label{display:block;margin:15px 0 5px}input{display:block;padding:10px;width:300px;font:18px system-ui}button{margin-top:18px;padding:10px 22px;background:#2851a0;color:white;border:0;border-radius:4px;font-size:16px}
</style></head><body><main><h1>Microsoft sign-in</h1><form id="credentials"><label for="email">Email</label><input id="email" type="email" autocomplete="off"><label for="password">Password</label><input id="password" type="password" autocomplete="off"><button>Next</button></form><form id="mfa" hidden><h2>Verify your identity</h2><p>Enter the verification code from your authenticator app.</p><label for="code">Verification code</label><input id="code" inputmode="numeric" autocomplete="off"><button>Verify</button></form></main><script>
document.querySelector('#credentials').onsubmit=event=>{event.preventDefault();if(document.querySelector('#email').value==='fixture@example.test'&&document.querySelector('#password').value==='fixture-secret'){document.querySelector('#credentials').hidden=true;document.querySelector('#mfa').hidden=false;}};
document.querySelector('#mfa').onsubmit=event=>{event.preventDefault();if(document.querySelector('#code').value==='123456'){document.body.dataset.authenticated='true';document.querySelector('main').textContent='Teams is ready';}};
</script></body></html>`;

before(async () => {
  fixtureDir = fs.mkdtempSync(path.join(os.tmpdir(), 'teams-mcp-login-ui-'));
  const main = path.join(fixtureDir, 'main.js');
  fs.writeFileSync(main, `const {app,BrowserWindow}=require('electron');app.setPath('userData',${JSON.stringify(path.join(fixtureDir, 'profile'))});let window;app.whenReady().then(()=>{window=new BrowserWindow({width:900,height:720,show:false,webPreferences:{backgroundThrottling:false}});window.loadURL('about:blank')});app.on('window-all-closed',()=>app.quit());`);
  const env = { ...process.env };
  delete env.ELECTRON_RUN_AS_NODE;
  application = await _electron.launch({ executablePath: require('electron'), args: [main], env, timeout: 30000 });
  target = await application.firstWindow();
  target.setDefaultTimeout(10000);
  await target.route('https://login.microsoftonline.com/mcp-fixture', route => route.fulfill({ contentType: 'text/html', body: fixture }));
  runtime = new TeamsRuntime({ loginOptions: { ttlMs: 60000 }, virtualDisplay: true });
  runtime.application = application;
  runtime.status = async () => ({ authenticated: await target.evaluate(() => document.body.dataset.authenticated === 'true') });
  runtime.getLoginPage = async () => target;
}, { timeout: 45000 });

after(async () => {
  if (runtime) await runtime.close();
  else await application?.close();
  if (fixtureDir) fs.rmSync(fixtureDir, { recursive: true, force: true });
});

async function openLink(width) {
  await target.goto('https://login.microsoftonline.com/mcp-fixture');
  const login = await runtime.run(() => runtime.beginLogin());
  if (!viewer) {
    const next = application.waitForEvent('window');
    await application.evaluate(({ BrowserWindow }, { url, width }) => {
      globalThis.loginTestViewer = new BrowserWindow({ width, height: 844, show: false, webPreferences: { backgroundThrottling: false } });
      globalThis.loginTestViewer.loadURL(url);
    }, { url: login.url, width });
    viewer = await next;
    viewer.setDefaultTimeout(10000);
    viewer.on('pageerror', error => errors.push(error.message));
  } else {
    const window = await application.browserWindow(viewer);
    await window.evaluate((window, width) => window.setSize(width, 844), width);
    await window.dispose();
    await viewer.goto(login.url);
  }
  await viewer.locator('#screen').waitFor({ state: 'visible' });
  await viewer.waitForFunction(() => document.querySelector('#screen').naturalWidth > 0);
  return login;
}

async function enter(selector, value) {
  const remote = await target.locator(selector).boundingBox();
  const image = await viewer.locator('#screen').boundingBox();
  const size = await target.evaluate(() => ({ width: innerWidth, height: innerHeight }));
  await viewer.locator('#screen').click({ position: {
    x: (remote.x + remote.width / 2) * image.width / size.width,
    y: (remote.y + remote.height / 2) * image.height / size.height,
  } });
  try { await target.waitForFunction(selector => document.activeElement === document.querySelector(selector), selector); }
  catch {
    const active = await target.evaluate(() => document.activeElement.id || document.activeElement.tagName);
    throw new Error(`Remote login field did not focus (${selector}; active ${active}). ${await viewer.locator('#error').innerText()}`);
  }
  await viewer.locator('#typing').fill(value);
  await target.waitForFunction(({ selector, value }) => document.querySelector(selector).value === value, { selector, value });
}

async function completeLogin() {
  await enter('#email', 'fixture@example.test');
  await enter('#password', 'fixture-secret');
  assert.equal((await viewer.locator('body').innerText()).includes('fixture-secret'), false);
  await viewer.locator('[data-key="Enter"]').click();
  await target.locator('#code').waitFor({ state: 'visible' });
  await viewer.locator('#refresh').click();
  await enter('#code', '123456');
  await viewer.locator('[data-key="Enter"]').click();
  await viewer.locator('#done').waitFor({ state: 'visible' });
  assert.equal((await runtime.status()).authenticated, true);
  assert.equal(runtime.loginPortal.closed, true);
  assert.equal(runtime.loginPortal.token, null);
  assert.equal(await viewer.locator('#screen').getAttribute('src'), null);
}

async function screenshot(file) {
  fs.writeFileSync(file, Buffer.from(await runtime.captureLoginScreen(viewer), 'base64'));
}

test('MCP browser link completes username, password, and MFA in the same headless session', { timeout: 45000 }, async () => {
  const login = await openLink(1120);
  assert.equal(await viewer.title(), 'Microsoft sign-in · Teams MCP');
  assert.equal(new URL(viewer.url()).origin, new URL(login.url).origin);
  assert.match(await viewer.locator('#status').innerText(), /Connected/);
  assert.equal(await viewer.locator('#origin').innerText(), 'https://login.microsoftonline.com');
  await screenshot('/tmp/teams-mcp-login-screen.png');
  await completeLogin();
  await screenshot('/tmp/teams-mcp-login-complete.png');
  assert.deepEqual(errors, []);
  assert.equal((await runtime.beginLogin()).url, undefined);
});

test('browser login controls work at a mobile width without horizontal overflow', { timeout: 45000 }, async () => {
  await openLink(390);
  assert.equal(await viewer.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
  await completeLogin();
  await screenshot('/tmp/teams-mcp-login-mobile.png');
  assert.deepEqual(errors, []);
});
