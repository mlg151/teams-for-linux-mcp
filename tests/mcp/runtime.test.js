'use strict';

const { test } = require('node:test');
const assert = require('node:assert/strict');
const { TeamsRuntime, isTeamsUrl } = require('../../app/mcp/runtime');
const { applyMcpConfig } = require('../../app/mcp/config');

test('only HTTPS Teams origins are trusted for session operations', () => {
  assert.equal(isTeamsUrl('https://teams.cloud.microsoft/v2/'), true);
  assert.equal(isTeamsUrl('https://login.microsoftonline.com/'), false);
  assert.equal(isTeamsUrl('https://teams.microsoft.com.evil.test/'), false);
  assert.equal(isTeamsUrl('http://teams.microsoft.com/'), false);
});

test('headless config is opt-in and preserves the existing partition/auth configuration', () => {
  const config = { partition: 'persist:existing', auth: { intune: { enabled: true } }, trayIconEnabled: true };
  const ordinary = structuredClone(config);
  applyMcpConfig(config, undefined);
  assert.deepEqual(config, ordinary);
  applyMcpConfig(config, 'headless');
  assert.equal(config.minimized, true);
  assert.equal(config.graphApi.enabled, true);
  assert.equal(config.trayIconEnabled, false);
  assert.equal(config.partition, ordinary.partition);
  assert.deepEqual(config.auth, ordinary.auth);
});

test('runtime serializes UI operations and recovers after a failed operation', async () => {
  const runtime = new TeamsRuntime();
  let release;
  const blocked = new Promise(resolve => { release = resolve; });
  const order = [];
  const first = runtime.run(async () => { order.push('first'); await blocked; throw new Error('failed'); });
  const rejection = assert.rejects(first, /failed/);
  const second = runtime.run(() => { order.push('second'); return 'ok'; });
  await new Promise(resolve => setImmediate(resolve));
  assert.deepEqual(order, ['first']);
  release();
  await rejection;
  assert.equal(await second, 'ok');
  assert.deepEqual(order, ['first', 'second']);
  await runtime.close();
  await assert.rejects(runtime.run(() => 'too late'), { code: 'shutting_down' });
});

test('cancelling a queued operation prevents a delayed write', async () => {
  const runtime = new TeamsRuntime();
  const controller = new AbortController();
  let release;
  let writes = 0;
  const first = runtime.run(() => new Promise(resolve => { release = resolve; }));
  const queued = runtime.run(() => { writes++; }, { signal: controller.signal });
  const rejected = assert.rejects(queued, { code: 'cancelled' });
  await new Promise(resolve => setImmediate(resolve));
  controller.abort();
  release();
  await first;
  await rejected;
  assert.equal(writes, 0);
  await runtime.close();
});
