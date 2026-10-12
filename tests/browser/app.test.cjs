const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');
const { spawn } = require('node:child_process');
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');
const root = path.resolve(__dirname, '../..');
const screenshots = path.join(root, '.local', 'screenshots');
let browser;

before(async () => {
  fs.mkdirSync(screenshots, { recursive: true });
  const executablePath = process.env.CHROMIUM_PATH || ['/usr/bin/chromium', '/usr/bin/chromium-browser', '/usr/bin/google-chrome'].find(p => fs.existsSync(p));
  browser = await chromium.launch({ ...(executablePath ? { executablePath } : {}), headless: true });
});
after(async () => { await browser?.close(); });

async function fixture(t, viewport = { width: 1280, height: 900 }) {
  const server = spawn(process.execPath, ['scripts/dev.cjs'], { cwd: root, env: { ...process.env, PORT: '0' }, stdio: ['ignore', 'pipe', 'pipe'] });
  t.after(() => server.kill());
  const port = await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('Development server did not start')), 10000);
    server.stdout.on('data', chunk => { const match = String(chunk).match(/on port (\d+)/); if (match) { clearTimeout(timer); resolve(Number(match[1])); } });
    server.on('error', error => { clearTimeout(timer); reject(error); });
    server.on('exit', code => { if (code) { clearTimeout(timer); reject(new Error('Server exited: ' + code)); } });
  });
  const context = await browser.newContext({ viewport, timezoneId: 'America/Los_Angeles' });
  t.after(() => context.close());
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  page.on('dialog', dialog => { if (dialog.type() === 'beforeunload') dialog.accept(); else { errors.push('Unexpected browser dialog: ' + dialog.type()); dialog.dismiss(); } });
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.locator('#entry').waitFor({ state: 'visible' });
  t.after(() => assert.deepEqual(errors, [], 'Browser JavaScript errors'));
  return { page, context, base: `http://127.0.0.1:${port}` };
}
async function entry(page, { amount = '680', note = '昼食', date = '2026-10-13', type = 'expense' } = {}) {
  await page.locator(`[data-type="${type}"]`).click();
  await page.locator('#entry-date').fill(date);
  await page.locator('#entry-payment').selectOption({ label: '現金' });
  await page.locator('#entry-category').selectOption({ label: type === 'income' ? '給与・アルバイト' : '食費' });
  await page.locator('#entry-amount').fill(amount);
  await page.locator('#entry-note').fill(note);
}
async function save(page) {
  await page.locator('#save-transaction').click();
  await page.locator('#status').filter({ hasText: /記録しました|記録を更新しました/ }).waitFor();
  await page.waitForFunction(() => !document.querySelector('#entry-fields').disabled);
}
async function go(page, view) {
  await page.locator(`[data-view="${view}"]`).click();
  await page.locator('#' + view).waitFor({ state: 'visible' });
}
async function month(page, value = '2026-10') {
  await page.locator('#dashboard-month').fill(value);
  await page.locator('#dashboard-month').dispatchEvent('change');
  await page.waitForFunction(() => document.querySelector('#dashboard-content').getAttribute('aria-busy') === 'false' && document.querySelector('#summary-expense').textContent !== '—');
}
async function history(page) {
  await go(page, 'history');
  await page.locator('#history-month').fill('2026-10');
  await page.locator('#history-filters button').click();
  await page.waitForFunction(() => document.querySelector('#history-list').getAttribute('aria-busy') === 'false');
}

test('desktop: expense/income entry, summaries, safe text, editing and logical deletion', { timeout: 60000 }, async t => {
  const { page } = await fixture(t);
  await entry(page); await save(page);
  assert.equal(await page.locator('#entry-payment').inputValue(), ''); assert.equal(await page.locator('#entry-category').inputValue(), ''); assert.equal(await page.locator('#entry-amount').inputValue(), ''); assert.equal(await page.locator('#entry-note').inputValue(), '');
  const today = await page.evaluate(() => new Intl.DateTimeFormat('sv-SE', { timeZone: 'Asia/Tokyo' }).format(new Date()));
  assert.equal(await page.locator('#entry-date').inputValue(), today);
  await entry(page, { type: 'income', amount: '200000', note: '給与' });
  assert.equal(await page.locator('#payment-label').textContent(), '入金先');
  assert.ok(!(await page.locator('#entry-category').textContent()).includes('食費'));
  assert.equal(await page.locator('#entry-amount').inputValue(), '200,000'); await save(page);
  await entry(page, { amount: '320', note: '<img src=x onerror=alert(1)>\n=1+1' }); await save(page);
  await go(page, 'dashboard'); await month(page);
  assert.equal(await page.locator('#summary-income').textContent(), '200,000円'); assert.equal(await page.locator('#summary-expense').textContent(), '1,000円'); assert.equal(await page.locator('#summary-balance').textContent(), '199,000円');
  assert.match(await page.locator('#category-legend').textContent(), /食費.*1,000円.*100.0%/);
  assert.equal(await page.locator('#trend-table tbody tr').count(), 12);
  await page.screenshot({ path: path.join(screenshots, 'desktop-dashboard.png'), fullPage: true });
  await history(page);
  const row = page.locator('.transaction-row').filter({ hasText: '昼食' }); await row.locator('button').filter({ hasText: '編集' }).click();
  await page.locator('#entry-amount').fill('900'); await save(page);
  await page.locator('#history').waitFor({ state: 'visible' });
  await page.locator('.transaction-row').filter({ hasText: '昼食' }).filter({ hasText: '900円' }).waitFor();
  assert.equal(await page.locator('#history-list img').count(), 0);
  await go(page, 'dashboard'); await month(page); assert.equal(await page.locator('#summary-expense').textContent(), '1,220円');
  await history(page);
  await page.locator('.transaction-row').filter({ hasText: '昼食' }).locator('button').filter({ hasText: '削除' }).click();
  await page.locator('#delete-dialog').waitFor({ state: 'visible' }); await page.locator('#confirm-delete').click();
  await page.locator('#status').filter({ hasText: '記録を削除しました' }).waitFor();
  await page.waitForFunction(() => document.querySelector('#history-count').textContent === '2件の記録');
  await go(page, 'dashboard'); await month(page); assert.equal(await page.locator('#summary-expense').textContent(), '320円');
});

test('mobile: required fields, invalid amounts, leap date and responsive layout', { timeout: 40000 }, async t => {
  const { page, context, base } = await fixture(t, { width: 375, height: 812 });
  assert.equal(await page.locator('#save-transaction').isDisabled(), true);
  await entry(page, { date: '2024-02-29' }); assert.equal(await page.locator('#save-transaction').isEnabled(), true);
  for (const value of ['0', '-1', '1.1', 'abc', '1,2', '9007199254740992']) {
    await page.locator('#entry-amount').fill(value); assert.equal(await page.locator('#save-transaction').isDisabled(), true, value);
  }
  await page.locator('#entry-amount').fill('12345'); assert.equal(await page.locator('#entry-amount').inputValue(), '12,345');
  await page.locator('#entry-category').selectOption(''); assert.equal(await page.locator('#save-transaction').isDisabled(), true);
  await page.locator('#entry-category').selectOption({ label: '食費' });
  assert.equal(await page.locator('#entry-date').inputValue(), '2024-02-29');
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
  await save(page);
  await go(page, 'dashboard'); await month(page, '2024-02');
  assert.equal(await page.locator('#summary-expense').textContent(), '12,345円');
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
  await page.setViewportSize({ width: 320, height: 740 });
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
  await page.setViewportSize({ width: 375, height: 812 });
  await page.screenshot({ path: path.join(screenshots, 'mobile-dashboard.png'), fullPage: true });
  const entryPage = await context.newPage();
  await entryPage.goto(base); await entryPage.locator('#entry').waitFor({ state: 'visible' });
  await entry(entryPage, { amount: '12345', date: '2024-02-29' });
  await entryPage.screenshot({ path: path.join(screenshots, 'mobile-entry.png'), fullPage: true });
});

test('confirmed save failure preserves input, enables correction and reuses request ID on unchanged retry', { timeout: 40000 }, async t => {
  const { page } = await fixture(t);
  let rejectedRequest;
  await page.route('**/rpc', async route => {
    const request = route.request().postDataJSON();
    if (request.method === 'createTransaction') {
      rejectedRequest = request.argument;
      await route.fulfill({ json: { ok: false, error: { code: 'VALIDATION', message: '検証用の保存失敗', retryable: false } } });
    } else await route.continue();
  });
  await entry(page); await page.locator('#save-transaction').click();
  await page.locator('#entry-error').filter({ hasText: '検証用の保存失敗' }).waitFor();
  assert.equal(await page.locator('#entry-note').inputValue(), '昼食'); assert.equal(await page.locator('#entry-amount').inputValue(), '680'); assert.equal(await page.locator('#entry-date').isDisabled(), false);
  await page.unroute('**/rpc');
  let retriedRequest;
  page.on('request', request => { if (request.url().endsWith('/rpc')) { const data = request.postDataJSON(); if (data.method === 'createTransaction') retriedRequest = data.argument; } });
  await save(page); assert.equal(retriedRequest.request_id, rejectedRequest.request_id);
  await history(page); assert.equal(await page.locator('#history-count').textContent(), '1件の記録');
});

test('lost response after successful save: same request survives reload, prevents duplicate row', { timeout: 40000 }, async t => {
  const { page } = await fixture(t);
  let firstRequest;
  await page.route('**/rpc', async route => {
    const request = route.request().postDataJSON();
    if (request.method === 'createTransaction') { firstRequest = request.argument; await route.fetch(); await route.abort('failed'); }
    else await route.continue();
  });
  await entry(page); await page.locator('#save-transaction').click();
  await page.locator('#pending-message').waitFor({ state: 'visible' });
  assert.equal(await page.locator('#entry-date').isDisabled(), true); assert.equal(await page.locator('#entry-note').inputValue(), '昼食');
  await page.unroute('**/rpc'); await page.reload();
  await page.locator('#pending-message').waitFor({ state: 'visible' }); assert.equal(await page.locator('#entry-amount').inputValue(), '680');
  let retry;
  page.on('request', request => { if (request.url().endsWith('/rpc')) { const data = request.postDataJSON(); if (data.method === 'createTransaction') retry = data.argument; } });
  await save(page); assert.equal(retry.request_id, firstRequest.request_id);
  await history(page); assert.equal(await page.locator('#history-count').textContent(), '1件の記録');
});

test('settings: rename/disable, historical labels, retained inactive edit, reactivation and duplicate names', { timeout: 60000 }, async t => {
  const { page } = await fixture(t);
  await entry(page); await save(page); await go(page, 'settings');
  const food = page.locator('#category-settings li').filter({ hasText: '食費' }); await food.locator('button').click();
  await page.locator('#master-name').fill('食事'); await page.locator('#master-active').uncheck(); await page.locator('#save-master').click();
  await page.locator('#master-dialog').waitFor({ state: 'hidden' }); await page.locator('#category-settings li').filter({ hasText: '食事' }).filter({ hasText: '使用停止中' }).waitFor();
  await go(page, 'entry'); assert.ok(!(await page.locator('#entry-category').textContent()).includes('食事'));
  await history(page); const row = page.locator('.transaction-row').filter({ hasText: '食事' }); await row.locator('button').filter({ hasText: '編集' }).click();
  assert.match(await page.locator('#entry-category').textContent(), /食事（使用停止中）/);
  await page.locator('#entry-note').fill('停止後に備考を編集'); await save(page);
  await go(page, 'settings'); await page.locator('#category-settings li').filter({ hasText: '食事' }).locator('button').click();
  await page.locator('#master-active').check(); await page.locator('#master-order').fill('0'); await page.locator('#save-master').click();
  await page.locator('#master-dialog').waitFor({ state: 'hidden' }); await page.locator('#category-settings li').first().filter({ hasText: '食事' }).waitFor();
  await page.locator('[data-add="category"]').click(); await page.locator('#master-name').fill('食事'); await page.locator('#save-master').click();
  await page.locator('#master-error').filter({ hasText: '同じ名称' }).waitFor();
  await page.locator('#master-name').fill('ペット'); await page.locator('#save-master').click(); await page.locator('#master-dialog').waitFor({ state: 'hidden' });
  await page.locator('#category-settings li').filter({ hasText: 'ペット' }).waitFor();
  await page.screenshot({ path: path.join(screenshots, 'desktop-settings.png'), fullPage: true });
  await go(page, 'entry'); assert.match(await page.locator('#entry-category').textContent(), /食事/); assert.match(await page.locator('#entry-category').textContent(), /ペット/);
});

test('empty month renders zero totals, empty category state, and correct year-crossing trend', { timeout: 40000 }, async t => {
  const { page } = await fixture(t);
  await go(page, 'dashboard'); await month(page, '2026-01');
  assert.equal(await page.locator('#summary-income').textContent(), '0円'); assert.equal(await page.locator('#summary-expense').textContent(), '0円'); assert.equal(await page.locator('#summary-balance').textContent(), '0円');
  assert.match(await page.locator('#category-chart').textContent(), /データがありません/);
  const rows = page.locator('#trend-table tbody tr'); assert.equal(await rows.count(), 12); assert.match(await rows.first().textContent(), /^2025-02/); assert.match(await rows.last().textContent(), /^2026-01/);
  await history(page); assert.match(await page.locator('#history-list').textContent(), /記録はありません/);
});
