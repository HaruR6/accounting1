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

async function fixture(t, viewport = { width: 1280, height: 900 }, initialize = true) {
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
  if (initialize) {
    const base = `http://127.0.0.1:${port}`;
    const bootstrap = await (await context.request.post(base + '/rpc', { data: { method: 'getBootstrap' } })).json();
    const balances = Object.fromEntries(bootstrap.data.accounts.map(a => [a.account_id, 0]));
    const initialized = await (await context.request.post(base + '/rpc', { data: { method: 'initializeAssets', argument: { request_id: 'browser-opening01', date: '2024-01-01', balances } } })).json();
    assert.equal(initialized.ok, true);
  }
  const errors = [];
  page.on('pageerror', error => errors.push(error.message));
  page.on('dialog', dialog => { if (dialog.type() === 'beforeunload') dialog.accept(); else { errors.push('Unexpected browser dialog: ' + dialog.type()); dialog.dismiss(); } });
  await page.goto(`http://127.0.0.1:${port}/`);
  await page.locator(initialize ? '#entry' : '#assets-initial-form').waitFor({ state: 'visible' });
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

async function assets(page, date = '2026-10-13') {
  await go(page, 'assets'); await page.locator('#assets-date').fill(date); await page.locator('#assets-date').dispatchEvent('change');
  await page.waitForFunction(() => document.querySelector('#assets-content').getAttribute('aria-busy') === 'false' && document.querySelector('#total-assets').textContent !== '—');
}
async function initialBalances(page, date = '2026-01-01') {
  await page.locator('#initial-date').fill(date);
  await page.locator('#initial-asset-cash').fill('10000'); await page.locator('#initial-asset-yucho').fill('100000');
  await page.locator('#initial-asset-financial').fill('200000'); await page.locator('#initial-debt-jal-card').fill('2000');
  await page.locator('#save-initial').click(); await page.locator('#assets-initial-form').waitFor({ state: 'hidden' });
  await page.waitForFunction(() => document.querySelector('#assets-content').getAttribute('aria-busy') === 'false');
}
test('first use: opening assets/debts, auto cash/card reflection, repayment and dated valuation history', { timeout: 60000 }, async t => {
  const { page } = await fixture(t, { width: 1280, height: 900 }, false);
  assert.equal(await page.title(), 'セミナー会計　早瀬ユウカ監修　家計簿アプリ');
  await initialBalances(page); await assets(page);
  assert.equal(await page.locator('#total-assets').textContent(), '310,000円'); assert.equal(await page.locator('#total-liabilities').textContent(), '2,000円'); assert.equal(await page.locator('#total-net-assets').textContent(), '308,000円');
  await go(page, 'entry'); await entry(page, { amount: '1000', date: '2026-01-02' }); await save(page);
  await entry(page, { amount: '500', date: '2026-01-02' }); await page.locator('#entry-payment').selectOption({ label: 'クレジットカード（JALカード）' }); await save(page);
  await assets(page); assert.equal(await page.locator('#total-assets').textContent(), '309,000円'); assert.equal(await page.locator('#total-liabilities').textContent(), '2,500円'); assert.equal(await page.locator('#total-net-assets').textContent(), '306,500円');
  await page.locator('#open-movement').click(); await page.locator('#movement-date').fill('2026-01-03'); await page.locator('#movement-from').selectOption('asset-yucho'); await page.locator('#movement-to').selectOption('debt-jal-card'); await page.locator('#movement-amount').fill('2500'); await page.locator('#save-movement').click();
  await page.locator('#movement-dialog').waitFor({ state: 'hidden' }); await page.waitForFunction(() => document.querySelector('#total-liabilities').textContent === '0円');
  assert.equal(await page.locator('#total-assets').textContent(), '306,500円'); assert.equal(await page.locator('#total-net-assets').textContent(), '306,500円');
  await page.locator('[data-account="asset-financial"] button').click(); await page.locator('#balance-date').fill('2026-03-31'); await page.locator('#balance-amount').fill('230000'); await page.locator('#save-balance').click();
  await page.locator('#balance-dialog').waitFor({ state: 'hidden' }); await page.waitForFunction(() => document.querySelector('#total-net-assets').textContent === '336,500円');
  await assets(page, '2026-02-28'); assert.match(await page.locator('[data-account="asset-financial"]').textContent(), /200,000円/);
  await assets(page); assert.match(await page.locator('[data-account="asset-financial"]').textContent(), /230,000円/);
  assert.equal(await page.locator('#assets-trend-table tbody tr').count(), 12);
  await page.screenshot({ path: path.join(screenshots, 'desktop-assets.png'), fullPage: true });
  await go(page, 'dashboard'); await month(page, '2026-01'); assert.equal(await page.locator('#summary-expense').textContent(), '1,500円');
});

test('sequential typing regroups every third digit, editing keeps digit position, UI contains only functional copy', { timeout: 40000 }, async t => {
  const { page } = await fixture(t, { width: 375, height: 812 });
  assert.equal(await page.locator('#entry-note').getAttribute('placeholder'), null);
  const body = await page.locator('body').textContent(); assert.ok(!body.includes('¥')); assert.ok(!body.includes('小さな記録')); assert.ok(!body.includes('DAILY RECORD'));
  await entry(page); const input = page.locator('#entry-amount'); await input.fill('');
  const expected = ['1', '12', '123', '1,234', '12,345', '123,456', '1,234,567', '12,345,678', '123,456,789'];
  for (let i = 0; i < expected.length; i++) { await input.press(String(i + 1)); assert.equal(await input.inputValue(), expected[i]); }
  await input.press('Backspace'); assert.equal(await input.inputValue(), '12,345,678');
  await input.fill('12345'); await input.evaluate(element => element.setSelectionRange(2, 2)); await input.press('9'); assert.equal(await input.inputValue(), '129,345');
  await input.press('8'); assert.equal(await input.inputValue(), '1,298,345');
  await assets(page); assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
  await page.setViewportSize({ width: 320, height: 740 }); assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
  await page.setViewportSize({ width: 375, height: 812 }); await page.screenshot({ path: path.join(screenshots, 'mobile-assets.png'), fullPage: true });
});

test('opening balance response loss survives reload without duplicate snapshots or lost input', { timeout: 40000 }, async t => {
  const { page, context, base } = await fixture(t, { width: 375, height: 812 }, false);
  let request;
  await page.route('**/rpc', async route => {
    const data = route.request().postDataJSON();
    if (data.method === 'initializeAssets') { request = data.argument; await route.fetch(); await route.abort('failed'); }
    else await route.continue();
  });
  await page.locator('#initial-asset-cash').fill(''); await page.locator('#initial-asset-cash').pressSequentially('1234567'); assert.equal(await page.locator('#initial-asset-cash').inputValue(), '1,234,567');
  await page.locator('#save-initial').click(); await page.locator('#initial-pending').waitFor({ state: 'visible' });
  assert.equal(await page.locator('#initial-asset-cash').isDisabled(), true);
  await page.unroute('**/rpc'); await page.reload(); await page.locator('#initial-pending').waitFor({ state: 'visible' });
  assert.equal(await page.locator('#initial-asset-cash').inputValue(), '1,234,567');
  let retry; page.on('request', r => { if (r.url().endsWith('/rpc')) { const d = r.postDataJSON(); if (d.method === 'initializeAssets') retry = d.argument; } });
  await page.locator('#save-initial').click(); await page.locator('#assets-initial-form').waitFor({ state: 'hidden' });
  assert.equal(retry.request_id, request.request_id); assert.equal(await page.locator('#total-assets').textContent(), '1,234,567円');
  const result = await (await context.request.post(base + '/rpc', { data: { method: 'getAssets' } })).json(); assert.equal(result.data.assets, 1234567);
});

test('asset movement deletion restores balances and mapping changes are saved through the asset tab', { timeout: 60000 }, async t => {
  const { page } = await fixture(t, { width: 1280, height: 900 }, false); await initialBalances(page); await assets(page);
  await page.locator('#open-movement').click(); await page.locator('#movement-date').fill('2026-01-03'); await page.locator('#movement-from').selectOption('asset-cash'); await page.locator('#movement-to').selectOption('asset-suica'); await page.locator('#movement-amount').fill('3000'); await page.locator('#save-movement').click();
  await page.locator('#movement-dialog').waitFor({ state: 'hidden' }); await page.locator('[data-account="asset-suica"]').filter({ hasText: '3,000円' }).waitFor();
  assert.equal(await page.locator('#total-assets').textContent(), '310,000円');
  await page.locator('#asset-movements button').click(); await page.locator('#confirm-delete').click(); await page.locator('#delete-dialog').waitFor({ state: 'hidden' }); await page.locator('[data-account="asset-suica"]').filter({ hasText: '0円' }).waitFor();
  await page.locator('#asset-links li').filter({ hasText: '現金' }).locator('button').click(); await page.locator('#link-account').selectOption('asset-paypay'); await page.locator('#save-link').click();
  await page.locator('#link-dialog').waitFor({ state: 'hidden' }); await page.locator('#asset-links li').filter({ hasText: '現金' }).filter({ hasText: 'PayPay' }).waitFor();
  await go(page, 'entry'); await entry(page, { amount: '500', date: '2026-01-04' }); assert.match(await page.locator('#entry-link-help').textContent(), /PayPay/); await save(page);
  await assets(page); assert.match(await page.locator('[data-account="asset-cash"]').textContent(), /10,000円/); assert.match(await page.locator('[data-account="asset-paypay"]').textContent(), /-500円/);
});

const saveMessages = {
  expense: ['先生！使いすぎです！', '先生、こんなに何に使ったんですか！', 'はあ、私がちゃんと見ていないがばっかりに、、、'],
  income: ['先生！よくがんばりましたね！', 'いつもありがとうございます、先生', 'こんなにいっぱい！でも、ちゃんと考えて使うんですよ？']
};
async function chooseSaveMessage(page, index) {
  await page.evaluate(value => { Math.random = () => value; }, [0, 0.5, 0.999999][index]);
}
async function assertSaveMessage(page, type, index) {
  const card = page.locator('#save-message-card');
  await card.waitFor({ state: 'visible' });
  assert.match(await card.textContent(), /保存しました/);
  assert.equal(await page.locator('#save-message-text').textContent(), saveMessages[type][index]);
  assert.equal(await card.evaluate((element, type) => element.classList.contains(type), type), true);
  assert.equal(await card.evaluate((element, other) => element.classList.contains(other), type === 'expense' ? 'income' : 'expense'), false);
}

test('save message: inclusive thresholds and each of the three exact texts for expense and income', { timeout: 60000 }, async t => {
  const { page } = await fixture(t);
  const card = page.locator('#save-message-card');
  assert.equal(await card.isVisible(), false);
  for (const [type, amount] of [['expense', '9999'], ['income', '29999']]) {
    await entry(page, { type, amount }); await save(page);
    assert.equal(await card.isVisible(), false, `${type} below threshold`);
  }
  for (const type of ['expense', 'income']) {
    const threshold = type === 'expense' ? 10000 : 30000;
    for (let index = 0; index < 3; index++) {
      await chooseSaveMessage(page, index);
      await entry(page, { type, amount: String(threshold + index) }); await save(page);
      await assertSaveMessage(page, type, index);
    }
  }
  // A smaller new registration must dismiss the previous qualifying card.
  await entry(page, { type: 'income', amount: '10000' }); await save(page);
  assert.equal(await card.isVisible(), false);
});

test('save message: no premature or failure popup; lost-response retry displays after confirmed success without duplicate entry', { timeout: 40000 }, async t => {
  const { page } = await fixture(t);
  const card = page.locator('#save-message-card');
  await entry(page, { type: 'income', amount: '30000', note: '次の保存前のメッセージ' }); await save(page);
  await card.waitFor({ state: 'visible' });
  const requests = [];
  let releaseFailure, requestStarted;
  const failureGate = new Promise(resolve => { releaseFailure = resolve; });
  const started = new Promise(resolve => { requestStarted = resolve; });
  await page.route('**/rpc', async route => {
    const data = route.request().postDataJSON();
    if (data.method !== 'createTransaction') return route.continue();
    requests.push(data.argument);
    requestStarted();
    await failureGate;
    await route.fulfill({ json: { ok: false, error: { code: 'VALIDATION', message: 'メッセージ検証用の保存失敗', retryable: false } } });
  });
  await chooseSaveMessage(page, 1);
  await entry(page, { amount: '10000' }); await page.locator('#save-transaction').click(); await started;
  assert.equal(await page.locator('#save-transaction').isDisabled(), true);
  assert.equal(await card.isVisible(), false, 'The previous card disappears as soon as the next save starts');
  releaseFailure();
  await page.locator('#entry-error').filter({ hasText: 'メッセージ検証用の保存失敗' }).waitFor();
  assert.equal(await card.isVisible(), false, 'No card on definitive failure');
  assert.equal(await page.locator('#entry-amount').inputValue(), '10,000');
  await page.unroute('**/rpc');
  await page.route('**/rpc', async route => {
    const data = route.request().postDataJSON();
    if (data.method !== 'createTransaction') return route.continue();
    requests.push(data.argument);
    await route.fetch();
    await route.abort('failed');
  });
  await page.locator('#save-transaction').click();
  await page.locator('#pending-message').waitFor({ state: 'visible' });
  assert.equal(await card.isVisible(), false, 'No card before the saved result is confirmed');
  await page.unroute('**/rpc');
  page.on('request', request => {
    if (request.url().endsWith('/rpc')) { const data = request.postDataJSON(); if (data.method === 'createTransaction') requests.push(data.argument); }
  });
  await save(page); await assertSaveMessage(page, 'expense', 1);
  assert.equal(requests.length, 3);
  assert.equal(new Set(requests.map(request => request.request_id)).size, 1);
  await history(page); assert.equal(await page.locator('#history-count').textContent(), '2件の記録');
});

test('save message: six-second lifetime, fresh timeout on replacement, manual close and continued entry focus', { timeout: 40000 }, async t => {
  const { page } = await fixture(t);
  await page.emulateMedia({ reducedMotion: 'reduce' });
  await page.clock.install({ time: new Date('2026-10-13T00:00:00Z') });
  await page.clock.pauseAt(new Date('2026-10-13T00:00:01Z'));
  const card = page.locator('#save-message-card');
  await chooseSaveMessage(page, 0);
  await entry(page, { amount: '10000' }); await save(page); await assertSaveMessage(page, 'expense', 0);
  assert.equal(await page.locator('dialog[open]').count(), 0);
  assert.equal(await card.evaluate(element => element.contains(document.activeElement)), false, 'Showing the card must not steal focus');
  await page.locator('#entry-amount').focus();
  assert.equal(await page.locator('#entry-amount').evaluate(element => element === document.activeElement), true);
  await page.clock.runFor(3000); assert.equal(await card.isVisible(), true);
  await chooseSaveMessage(page, 2);
  await entry(page, { type: 'income', amount: '30000' });
  // Clear the previous status so save() waits for this registration's result.
  await page.locator('#status').evaluate(element => { element.textContent = ''; });
  await save(page); await assertSaveMessage(page, 'income', 2);
  await page.locator('#entry-note').focus();
  await page.clock.runFor(3000); assert.equal(await card.isVisible(), true, 'The old card timer must not close the replacement');
  await page.clock.runFor(2999); assert.equal(await card.isVisible(), true);
  await page.clock.runFor(1); assert.equal(await card.isVisible(), false, 'The replacement closes at six seconds');
  assert.equal(await page.locator('#entry-note').evaluate(element => element === document.activeElement), true, 'Automatic dismissal keeps input focus');
  await entry(page, { amount: '10000' });
  await page.locator('#status').evaluate(element => { element.textContent = ''; });
  await save(page); await card.waitFor({ state: 'visible' });
  await page.locator('#close-save-message').click(); assert.equal(await card.isVisible(), false);
  assert.equal(await card.evaluate(element => element.contains(document.activeElement)), false, 'Manual dismissal does not leave focus in a hidden card');
});

test('save message: mobile card stays within the screen above the navigation and leaves navigation usable', { timeout: 40000 }, async t => {
  const { page } = await fixture(t, { width: 375, height: 812 });
  await chooseSaveMessage(page, 2);
  await entry(page, { amount: '10000' }); await save(page); await assertSaveMessage(page, 'expense', 2);
  for (const viewport of [{ width: 375, height: 812 }, { width: 320, height: 640 }]) {
    await page.setViewportSize(viewport);
    const card = await page.locator('#save-message-card').boundingBox();
    const nav = await page.locator('.main-nav').boundingBox();
    assert.ok(card && nav);
    assert.ok(card.x >= 0 && card.x + card.width <= viewport.width, 'Card fits horizontally');
    assert.ok(card.y >= 0 && card.y + card.height < nav.y, 'Card remains above the menu without overlap');
    assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth), true);
  }
  await go(page, 'history'); await page.locator('#history').waitFor({ state: 'visible' });
  await go(page, 'entry');
  await page.screenshot({ path: path.join(screenshots, 'mobile-save-message.png'), fullPage: true });
});

test('save message: editing, deletion, settings and asset balance saves do not display registration messages', { timeout: 60000 }, async t => {
  const { page } = await fixture(t, { width: 1280, height: 900 }, false);
  const card = page.locator('#save-message-card');
  await initialBalances(page); assert.equal(await card.isVisible(), false, 'Opening balances are not income');
  await go(page, 'entry'); await entry(page, { amount: '10000', note: 'メッセージ除外の検証' }); await save(page);
  await card.waitFor({ state: 'visible' }); await page.locator('#close-save-message').click();
  await history(page);
  await page.locator('.transaction-row').filter({ hasText: 'メッセージ除外の検証' }).locator('button').filter({ hasText: '編集' }).click();
  await page.locator('#entry-amount').fill('20000'); await save(page);
  await page.locator('#history').waitFor({ state: 'visible' }); assert.equal(await card.isVisible(), false, 'No card for a qualifying edited amount');
  await page.locator('.transaction-row').filter({ hasText: 'メッセージ除外の検証' }).locator('button').filter({ hasText: '削除' }).click();
  await page.locator('#confirm-delete').click(); await page.locator('#status').filter({ hasText: '記録を削除しました' }).waitFor();
  assert.equal(await card.isVisible(), false, 'No card for deletion');
  await go(page, 'settings'); await page.locator('#category-settings li').filter({ hasText: '食費' }).locator('button').click();
  await page.locator('#master-name').fill('食事'); await page.locator('#save-master').click(); await page.locator('#master-dialog').waitFor({ state: 'hidden' });
  assert.equal(await card.isVisible(), false, 'No card for settings');
  await assets(page); await page.locator('[data-account="asset-financial"] button').click();
  await page.locator('#balance-amount').fill('50000'); await page.locator('#save-balance').click(); await page.locator('#balance-dialog').waitFor({ state: 'hidden' });
  assert.equal(await card.isVisible(), false, 'A valuation change is not income');
});
