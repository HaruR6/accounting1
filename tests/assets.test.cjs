const test = require('node:test');
const assert = require('node:assert/strict');
const { createRuntime } = require('../scripts/gas-runtime.cjs');
function ok(result) { assert.equal(result.ok, true, result.error?.message); return result.data; }
function fail(result, code) { assert.equal(result.ok, false); assert.equal(result.error.code, code); }
function opening(r, balances = {}, date = '2026-01-01') {
  const accounts = ok(r.invoke('getBootstrap')).accounts;
  return { request_id: 'opening-request-001', date, balances: Object.fromEntries(accounts.map(a => [a.account_id, balances[a.account_id] || 0])) };
}
function initialized(balances = {}) { const r = createRuntime(); ok(r.invoke('initializeAssets', opening(r, balances))); return r; }
function balances(r, date = '2026-10-13') { return ok(r.invoke('getAssets', { date })); }
function amount(r, id, date) { return balances(r, date).accounts.find(a => a.account_id === id).amount; }
function transaction(r, fields = {}) {
  const bootstrap = ok(r.invoke('getBootstrap'));
  return { request_id: 'asset-transaction-001', date: '2026-01-02', type: 'expense', payment_method_id: 'initial-payment-1', category_id: bootstrap.categories.find(c => c.type === (fields.type || 'expense')).category_id, amount: 1000, note: '', ...fields };
}
function revalue(r, fields = {}) { return { request_id: 'revalue-request-001', date: '2026-03-31', balances: { 'asset-financial': 230000 }, expected_revision: balances(r).revision, ...fields }; }

test('assets start unregistered; opening balances produce assets minus liabilities and survive setup rerun', () => {
  const r = createRuntime(); assert.equal(ok(r.invoke('getBootstrap')).assetsInitialized, false);
  assert.equal(balances(r).known, false); assert.equal(balances(r).accounts[0].amount, null);
  const input = opening(r, { 'asset-cash': 10000, 'asset-yucho': 100000, 'asset-jrebank': 50000, 'asset-financial': 200000, 'asset-suica': 1000, 'asset-paypay': 5000, 'debt-jal-card': 2000, 'debt-jre-card': 3000, 'debt-other': 10000 });
  ok(r.invoke('initializeAssets', input)); const value = balances(r);
  assert.equal(value.assets, 366000); assert.equal(value.liabilities, 15000); assert.equal(value.netAssets, 351000); assert.equal(value.known, true);
  const snapshot = JSON.stringify(r.book().getSheetByName('AssetSnapshots').rows);
  r.invoke('setupDatabase'); assert.equal(JSON.stringify(r.book().getSheetByName('AssetSnapshots').rows), snapshot);
  assert.equal(r.book().getSheetByName('AssetAccounts').getLastRow(), 10); assert.equal(ok(r.invoke('getBootstrap')).assetsInitialized, true);
});

test('expenses, income and card refunds update the correct asset or liability; edits and deletions recalculate', () => {
  const r = initialized({ 'asset-cash': 10000, 'debt-jal-card': 2000 });
  const cash = ok(r.invoke('createTransaction', transaction(r))).transaction;
  assert.equal(cash.asset_account_id, 'asset-cash'); assert.equal(amount(r, 'asset-cash'), 9000);
  ok(r.invoke('createTransaction', transaction(r, { request_id: 'income-request-001', type: 'income', amount: 3000 })));
  ok(r.invoke('createTransaction', transaction(r, { request_id: 'card-expense-0001', payment_method_id: 'initial-payment-2', amount: 500 })));
  ok(r.invoke('createTransaction', transaction(r, { request_id: 'card-refund-00001', payment_method_id: 'initial-payment-2', type: 'income', amount: 100 })));
  assert.equal(amount(r, 'asset-cash'), 12000); assert.equal(amount(r, 'debt-jal-card'), 2400); assert.equal(balances(r).netAssets, 9600);
  const edit = ok(r.invoke('updateTransaction', { ...cash, amount: 1500, expected_updated_at: cash.updated_at })).transaction;
  assert.equal(amount(r, 'asset-cash'), 11500);
  ok(r.invoke('deleteTransaction', { transaction_id: edit.transaction_id, expected_updated_at: edit.updated_at }));
  assert.equal(amount(r, 'asset-cash'), 13000); assert.equal(balances(r).netAssets, 10600);
});

test('recorded expenses before opening are not counted twice; same-day changes are reflected', () => {
  const r = createRuntime();
  ok(r.invoke('createTransaction', transaction(r, { date: '2025-12-31', request_id: 'old-transaction-001', amount: 10000 })));
  const prior = ok(r.invoke('createTransaction', transaction(r, { date: '2026-01-01', request_id: 'same-day-old-001', amount: 100 }))).transaction;
  ok(r.invoke('initializeAssets', opening(r, { 'asset-cash': 1000 })));
  assert.equal(amount(r, 'asset-cash'), 1000);
  ok(r.invoke('createTransaction', transaction(r, { date: '2026-01-01', request_id: 'same-day-new-001', amount: 200 })));
  assert.equal(amount(r, 'asset-cash'), 800);
  const edit = ok(r.invoke('updateTransaction', { ...prior, amount: 150, expected_updated_at: prior.updated_at })).transaction;
  assert.equal(amount(r, 'asset-cash'), 750);
  ok(r.invoke('deleteTransaction', { transaction_id: edit.transaction_id, expected_updated_at: edit.updated_at })); assert.equal(amount(r, 'asset-cash'), 900);
});

test('financial valuation is dated, historical values persist, and same-day reconciliation is absolute', () => {
  const r = initialized({ 'asset-financial': 200000, 'asset-cash': 10000 });
  ok(r.invoke('saveAssetBalances', revalue(r)));
  assert.equal(amount(r, 'asset-financial', '2026-02-28'), 200000); assert.equal(amount(r, 'asset-financial', '2026-04-30'), 230000);
  ok(r.invoke('createTransaction', transaction(r, { amount: 1000, date: '2026-04-01' })));
  assert.equal(amount(r, 'asset-cash'), 9000);
  ok(r.invoke('saveAssetBalances', revalue(r, { request_id: 'cash-reconcile-001', date: '2026-04-01', balances: { 'asset-cash': 12000 } })));
  assert.equal(amount(r, 'asset-cash'), 12000);
  ok(r.invoke('createTransaction', transaction(r, { request_id: 'post-reconcile-001', amount: 500, date: '2026-04-01' })));
  assert.equal(amount(r, 'asset-cash'), 11500);
  const dashboard = ok(r.invoke('getDashboard', '2026-03')); assert.equal(dashboard.income, 0); assert.equal(dashboard.expense, 0);
});

test('transfers, repayments and borrowing change both accounts and never count as income or expense', () => {
  const r = initialized({ 'asset-cash': 10000, 'asset-suica': 1000, 'asset-yucho': 100000, 'debt-jal-card': 2000 });
  const base = { request_id: 'transfer-request-001', date: '2026-01-05', from_account_id: 'asset-cash', to_account_id: 'asset-suica', amount: 3000 };
  const transfer = ok(r.invoke('createAssetMovement', base)); assert.equal(amount(r, 'asset-cash'), 7000); assert.equal(amount(r, 'asset-suica'), 4000);
  assert.equal(balances(r).netAssets, 109000);
  ok(r.invoke('createAssetMovement', { ...base, request_id: 'repayment-request1', from_account_id: 'asset-yucho', to_account_id: 'debt-jal-card', amount: 1500 }));
  assert.equal(balances(r).assets, 109500); assert.equal(balances(r).liabilities, 500); assert.equal(balances(r).netAssets, 109000);
  ok(r.invoke('createAssetMovement', { ...base, request_id: 'borrowing-request1', from_account_id: 'debt-other', to_account_id: 'asset-yucho', amount: 20000 }));
  assert.equal(balances(r).assets, 129500); assert.equal(balances(r).liabilities, 20500); assert.equal(balances(r).netAssets, 109000);
  assert.equal(ok(r.invoke('getDashboard', '2026-01')).expense, 0); assert.equal(ok(r.invoke('getDashboard', '2026-01')).income, 0);
  ok(r.invoke('deleteAssetMovement', { movement_id: transfer.movement_id })); ok(r.invoke('deleteAssetMovement', { movement_id: transfer.movement_id }));
  assert.equal(amount(r, 'asset-cash'), 10000); assert.equal(amount(r, 'asset-suica'), 1000); assert.equal(balances(r).movementCount, 2);
  assert.equal(r.book().getSheetByName('AssetMovements').getLastRow(), 4);
});

test('opening balances, revaluations and movements retry safely after an interrupted response', () => {
  const r = createRuntime(), input = opening(r, { 'asset-cash': 10000 });
  r.flushFailure = true; fail(r.invoke('initializeAssets', input), 'SERVER');
  assert.equal(ok(r.invoke('initializeAssets', input)).duplicate, true); assert.equal(r.book().getSheetByName('AssetSnapshots').getLastRow(), 2);
  const edit = revalue(r); r.flushFailure = true; fail(r.invoke('saveAssetBalances', edit), 'SERVER');
  assert.equal(ok(r.invoke('saveAssetBalances', edit)).duplicate, true); assert.equal(r.book().getSheetByName('AssetSnapshots').getLastRow(), 3);
  const movement = { request_id: 'lost-movement-001', date: '2026-04-01', from_account_id: 'asset-cash', to_account_id: 'asset-paypay', amount: 1000 };
  r.flushFailure = true; fail(r.invoke('createAssetMovement', movement), 'SERVER');
  assert.equal(ok(r.invoke('createAssetMovement', movement)).duplicate, true); assert.equal(amount(r, 'asset-cash'), 9000); assert.equal(amount(r, 'asset-paypay'), 1000);
  fail(r.invoke('createAssetMovement', { ...movement, amount: 2000 }), 'CONFLICT');
  fail(r.invoke('saveAssetBalances', { ...edit, balances: { 'asset-financial': 999 } }), 'CONFLICT');
});

test('snapshot revisions detect intervening transactions and lock contention prevents changes', () => {
  const r = initialized({ 'asset-cash': 10000 }), input = revalue(r);
  ok(r.invoke('createTransaction', transaction(r))); fail(r.invoke('saveAssetBalances', input), 'CONFLICT');
  assert.equal(amount(r, 'asset-financial'), 0);
  const refreshed = revalue(r); r.busy = true; fail(r.invoke('saveAssetBalances', refreshed), 'BUSY'); r.busy = false;
  ok(r.invoke('saveAssetBalances', refreshed)); assert.equal(amount(r, 'asset-financial'), 230000); assert.equal(r.lockHeld, false);
});

test('payment mapping updates affect future transactions while preserving old transaction accounts', () => {
  const r = initialized({ 'asset-cash': 10000, 'asset-paypay': 5000 });
  const prior = ok(r.invoke('createTransaction', transaction(r))).transaction;
  const link = ok(r.invoke('getBootstrap')).paymentLinks.find(l => l.payment_method_id === 'initial-payment-1');
  ok(r.invoke('savePaymentAccountLink', { payment_method_id: link.payment_method_id, account_id: 'asset-paypay', expected_updated_at: link.updated_at }));
  ok(r.invoke('createTransaction', transaction(r, { request_id: 'remapped-new-001', amount: 2000 })));
  assert.equal(amount(r, 'asset-cash'), 9000); assert.equal(amount(r, 'asset-paypay'), 3000);
  const edit = ok(r.invoke('updateTransaction', { ...prior, amount: 1500, expected_updated_at: prior.updated_at })).transaction;
  assert.equal(edit.asset_account_id, 'asset-cash'); assert.equal(amount(r, 'asset-cash'), 8500);
  fail(r.invoke('savePaymentAccountLink', { payment_method_id: link.payment_method_id, account_id: 'asset-suica', expected_updated_at: link.updated_at }), 'CONFLICT');
  const custom = ok(r.invoke('savePaymentMethod', { request_id: 'custom-paymethod1', name: '証券口座', sort_order: 100, is_active: true }));
  fail(r.invoke('createTransaction', transaction(r, { request_id: 'custom-purchase1', payment_method_id: custom.payment_method_id })), 'VALIDATION');
  ok(r.invoke('savePaymentAccountLink', { payment_method_id: custom.payment_method_id, account_id: 'asset-financial', expected_updated_at: '' }));
  const created = ok(r.invoke('createTransaction', transaction(r, { request_id: 'custom-purchase1', payment_method_id: custom.payment_method_id }))).transaction;
  assert.equal(created.asset_account_id, 'asset-financial');
});

test('12-month asset history crosses years, preserves unregistered months and supports negative net assets', () => {
  const r = initialized({ 'asset-cash': 1000, 'debt-other': 2000 });
  const data = balances(r, '2026-01-31'); assert.equal(data.netAssets, -1000); assert.equal(data.months.length, 12);
  assert.equal(data.months[0].month, '2025-02'); assert.equal(data.months[0].assets, null); assert.equal(data.months[0].known, false);
  assert.equal(data.months[11].month, '2026-01'); assert.equal(data.months[11].netAssets, -1000);
  const before = balances(r, '2025-12-31'); assert.equal(before.known, false); assert.equal(before.accounts[0].amount, null);
  ok(r.invoke('createTransaction', transaction(r, { request_id: 'future-dated-001', date: '2026-12-01', amount: 100 })));
  assert.equal(balances(r, '2026-10-13').assets, 1000);
});

test('invalid balances and transfers are rejected before writing; assets APIs enforce owner identity', () => {
  const r = createRuntime(), input = opening(r);
  for (const bad of [{ ...input, balances: {} }, { ...input, balances: { ...input.balances, 'asset-cash': -1 } }, { ...input, balances: { ...input.balances, 'asset-cash': 1.1 } }, { ...input, date: '9999-12-31' }, { ...input, balances: { ...input.balances, unknown: 1 } }]) fail(r.invoke('initializeAssets', bad), 'VALIDATION');
  assert.equal(r.book().getSheetByName('AssetSnapshots').getLastRow(), 1);
  ok(r.invoke('initializeAssets', input)); fail(r.invoke('initializeAssets', { ...input, request_id: 'second-opening01' }), 'CONFLICT');
  const base = { request_id: 'invalid-move-001', date: '2026-01-01', from_account_id: 'asset-cash', to_account_id: 'asset-cash', amount: 1 };
  fail(r.invoke('createAssetMovement', base), 'VALIDATION'); fail(r.invoke('createAssetMovement', { ...base, to_account_id: 'asset-suica', amount: 0 }), 'VALIDATION');
  fail(r.invoke('saveAssetBalances', revalue(r, { date: '2025-12-31' })), 'VALIDATION');
  r.activeEmail = 'other@example.test';
  for (const method of ['getAssets', 'initializeAssets', 'saveAssetBalances', 'savePaymentAccountLink', 'createAssetMovement', 'deleteAssetMovement']) fail(r.invoke(method, {}), 'AUTH');
});

test('upgrade from the 13-column app preserves A-M, adds accounts and maps renamed initial methods', () => {
  const r = createRuntime(); const created = ok(r.invoke('createTransaction', transaction(r))).transaction;
  const boot = ok(r.invoke('getBootstrap')), cash = boot.paymentMethods.find(m => m.payment_method_id === 'initial-payment-1');
  ok(r.invoke('savePaymentMethod', { ...cash, name: '財布', expected_updated_at: cash.updated_at }));
  const sheet = r.book().getSheetByName('Transactions'); sheet.rows = sheet.rows.map(row => row.slice(0, 13));
  for (const name of ['AssetAccounts', 'PaymentAccountLinks', 'AssetSnapshots', 'AssetMovements']) r.book().sheets.delete(name);
  r.properties.delete('ASSET_MASTERS_SEEDED'); const original = JSON.stringify(sheet.rows);
  r.invoke('setupDatabase'); assert.equal(JSON.stringify(sheet.rows.map(row => row.slice(0, 13))), original);
  assert.equal(sheet.rows[0][13], 'asset_account_id'); assert.equal(sheet.rows[1][13], 'asset-cash');
  assert.equal(ok(r.invoke('listTransactions')).items[0].transaction_id, created.transaction_id);
  assert.equal(ok(r.invoke('getBootstrap')).paymentMethods.find(m => m.payment_method_id === cash.payment_method_id).name, '財布');
});
