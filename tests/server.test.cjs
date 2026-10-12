const test = require('node:test');
const assert = require('node:assert/strict');
const { createRuntime, Book } = require('../scripts/gas-runtime.cjs');

function payload(runtime, extra = {}) {
  const data = ok(runtime.invoke('getBootstrap'));
  return { request_id: 'request-00000001', date: '2026-10-13', type: 'expense', payment_method_id: data.paymentMethods[0].payment_method_id, category_id: data.categories.find(c => c.type === 'expense' && c.name === '食費').category_id, amount: 680, note: '昼食', ...extra };
}
function ok(response) { assert.equal(response.ok, true, response.error?.message); return response.data; }
function fail(response, code) { assert.equal(response.ok, false); assert.equal(response.error.code, code); }
function rows(runtime, name = 'Transactions') { return runtime.book().getSheetByName(name).getLastRow() - 1; }

test('initialization seeds 7 payment methods and 15 categories; rerun preserves data and settings', () => {
  const r = createRuntime(), bootstrap = ok(r.invoke('getBootstrap'));
  assert.equal(bootstrap.paymentMethods.length, 7); assert.equal(bootstrap.categories.length, 15);
  assert.equal(bootstrap.categories.filter(c => c.type === 'expense').length, 11);
  const created = ok(r.invoke('createTransaction', payload(r))).transaction;
  const category = bootstrap.categories[0];
  ok(r.invoke('saveCategory', { ...category, name: '変更した名称', is_active: false, expected_updated_at: category.updated_at }));
  const before = JSON.stringify(r.book().getSheetByName('Transactions').rows);
  const id = r.properties.get('SPREADSHEET_ID');
  r.invoke('setupDatabase');
  assert.equal(r.properties.get('SPREADSHEET_ID'), id);
  assert.equal(JSON.stringify(r.book().getSheetByName('Transactions').rows), before);
  assert.equal(ok(r.invoke('getBootstrap')).categories.find(c => c.category_id === category.category_id).name, '変更した名称');
  assert.equal(rows(r), 1); assert.ok(created.transaction_id);
});

test('new expense stores a numeric amount, Tokyo dates, IDs and ordered schema', () => {
  const r = createRuntime(); const data = ok(r.invoke('createTransaction', payload(r)));
  const row = r.book().getSheetByName('Transactions').rows[1];
  assert.equal(row[1], '2026-10-13'); assert.equal(row[5], 680); assert.equal(typeof row[5], 'number'); assert.equal(row[6], '昼食');
  assert.match(row[7], /\+09:00$/); assert.equal(row[9], ''); assert.equal(row[10], 'request-00000001');
  assert.equal(row[11], data.transaction.category_id); assert.equal(row[12], '食費'); assert.equal(data.duplicate, false);
  assert.equal(r.lockHeld, false);
});

test('income uses income category and same payment-method master', () => {
  const r = createRuntime(); const category = ok(r.invoke('getBootstrap')).categories.find(c => c.type === 'income');
  const data = ok(r.invoke('createTransaction', payload(r, { type: 'income', category_id: category.category_id, amount: 100000 })));
  assert.equal(data.transaction.type, 'income'); assert.equal(ok(r.invoke('getDashboard', '2026-10')).income, 100000);
});

test('rejects missing fields, unsafe amounts, malformed types, IDs and dates without writing', () => {
  const r = createRuntime();
  for (const change of [{ amount: 0 }, { amount: -1 }, { amount: 1.1 }, { amount: '680' }, { amount: Number.MAX_SAFE_INTEGER + 1 }, { amount: null }, { payment_method_id: '' }, { category_id: '' }, { type: 'transfer' }, { date: '2026-02-29' }, { date: '2026-13-01' }, { date: '2026-04-31' }, { date: '2026-1-01' }, { date: '' }, { date: '1899-01-01' }, { request_id: '=formula()' }, { note: null }, { note: 'x'.repeat(10001) }]) {
    fail(r.invoke('createTransaction', payload(r, change)), 'VALIDATION');
  }
  const category = ok(r.invoke('getBootstrap')).categories.find(c => c.type === 'income');
  fail(r.invoke('createTransaction', payload(r, { category_id: category.category_id })), 'VALIDATION');
  fail(r.invoke('createTransaction', null), 'VALIDATION'); assert.equal(rows(r), 0); assert.equal(r.lockHeld, false);
});

test('leap day handles Gregorian century rules', () => {
  const r = createRuntime();
  for (const date of ['2024-02-29', '2000-02-29', '2026-02-28']) ok(r.invoke('createTransaction', payload(r, { date, request_id: 'request-' + date })));
  for (const date of ['1900-02-29', '2100-02-29', '2025-02-29', '2026-00-10', '2026-02-00']) fail(r.invoke('createTransaction', payload(r, { date })), 'VALIDATION');
  assert.equal(rows(r), 3);
});

test('request retries are deduplicated even after a write succeeds but the response fails', () => {
  const r = createRuntime(), input = payload(r);
  r.flushFailure = true;
  fail(r.invoke('createTransaction', input), 'SERVER'); assert.equal(rows(r), 1);
  const retry = ok(r.invoke('createTransaction', input)); assert.equal(retry.duplicate, true); assert.equal(rows(r), 1);
  fail(r.invoke('createTransaction', { ...input, amount: 700 }), 'CONFLICT'); assert.equal(rows(r), 1);
  assert.equal(r.lockHeld, false);
});

test('busy lock prevents write; successful retries use same ID', () => {
  const r = createRuntime(), input = payload(r); r.busy = true;
  fail(r.invoke('createTransaction', input), 'BUSY'); assert.equal(rows(r), 0);
  r.busy = false; ok(r.invoke('createTransaction', input)); assert.equal(rows(r), 1);
});

test('inactive masters stay visible for old transactions, allow retained edit, reject new selection', () => {
  const r = createRuntime(), input = payload(r); const original = ok(r.invoke('createTransaction', input)).transaction;
  const boot = ok(r.invoke('getBootstrap')); const cat = boot.categories.find(c => c.category_id === input.category_id); const method = boot.paymentMethods.find(m => m.payment_method_id === input.payment_method_id);
  const newCat = ok(r.invoke('saveCategory', { ...cat, name: '食事', is_active: false, expected_updated_at: cat.updated_at }));
  const newMethod = ok(r.invoke('savePaymentMethod', { ...method, name: 'お財布', is_active: false, expected_updated_at: method.updated_at }));
  fail(r.invoke('createTransaction', { ...input, request_id: 'different-request' }), 'VALIDATION');
  assert.equal(ok(r.invoke('createTransaction', input)).duplicate, true);
  const edited = ok(r.invoke('updateTransaction', { ...input, transaction_id: original.transaction_id, expected_updated_at: original.updated_at, amount: 900 })).transaction;
  assert.equal(edited.category_name, '食事'); assert.equal(edited.payment_method_name, 'お財布'); assert.equal(edited.transaction_id, original.transaction_id); assert.equal(edited.created_at, original.created_at);
  assert.notEqual(edited.updated_at, original.updated_at);
  assert.equal(ok(r.invoke('getDashboard', '2026-10')).categories[0].name, '食事');
  ok(r.invoke('saveCategory', { ...newCat, is_active: true, expected_updated_at: newCat.updated_at }));
  ok(r.invoke('savePaymentMethod', { ...newMethod, is_active: true, expected_updated_at: newMethod.updated_at }));
  ok(r.invoke('createTransaction', { ...input, request_id: 'reactivated-request' })); assert.equal(rows(r), 2);
});

test('master additions are idempotent; duplicates include inactive and normalized names', () => {
  const r = createRuntime(); const input = { request_id: 'new-category-001', type: 'expense', name: 'ペット', sort_order: 120, is_active: true };
  const category = ok(r.invoke('saveCategory', input)); assert.equal(ok(r.invoke('saveCategory', input)).category_id, category.category_id);
  fail(r.invoke('saveCategory', { ...input, request_id: 'new-category-002' }), 'VALIDATION');
  ok(r.invoke('saveCategory', { ...input, type: 'income', request_id: 'new-category-003' }));
  fail(r.invoke('saveCategory', { ...category, type: 'income', expected_updated_at: category.updated_at }), 'VALIDATION');
  fail(r.invoke('saveCategory', { ...category, sort_order: -1, expected_updated_at: category.updated_at }), 'VALIDATION');
  const method = { request_id: 'new-payment-0001', name: 'ABC', sort_order: 0, is_active: false };
  ok(r.invoke('savePaymentMethod', method)); fail(r.invoke('savePaymentMethod', { ...method, name: 'ＡＢＣ', request_id: 'new-payment-0002' }), 'VALIDATION');
  assert.equal(ok(r.invoke('getBootstrap')).paymentMethods[0].name, 'ABC');
});

test('edit concurrency conflict preserves newer data; deletion is logical and idempotent', () => {
  const r = createRuntime(), input = payload(r); const original = ok(r.invoke('createTransaction', input)).transaction;
  const newer = ok(r.invoke('updateTransaction', { ...input, transaction_id: original.transaction_id, expected_updated_at: original.updated_at, amount: 800 })).transaction;
  fail(r.invoke('updateTransaction', { ...input, transaction_id: original.transaction_id, expected_updated_at: original.updated_at, amount: 900 }), 'CONFLICT');
  fail(r.invoke('deleteTransaction', { transaction_id: original.transaction_id, expected_updated_at: original.updated_at }), 'CONFLICT');
  ok(r.invoke('deleteTransaction', { transaction_id: newer.transaction_id, expected_updated_at: newer.updated_at }));
  ok(r.invoke('deleteTransaction', { transaction_id: newer.transaction_id, expected_updated_at: newer.updated_at }));
  assert.equal(rows(r), 1); assert.equal(r.book().getSheetByName('Transactions').rows[1][5], 800); assert.ok(r.book().getSheetByName('Transactions').rows[1][9]);
  assert.equal(ok(r.invoke('listTransactions')).total, 0); assert.equal(ok(r.invoke('getDashboard', '2026-10')).expense, 0);
  fail(r.invoke('updateTransaction', { ...newer, expected_updated_at: newer.updated_at }), 'NOT_FOUND');
});

test('lost update response can be retried without a false concurrency conflict', () => {
  const r = createRuntime(), input = payload(r), old = ok(r.invoke('createTransaction', input)).transaction;
  const edit = { ...input, transaction_id: old.transaction_id, expected_updated_at: old.updated_at, amount: 720 };
  r.flushFailure = true; fail(r.invoke('updateTransaction', edit), 'SERVER');
  const retried = ok(r.invoke('updateTransaction', edit)); assert.equal(retried.duplicate, true); assert.equal(retried.transaction.amount, 720); assert.equal(rows(r), 1);
});

test('lost master edit response can be retried and initialization can resume partial seeding', () => {
  const r = createRuntime(), category = ok(r.invoke('getBootstrap')).categories[0];
  const edit = { ...category, name: '変更後', expected_updated_at: category.updated_at };
  r.flushFailure = true; fail(r.invoke('saveCategory', edit), 'SERVER');
  assert.equal(ok(r.invoke('saveCategory', edit)).name, '変更後');
  const partial = createRuntime({ initialized: false }); partial.flushFailure = true;
  assert.throws(() => partial.invoke('setupDatabase'), /interruption/);
  partial.invoke('setupDatabase');
  assert.equal(ok(partial.invoke('getBootstrap')).categories.length, 15); assert.equal(ok(partial.invoke('getBootstrap')).paymentMethods.length, 7);
});

test('monthly totals, category ratios and 12-month trend use transaction date across years', () => {
  const r = createRuntime(); const boot = ok(r.invoke('getBootstrap')); const food = boot.categories.find(c => c.name === '食費'), daily = boot.categories.find(c => c.name === '日用品'), salary = boot.categories.find(c => c.type === 'income');
  for (const [i, extra] of [
    { date: '2026-01-03', category_id: food.category_id, amount: 100 },
    { date: '2026-01-03', category_id: food.category_id, amount: 200 },
    { date: '2026-01-05', category_id: daily.category_id, amount: 700 },
    { date: '2026-01-05', category_id: salary.category_id, type: 'income', amount: 2000 },
    { date: '2025-12-31', amount: 500 }, { date: '2025-01-01', amount: 999 }
  ].entries()) ok(r.invoke('createTransaction', payload(r, { ...extra, request_id: 'aggregation-' + String(i).padStart(4, '0') })));
  const d = ok(r.invoke('getDashboard', '2026-01'));
  assert.equal(d.income, 2000); assert.equal(d.expense, 1000); assert.equal(d.balance, 1000);
  assert.equal(d.categories.find(c => c.category_id === food.category_id).amount, 300);
  assert.equal(d.categories.find(c => c.category_id === food.category_id).ratio, .3);
  assert.equal(d.months.length, 12); assert.equal(d.months[0].month, '2025-02'); assert.equal(d.months[10].month, '2025-12'); assert.equal(d.months[10].expense, 500);
  assert.equal(d.months.reduce((sum, m) => sum + m.expense, 0), 1500);
  const empty = ok(r.invoke('getDashboard', '2027-02')); assert.equal(empty.income, 0); assert.equal(empty.expense, 0); assert.deepEqual(empty.categories, []);
});

test('history is sorted, filtered, paginated in 50-row pages, and excludes deletions', () => {
  const r = createRuntime();
  for (let i = 0; i < 53; i++) ok(r.invoke('createTransaction', payload(r, { request_id: 'history-' + String(i).padStart(4, '0'), date: i === 0 ? '2026-09-30' : '2026-10-13', amount: i + 1 })));
  const sheet = r.book().getSheetByName('Transactions');
  sheet.rows[1][7] = '2026-10-20T10:00:00.000+09:00'; // Registration time cannot override the date ordering.
  sheet.rows[2][7] = '2026-10-15T10:00:00.000+09:00';
  const first = ok(r.invoke('listTransactions', { month: '2026-10', type: 'expense', page: 1 }));
  assert.equal(first.total, 52); assert.equal(first.items.length, 50); assert.equal(first.pages, 2); assert.equal(first.items[0].amount, 2);
  assert.equal(ok(r.invoke('listTransactions', { month: '2026-10', page: 2 })).items.length, 2);
  assert.equal(ok(r.invoke('listTransactions', { type: 'income' })).total, 0);
  assert.equal(ok(r.invoke('listTransactions')).items[0].date, '2026-10-13');
  fail(r.invoke('listTransactions', { page: 0 }), 'VALIDATION');
});

test('literal names and multiline notes cannot become spreadsheet formulas; values round-trip', () => {
  const r = createRuntime();
  for (const [i, note] of ['=IMPORTXML("https://evil.test","//x")', '+SUM(A1)', '-1+2', '@evil', "'quoted", '<script>alert(1)</script>\n昼食', '\t=1+1'].entries()) {
    const transaction = ok(r.invoke('createTransaction', payload(r, { request_id: 'literal-' + String(i).padStart(4, '0'), note }))).transaction;
    assert.equal(transaction.note, note);
  }
  const sheet = r.book().getSheetByName('Transactions'); assert.equal(sheet.formulas.size, 0);
  const method = ok(r.invoke('savePaymentMethod', { request_id: 'literal-payment1', name: '=1+1', sort_order: 1, is_active: true }));
  assert.equal(method.name, '=1+1'); assert.equal(r.book().getSheetByName('PaymentMethods').formulas.size, 0);
});

test('server requires owner identity for reads and writes; blank identities are denied', () => {
  const r = createRuntime(), input = payload(r);
  for (const email of ['other@example.test', '']) {
    r.activeEmail = email;
    for (const [method, argument] of [['getBootstrap'], ['createTransaction', input], ['listTransactions'], ['getDashboard', '2026-10'], ['saveCategory', {}], ['deleteTransaction', {}], ['updateTransaction', {}]]) fail(r.invoke(method, argument), 'AUTH');
    assert.throws(() => r.invoke('setupDatabase'));
  }
  assert.equal(rows(r), 0);
});

test('v1.0 A-K migration appends category headers and keeps all legacy values', () => {
  const r = createRuntime({ initialized: false }), book = new Book('existing-book'); r.books.set(book.id, book); r.properties.set('SPREADSHEET_ID', book.id);
  const sheet = book.insertSheet('Transactions');
  const headers = Array.from(r.context.HEADERS_.Transactions).slice(0, 11);
  sheet.getRange(1, 1, 1, 11).setValues([headers]);
  const original = ['legacy-transaction', '2026-10-13', 'expense', 'initial-payment-1', '現金', 680, '以前の記録', '2026-10-13T01:00:00.000+09:00', '2026-10-13T01:00:00.000+09:00', '', 'legacy-request01'];
  sheet.getRange(2, 1, 1, 11).setValues([original]); r.invoke('setupDatabase');
  assert.deepEqual(sheet.rows[1].slice(0, 11), original); assert.equal(sheet.rows[0][11], 'category_id');
  assert.equal(ok(r.invoke('getDashboard', '2026-10')).categories[0].name, '未分類');
  assert.equal(ok(r.invoke('listTransactions')).items[0].category_name, '未分類');
});

test('incompatible existing schema is rejected without overwriting sheet contents', () => {
  const r = createRuntime({ initialized: false }), book = new Book('existing-book'); r.books.set(book.id, book); r.properties.set('SPREADSHEET_ID', book.id);
  const sheet = book.insertSheet('Transactions'); sheet.getRange(1, 1).setValues([['wrong-header']]);
  const before = JSON.stringify(sheet.rows); assert.throws(() => r.invoke('setupDatabase'), /ヘッダー/); assert.equal(JSON.stringify(sheet.rows), before);
});

test('timestamp conversion uses Tokyo date and amounts cannot overflow aggregation silently', () => {
  const r = createRuntime(); const input = payload(r); ok(r.invoke('createTransaction', input));
  const sheet = r.book().getSheetByName('Transactions'); sheet.rows[1][1] = new Date('2026-10-12T15:30:00Z');
  assert.equal(ok(r.invoke('listTransactions')).items[0].date, '2026-10-13');
  ok(r.invoke('createTransaction', payload(r, { request_id: 'large-request-001', amount: Number.MAX_SAFE_INTEGER })));
  fail(r.invoke('getDashboard', '2026-10'), 'DATA');
});
