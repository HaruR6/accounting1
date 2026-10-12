/* Google Apps Script V8. Public functions return JSON-safe envelopes. */
var TZ_ = 'Asia/Tokyo';
var HEADERS_ = {
  Transactions: ['transaction_id', 'date', 'type', 'payment_method_id', 'payment_method_name', 'amount', 'note', 'created_at', 'updated_at', 'deleted_at', 'request_id', 'category_id', 'category_name', 'asset_account_id'],
  PaymentMethods: ['payment_method_id', 'name', 'sort_order', 'is_active', 'created_at', 'updated_at'],
  Categories: ['category_id', 'type', 'name', 'sort_order', 'is_active', 'created_at', 'updated_at'],
  AssetAccounts: ['account_id', 'name', 'kind', 'sort_order'],
  PaymentAccountLinks: ['payment_method_id', 'account_id', 'updated_at'],
  AssetSnapshots: ['snapshot_id', 'date', 'balances_json', 'included_effects_json', 'created_at', 'request_id'],
  AssetMovements: ['movement_id', 'date', 'from_account_id', 'to_account_id', 'amount', 'created_at', 'deleted_at', 'request_id']
};
var ACCOUNT_DEFAULTS_ = [
  ['asset-cash', '現金', 'asset'], ['asset-yucho', 'ゆうちょ銀行', 'asset'], ['asset-jrebank', 'JREBANK', 'asset'],
  ['asset-financial', '金融資産', 'asset'], ['asset-suica', 'SUICA', 'asset'], ['asset-paypay', 'PayPay', 'asset'],
  ['debt-jal-card', 'JALカード未払額', 'liability'], ['debt-jre-card', 'JREカード未払額', 'liability'], ['debt-other', 'その他の負債', 'liability']
];
var PAYMENT_DEFAULTS_ = ['現金', 'クレジットカード（JALカード）', 'クレジットカード（JREカード）', 'Suica', 'ゆうちょ銀行', 'JREバンク', 'PayPay'];
var CATEGORY_DEFAULTS_ = {
  expense: ['食費', '日用品', '交通費', '住居費', '光熱費', '通信費', '医療費', '趣味・娯楽', '教育費', '交際費', 'その他'],
  income: ['給与・アルバイト', '仕送り', '臨時収入', 'その他']
};

function doGet() {
  authorize_();
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('セミナー会計　早瀬ユウカ監修　家計簿アプリ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include_(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

// Run once in the Apps Script editor. Safe to run again; never clears data.
function setupDatabase() {
  var email = Session.getActiveUser().getEmail();
  var effective = Session.getEffectiveUser().getEmail();
  require_(email && email === effective, 'AUTH', '本人のアカウントからスクリプトエディタで実行してください。');
  var properties = PropertiesService.getScriptProperties();
  var owner = properties.getProperty('OWNER_EMAIL');
  require_(!owner || owner === email, 'AUTH', '設定済みの所有者とアカウントが異なります。');
  return locked_(function () {
    var id = properties.getProperty('SPREADSHEET_ID');
    var book = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('セミナー会計　早瀬ユウカ監修　家計簿アプリ');
    // Save immediately so a retry after interrupted initialization reuses the book.
    if (!id) properties.setProperty('SPREADSHEET_ID', book.getId());
    book.setSpreadsheetTimeZone(TZ_);
    Object.keys(HEADERS_).forEach(function (name) { initializeSheet_(book, name); });
    seedMasters_(book);
    seedAssetMasters_(book);
    properties.setProperty('OWNER_EMAIL', email);
    return { spreadsheetId: book.getId(), message: '初期化完了。既存データは保持しました。' };
  });
}

function getBootstrap() {
  return api_(function () {
    var book = database_();
    return { today: today_(), paymentMethods: masters_(book, 'PaymentMethods'), categories: masters_(book, 'Categories'),
      accounts: masters_(book, 'AssetAccounts'), paymentLinks: records_(book, 'PaymentAccountLinks').map(publicRecord_), assetsInitialized: records_(book, 'AssetSnapshots').length > 0 };
  });
}

function createTransaction(input) {
  return api_(function () {
    return locked_(function () {
      var book = database_();
      require_(input && typeof input === 'object', 'VALIDATION', '登録内容がありません。');
      validId_(input.request_id, 'リクエストID');
      var rows = records_(book, 'Transactions');
      var duplicate = rows.find(function (r) { return r.request_id === input.request_id; });
      if (duplicate) {
        // Do not revalidate master activity: an earlier save may already have succeeded.
        require_(['date', 'type', 'payment_method_id', 'category_id', 'amount', 'note'].every(function (key) {
          return duplicate[key] === input[key];
        }), 'CONFLICT', '同じリクエストIDの保存内容が異なります。履歴を確認してください。');
        return { transaction: displayTransaction_(duplicate, book), duplicate: true };
      }
      var data = validateTransaction_(input, book, null);
      var now = timestamp_();
      var row = Object.assign(data, { transaction_id: Utilities.getUuid(), created_at: now, updated_at: now, deleted_at: '', request_id: input.request_id });
      appendRecord_(book, 'Transactions', row);
      SpreadsheetApp.flush();
      return { transaction: displayTransaction_(row, book), duplicate: false };
    });
  });
}

function updateTransaction(input) {
  return api_(function () {
    return locked_(function () {
      var book = database_();
      require_(input && typeof input === 'object', 'VALIDATION', '編集内容がありません。');
      var old = transactionById_(book, input.transaction_id);
      require_(!old.deleted_at, 'NOT_FOUND', 'この取引は削除済みです。');
      // Recover an update whose response was lost without overwriting newer data.
      if (input.expected_updated_at !== old.updated_at && ['date', 'type', 'payment_method_id', 'category_id', 'amount', 'note'].every(function (key) { return old[key] === input[key]; })) {
        return { transaction: displayTransaction_(old, book), duplicate: true };
      }
      require_(input.expected_updated_at === old.updated_at, 'CONFLICT', '別の画面で更新されました。履歴を読み直してください。');
      var data = validateTransaction_(input, book, old);
      var row = Object.assign({}, old, data, { updated_at: timestamp_(old.updated_at) });
      writeRecord_(book, 'Transactions', old._row, row);
      SpreadsheetApp.flush();
      return { transaction: displayTransaction_(row, book) };
    });
  });
}

function deleteTransaction(input) {
  return api_(function () {
    return locked_(function () {
      var book = database_();
      require_(input && typeof input === 'object', 'VALIDATION', '取引を指定してください。');
      var row = transactionById_(book, input.transaction_id);
      if (row.deleted_at) return { transaction_id: row.transaction_id, deleted: true };
      require_(input.expected_updated_at === row.updated_at, 'CONFLICT', '別の画面で更新されました。履歴を読み直してください。');
      row.deleted_at = timestamp_(row.updated_at);
      row.updated_at = row.deleted_at;
      writeRecord_(book, 'Transactions', row._row, row);
      SpreadsheetApp.flush();
      return { transaction_id: row.transaction_id, deleted: true };
    });
  });
}

function listTransactions(filter) {
  return api_(function () {
    filter = filter || {};
    var book = database_();
    if (filter.month) validMonth_(filter.month);
    if (filter.type) validType_(filter.type);
    if (filter.category_id) validId_(filter.category_id, 'カテゴリID');
    var page = filter.page === undefined ? 1 : filter.page;
    require_(Number.isSafeInteger(page) && page >= 1, 'VALIDATION', 'ページ番号が不正です。');
    var rows = records_(book, 'Transactions').filter(function (r) {
      return !r.deleted_at && (!filter.month || r.date.slice(0, 7) === filter.month) &&
        (!filter.type || r.type === filter.type) && (!filter.category_id || r.category_id === filter.category_id);
    }).sort(function (a, b) {
      return b.date.localeCompare(a.date) || b.created_at.localeCompare(a.created_at) || b.transaction_id.localeCompare(a.transaction_id);
    });
    var pages = Math.max(1, Math.ceil(rows.length / 50));
    page = Math.min(page, pages);
    var methods = masters_(book, 'PaymentMethods');
    var categories = masters_(book, 'Categories');
    return { items: rows.slice((page - 1) * 50, page * 50).map(function (row) {
      return displayWithMasters_(row, methods, categories);
    }), page: page, pages: pages, total: rows.length, pageSize: 50 };
  });
}

function getDashboard(month) {
  return api_(function () {
    validMonth_(month);
    var book = database_();
    var categories = masters_(book, 'Categories');
    var months = [];
    var year = Number(month.slice(0, 4));
    var m = Number(month.slice(5));
    for (var offset = 11; offset >= 0; offset--) {
      var index = year * 12 + m - 1 - offset;
      months.push({ month: String(Math.floor(index / 12)).padStart(4, '0') + '-' + String(index % 12 + 1).padStart(2, '0'), income: 0, expense: 0 });
    }
    var byMonth = {};
    months.forEach(function (r) { byMonth[r.month] = r; });
    var byCategory = {};
    records_(book, 'Transactions').forEach(function (r) {
      if (r.deleted_at) return;
      var key = r.date.slice(0, 7);
      if (byMonth[key]) byMonth[key][r.type] = safeAdd_(byMonth[key][r.type], r.amount);
      if (key === month && r.type === 'expense') {
        var id = r.category_id || '__uncategorized';
        if (!byCategory[id]) {
          var master = categories.find(function (c) { return c.category_id === id; });
          byCategory[id] = { category_id: id, name: master ? master.name : (r.category_name || '未分類'), amount: 0 };
        }
        byCategory[id].amount = safeAdd_(byCategory[id].amount, r.amount);
      }
    });
    var selected = byMonth[month];
    var breakdown = Object.keys(byCategory).map(function (id) {
      var item = byCategory[id];
      item.ratio = selected.expense ? item.amount / selected.expense : 0;
      return item;
    }).sort(function (a, b) { return b.amount - a.amount || a.name.localeCompare(b.name, 'ja'); });
    return { month: month, income: selected.income, expense: selected.expense, balance: selected.income - selected.expense, categories: breakdown, months: months };
  });
}

function savePaymentMethod(input) {
  return api_(function () { return saveMaster_(input, 'PaymentMethods'); });
}

function saveCategory(input) {
  return api_(function () { return saveMaster_(input, 'Categories'); });
}

function saveMaster_(input, sheetName) {
  return locked_(function () {
    var book = database_();
    require_(input && typeof input === 'object', 'VALIDATION', '設定内容がありません。');
    var idKey = sheetName === 'Categories' ? 'category_id' : 'payment_method_id';
    var name = typeof input.name === 'string' ? input.name.trim() : '';
    require_(name && name.length <= 100 && !/[\r\n\t]/.test(name), 'VALIDATION', '名称は改行のない1〜100文字で入力してください。');
    require_(Number.isSafeInteger(input.sort_order) && input.sort_order >= 0 && input.sort_order <= 1000000, 'VALIDATION', '表示順は0〜1,000,000の整数で入力してください。');
    require_(typeof input.is_active === 'boolean', 'VALIDATION', '使用状態が不正です。');
    if (sheetName === 'Categories') validType_(input.type);
    var rows = records_(book, sheetName);
    var old = null;
    if (input[idKey]) {
      validId_(input[idKey], '設定ID');
      old = rows.find(function (r) { return r[idKey] === input[idKey]; });
      require_(old, 'NOT_FOUND', '設定項目が見つかりません。');
      if (sheetName === 'Categories') require_(old.type === input.type, 'VALIDATION', 'カテゴリの収支区分は変更できません。');
      if (input.expected_updated_at !== old.updated_at && old.name === name && old.sort_order === input.sort_order && old.is_active === input.is_active) return publicRecord_(old);
      require_(input.expected_updated_at === old.updated_at, 'CONFLICT', '設定が更新されました。読み直してください。');
    } else {
      validId_(input.request_id, 'リクエストID');
      // The caller-generated UUID is also the master ID: safe to retry additions.
      old = rows.find(function (r) { return r[idKey] === input.request_id; });
      if (old) {
        require_(old.name === name && old.sort_order === input.sort_order && old.is_active === input.is_active &&
          (sheetName !== 'Categories' || old.type === input.type), 'CONFLICT', '同じリクエストIDの設定内容が異なります。');
        return publicRecord_(old);
      }
    }
    require_(!rows.some(function (r) {
      return (!old || r[idKey] !== old[idKey]) && r.name.normalize('NFKC').toLocaleLowerCase() === name.normalize('NFKC').toLocaleLowerCase() &&
        (sheetName !== 'Categories' || r.type === input.type);
    }), 'VALIDATION', '同じ名称の項目が既にあります（使用停止中を含む）。');
    var now = timestamp_(old && old.updated_at);
    var row = Object.assign({}, old || {}, { name: name, sort_order: input.sort_order, is_active: input.is_active, created_at: old ? old.created_at : now, updated_at: now });
    row[idKey] = old ? old[idKey] : input.request_id;
    if (sheetName === 'Categories') row.type = input.type;
    if (old) writeRecord_(book, sheetName, old._row, row);
    else appendRecord_(book, sheetName, row);
    SpreadsheetApp.flush();
    return publicRecord_(row);
  });
}

function authorize_() {
  var owner = PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL');
  require_(owner, 'SETUP', 'スクリプトエディタで setupDatabase を実行してください。');
  require_(Session.getActiveUser().getEmail() === owner, 'AUTH', '本人のGoogleアカウントでアクセスしてください。');
}

function api_(fn) {
  try {
    authorize_();
    return { ok: true, data: fn() };
  } catch (error) {
    var known = error.appCode;
    if (!known) console.error(error.stack || error.message);
    return { ok: false, error: { code: known || 'SERVER', message: known ? error.message : '保存先との処理に失敗しました。時間をおいて再試行してください。', retryable: !known || known === 'BUSY' } };
  }
}

function require_(condition, code, message) {
  if (!condition) { var error = new Error(message); error.appCode = code; throw error; }
}

function locked_(fn) {
  var lock = LockService.getScriptLock();
  require_(lock.tryLock(10000), 'BUSY', 'ほかの保存処理が実行中です。少し待って再試行してください。');
  try { return fn(); } finally { lock.releaseLock(); }
}

function database_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  require_(id, 'SETUP', '保存先が未設定です。setupDatabase を実行してください。');
  var book = SpreadsheetApp.openById(id);
  Object.keys(HEADERS_).forEach(function (name) {
    var sheet = book.getSheetByName(name);
    require_(sheet, 'SETUP', name + ' シートがありません。setupDatabase を実行してください。');
    var actual = sheet.getRange(1, 1, 1, HEADERS_[name].length).getValues()[0];
    require_(HEADERS_[name].every(function (header, i) { return actual[i] === header; }), 'SETUP', name + ' のヘッダーが仕様と異なります。READMEを確認してください。');
  });
  return book;
}

function initializeSheet_(book, name) {
  var sheet = book.getSheetByName(name) || book.insertSheet(name);
  var expected = HEADERS_[name];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    sheet.setFrozenRows(1);
    return;
  }
  var actual = sheet.getRange(1, 1, 1, expected.length).getValues()[0];
  // v1.0 migration: append only L/M headers, preserving every A-K cell.
  if (name === 'Transactions' && actual.slice(0, 11).every(function (h, i) { return h === expected[i]; }) && !actual[11] && !actual[12]) {
    require_(sheet.getLastColumn() <= 11, 'SETUP', 'L・M列に既存データがあります。自動移行を中止しました。');
    sheet.getRange(1, 12, 1, 2).setValues([expected.slice(11, 13)]);
    actual = sheet.getRange(1, 1, 1, expected.length).getValues()[0];
  }
  if (name === 'Transactions' && actual.slice(0, 13).every(function (h, i) { return h === expected[i]; }) && !actual[13]) {
    require_(sheet.getLastColumn() <= 13, 'SETUP', 'N列に既存データがあります。自動移行を中止しました。');
    sheet.getRange(1, 14).setValues([['asset_account_id']]);
    actual = expected;
  }
  require_(expected.every(function (h, i) { return actual[i] === h; }), 'SETUP', name + ' の既存ヘッダーが異なります。データは変更していません。');
}

function seedMasters_(book) {
  // Deterministic initial IDs let interrupted initialization resume safely.
  var now = timestamp_();
  var properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('MASTERS_SEEDED') === '1') return;
  PAYMENT_DEFAULTS_.forEach(function (name, i) {
    var id = 'initial-payment-' + (i + 1);
    var rows = records_(book, 'PaymentMethods');
    if (!rows.some(function (r) { return r.payment_method_id === id || r.name === name; })) {
      appendRecord_(book, 'PaymentMethods', { payment_method_id: id, name: name, sort_order: (i + 1) * 10, is_active: true, created_at: now, updated_at: now });
    }
  });
  Object.keys(CATEGORY_DEFAULTS_).forEach(function (type) {
    CATEGORY_DEFAULTS_[type].forEach(function (name, i) {
      var id = 'initial-' + type + '-' + (i + 1);
      var rows = records_(book, 'Categories');
      if (!rows.some(function (r) { return r.category_id === id || (r.type === type && r.name === name); })) {
        appendRecord_(book, 'Categories', { category_id: id, type: type, name: name, sort_order: (i + 1) * 10, is_active: true, created_at: now, updated_at: now });
      }
    });
  });
  SpreadsheetApp.flush();
  properties.setProperty('MASTERS_SEEDED', '1');
}

function records_(book, name) {
  var sheet = book.getSheetByName(name);
  if (sheet.getLastRow() < 2) return [];
  var headers = HEADERS_[name];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues().map(function (values, index) {
    var record = { _row: index + 2 };
    headers.forEach(function (key, i) {
      var value = values[i];
      if (value instanceof Date) value = Utilities.formatDate(value, TZ_, key === 'date' ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ss.SSS") + (key === 'date' ? '' : '+09:00');
      if (key === 'is_active') value = value === true || value === 'TRUE' || value === 'true';
      record[key] = value;
    });
    return record;
  }).filter(function (r) { return r[headers[0]] !== ''; });
}

function writeRecord_(book, name, rowNumber, record) {
  var sheet = book.getSheetByName(name);
  var values = HEADERS_[name].map(function (key) {
    var value = record[key] === undefined ? '' : record[key];
    // An apostrophe is Sheets' literal text escape; getValues returns the text.
    // Escape all user-provided names/notes, including leading apostrophes.
    return (key === 'name' || key === 'note' || key === 'category_name' || key === 'payment_method_name') && value !== '' ? "'" + value : value;
  });
  HEADERS_[name].forEach(function (key, i) {
    if (key !== 'amount' && key !== 'sort_order' && key !== 'is_active') sheet.getRange(rowNumber, i + 1).setNumberFormat('@');
  });
  sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
}

function appendRecord_(book, name, record) { writeRecord_(book, name, book.getSheetByName(name).getLastRow() + 1, record); }
function publicRecord_(record) { var result = Object.assign({}, record); delete result._row; return result; }
function masters_(book, name) { return records_(book, name).map(publicRecord_).sort(function (a, b) { return a.sort_order - b.sort_order || a.name.localeCompare(b.name, 'ja'); }); }
function displayTransaction_(row, book) { return displayWithMasters_(row, masters_(book, 'PaymentMethods'), masters_(book, 'Categories')); }
function displayWithMasters_(row, methods, categories) {
  var result = publicRecord_(row);
  var method = methods.find(function (m) { return m.payment_method_id === row.payment_method_id; });
  var category = categories.find(function (c) { return c.category_id === row.category_id; });
  result.payment_method_name = method ? method.name : row.payment_method_name;
  result.category_name = category ? category.name : (row.category_name || '未分類');
  return result;
}

function transactionById_(book, id) {
  validId_(id, '取引ID');
  var row = records_(book, 'Transactions').find(function (r) { return r.transaction_id === id; });
  require_(row, 'NOT_FOUND', '取引が見つかりません。');
  return row;
}

function validateTransaction_(input, book, old) {
  validDate_(input.date);
  validType_(input.type);
  require_(Number.isSafeInteger(input.amount) && input.amount >= 1, 'VALIDATION', '金額は1円以上の安全に扱える整数で入力してください。');
  require_(typeof input.note === 'string' && input.note.length <= 10000, 'VALIDATION', '備考は10,000文字以内で入力してください。');
  validId_(input.payment_method_id, '支払・入金方法');
  validId_(input.category_id, 'カテゴリ');
  var method = masters_(book, 'PaymentMethods').find(function (m) { return m.payment_method_id === input.payment_method_id; });
  var category = masters_(book, 'Categories').find(function (c) { return c.category_id === input.category_id; });
  require_(method && (method.is_active || (old && old.payment_method_id === input.payment_method_id)), 'VALIDATION', '有効な支払・入金方法を選択してください。');
  require_(category && category.type === input.type && (category.is_active || (old && old.category_id === input.category_id && old.type === input.type)), 'VALIDATION', '収支区分に対応する有効なカテゴリを選択してください。');
  var link = records_(book, 'PaymentAccountLinks').find(function (r) { return r.payment_method_id === method.payment_method_id; });
  require_(link || (old && old.payment_method_id === method.payment_method_id && old.asset_account_id), 'VALIDATION', '資産管理タブでこの支払・入金方法の残高反映先を設定してください。');
  return { date: input.date, type: input.type, payment_method_id: method.payment_method_id, payment_method_name: method.name, category_id: category.category_id, category_name: category.name, amount: input.amount, note: input.note,
    asset_account_id: old && old.payment_method_id === method.payment_method_id && old.asset_account_id ? old.asset_account_id : (link ? link.account_id : '') };
}

function seedAssetMasters_(book) {
  ACCOUNT_DEFAULTS_.forEach(function (item, i) {
    if (!records_(book, 'AssetAccounts').some(function (r) { return r.account_id === item[0]; })) {
      appendRecord_(book, 'AssetAccounts', { account_id: item[0], name: item[1], kind: item[2], sort_order: i });
    }
  });
  var properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('ASSET_MASTERS_SEEDED') === '1') return;
  var targets = ['asset-cash', 'debt-jal-card', 'debt-jre-card', 'asset-suica', 'asset-yucho', 'asset-jrebank', 'asset-paypay'];
  masters_(book, 'PaymentMethods').forEach(function (method) {
    var index = PAYMENT_DEFAULTS_.indexOf(method.name);
    // Stable initial IDs also survive master renaming in an existing installation.
    if (/^initial-payment-[1-7]$/.test(method.payment_method_id)) index = Number(method.payment_method_id.slice(-1)) - 1;
    if (index >= 0 && !records_(book, 'PaymentAccountLinks').some(function (r) { return r.payment_method_id === method.payment_method_id; })) {
      appendRecord_(book, 'PaymentAccountLinks', { payment_method_id: method.payment_method_id, account_id: targets[index], updated_at: timestamp_() });
    }
  });
  var links = records_(book, 'PaymentAccountLinks');
  records_(book, 'Transactions').forEach(function (transaction) {
    var link = links.find(function (r) { return r.payment_method_id === transaction.payment_method_id; });
    if (!transaction.asset_account_id && link) book.getSheetByName('Transactions').getRange(transaction._row, 14).setNumberFormat('@').setValues([[link.account_id]]);
  });
  SpreadsheetApp.flush();
  properties.setProperty('ASSET_MASTERS_SEEDED', '1');
}

function assetContext_(book) {
  return { accounts: masters_(book, 'AssetAccounts'), snapshots: records_(book, 'AssetSnapshots'), transactions: records_(book, 'Transactions'), movements: records_(book, 'AssetMovements') };
}

function assetRevision_(context) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(context));
  return bytes.map(function (b) { return (b & 255).toString(16).padStart(2, '0'); }).join('');
}

function assetEffects_(context, start, end) {
  var effects = {};
  context.accounts.forEach(function (a) { effects[a.account_id] = 0; });
  function add(id, amount) { if (Object.prototype.hasOwnProperty.call(effects, id)) effects[id] = safeAdd_(effects[id], amount); }
  context.transactions.forEach(function (t) {
    if (t.deleted_at || t.date < start || t.date > end || !t.asset_account_id) return;
    var account = context.accounts.find(function (a) { return a.account_id === t.asset_account_id; });
    if (account) add(account.account_id, t.amount * (t.type === 'income' ? 1 : -1) * (account.kind === 'asset' ? 1 : -1));
  });
  context.movements.forEach(function (m) {
    if (m.deleted_at || m.date < start || m.date > end) return;
    var from = context.accounts.find(function (a) { return a.account_id === m.from_account_id; });
    var to = context.accounts.find(function (a) { return a.account_id === m.to_account_id; });
    if (from) add(from.account_id, m.amount * (from.kind === 'asset' ? -1 : 1));
    if (to) add(to.account_id, m.amount * (to.kind === 'asset' ? 1 : -1));
  });
  return effects;
}

function balancesAt_(context, date) {
  var snapshots = context.snapshots.filter(function (s) { return s.date <= date; }).sort(function (a, b) { return b.date.localeCompare(a.date) || b._row - a._row; });
  var totalAssets = 0, totalLiabilities = 0;
  var accounts = context.accounts.map(function (a) {
    var base = snapshots.find(function (s) { return Object.prototype.hasOwnProperty.call(JSON.parse(s.balances_json), a.account_id); });
    var amount = null;
    if (base) {
      var baseline = JSON.parse(base.balances_json)[a.account_id];
      var included = JSON.parse(base.included_effects_json)[a.account_id] || 0;
      amount = safeAdd_(safeAdd_(baseline, assetEffects_(context, base.date, date)[a.account_id]), -included);
      if (a.kind === 'asset') totalAssets = safeAdd_(totalAssets, amount);
      else totalLiabilities = safeAdd_(totalLiabilities, amount);
    }
    return Object.assign({}, a, { amount: amount, confirmed_date: base ? base.date : null });
  });
  return { date: date, known: accounts.every(function (a) { return a.amount !== null; }), accounts: accounts,
    assets: totalAssets, liabilities: totalLiabilities, netAssets: safeAdd_(totalAssets, -totalLiabilities) };
}

function initializeAssets(input) {
  return api_(function () {
    return locked_(function () {
      var book = database_(), context = assetContext_(book);
      validSnapshotInput_(input, context.accounts, true);
      var duplicate = context.snapshots.find(function (s) { return s.request_id === input.request_id; });
      if (duplicate) { sameSnapshot_(duplicate, input); return { snapshot_id: duplicate.snapshot_id, duplicate: true }; }
      require_(!context.snapshots.length, 'CONFLICT', '開始残高は登録済みです。資産管理から残高を変更してください。');
      return appendSnapshot_(book, context, input);
    });
  });
}

function saveAssetBalances(input) {
  return api_(function () {
    return locked_(function () {
      var book = database_(), context = assetContext_(book);
      validSnapshotInput_(input, context.accounts, false);
      var duplicate = context.snapshots.find(function (s) { return s.request_id === input.request_id; });
      if (duplicate) { sameSnapshot_(duplicate, input); return { snapshot_id: duplicate.snapshot_id, duplicate: true }; }
      require_(context.snapshots.length, 'SETUP', '先に開始残高を登録してください。');
      var first = context.snapshots.reduce(function (date, s) { return s.date < date ? s.date : date; }, context.snapshots[0].date);
      require_(input.date >= first, 'VALIDATION', '開始残高の日付以降を指定してください。');
      require_(input.expected_revision === assetRevision_(context), 'CONFLICT', '残高または取引が更新されました。資産管理を更新してから変更してください。');
      return appendSnapshot_(book, context, input);
    });
  });
}

function validSnapshotInput_(input, accounts, all) {
  require_(input && typeof input === 'object', 'VALIDATION', '残高を指定してください。');
  validId_(input.request_id, 'リクエストID'); validDate_(input.date);
  require_(input.date <= today_(), 'VALIDATION', '残高確認日は今日以前を指定してください。');
  require_(input.balances && typeof input.balances === 'object' && !Array.isArray(input.balances), 'VALIDATION', '残高の形式が不正です。');
  var ids = Object.keys(input.balances);
  require_(ids.length > 0 && (!all || ids.length === accounts.length), 'VALIDATION', '全項目の開始残高を入力してください。ない項目は0円です。');
  ids.forEach(function (id) {
    require_(accounts.some(function (a) { return a.account_id === id; }) && Number.isSafeInteger(input.balances[id]) && input.balances[id] >= 0, 'VALIDATION', '残高は0円以上の整数で入力してください。');
  });
  // Fail before writing if the proposed asset, liability or net totals cannot be represented.
  var assets = 0, liabilities = 0;
  accounts.forEach(function (a) { if (Object.prototype.hasOwnProperty.call(input.balances, a.account_id)) { if (a.kind === 'asset') assets = safeAdd_(assets, input.balances[a.account_id]); else liabilities = safeAdd_(liabilities, input.balances[a.account_id]); } });
  safeAdd_(assets, -liabilities);
}

function sameSnapshot_(snapshot, input) {
  var saved = JSON.parse(snapshot.balances_json);
  require_(snapshot.date === input.date && Object.keys(saved).length === Object.keys(input.balances).length && Object.keys(saved).every(function (id) { return saved[id] === input.balances[id]; }), 'CONFLICT', '同じリクエストIDの残高内容が異なります。');
}

function appendSnapshot_(book, context, input) {
  var balances = {}, included = {}, effects = assetEffects_(context, input.date, input.date);
  context.accounts.forEach(function (a) { if (Object.prototype.hasOwnProperty.call(input.balances, a.account_id)) { balances[a.account_id] = input.balances[a.account_id]; included[a.account_id] = effects[a.account_id]; } });
  var snapshot = { snapshot_id: Utilities.getUuid(), date: input.date, balances_json: JSON.stringify(balances), included_effects_json: JSON.stringify(included), created_at: timestamp_(), request_id: input.request_id };
  var proposed = Object.assign({}, context, { snapshots: context.snapshots.concat([Object.assign({ _row: book.getSheetByName('AssetSnapshots').getLastRow() + 1 }, snapshot)]) });
  balancesAt_(proposed, today_()); // Validate derived totals before committing any row.
  appendRecord_(book, 'AssetSnapshots', snapshot); SpreadsheetApp.flush();
  return { snapshot_id: snapshot.snapshot_id, duplicate: false };
}

function getAssets(input) {
  return api_(function () {
    input = input || {};
    var date = input.date || today_(); validDate_(date);
    var context = assetContext_(database_());
    var current = balancesAt_(context, date), months = [];
    var index = Number(date.slice(0, 4)) * 12 + Number(date.slice(5, 7)) - 1;
    for (var offset = 11; offset >= 0; offset--) {
      var n = index - offset, year = Math.floor(n / 12), month = n % 12 + 1;
      var monthKey = String(year).padStart(4, '0') + '-' + String(month).padStart(2, '0');
      var last = new Date(Date.UTC(year, month, 0)).getUTCDate();
      var end = monthKey === date.slice(0, 7) ? date : monthKey + '-' + String(last).padStart(2, '0');
      var values = balancesAt_(context, end);
      months.push({ month: monthKey, date: end, known: values.known, assets: values.known ? values.assets : null, liabilities: values.known ? values.liabilities : null, netAssets: values.known ? values.netAssets : null });
    }
    var movements = context.movements.filter(function (m) { return !m.deleted_at && m.date <= date; }).sort(function (a, b) { return b.date.localeCompare(a.date) || b._row - a._row; });
    return Object.assign(current, { initialized: context.snapshots.length > 0, revision: assetRevision_(context), months: months,
      movements: movements.slice(0, 50).map(publicRecord_), movementCount: movements.length,
      unmappedTransactions: context.transactions.filter(function (t) { return !t.deleted_at && !t.asset_account_id && t.date <= date; }).length });
  });
}

function savePaymentAccountLink(input) {
  return api_(function () {
    return locked_(function () {
      var book = database_();
      require_(input && typeof input === 'object', 'VALIDATION', '反映先を指定してください。');
      var method = masters_(book, 'PaymentMethods').find(function (m) { return m.payment_method_id === input.payment_method_id; });
      var account = masters_(book, 'AssetAccounts').find(function (a) { return a.account_id === input.account_id; });
      require_(method && account, 'VALIDATION', '支払方法と反映先を選択してください。');
      var old = records_(book, 'PaymentAccountLinks').find(function (r) { return r.payment_method_id === input.payment_method_id; });
      if (old && old.account_id === input.account_id) return publicRecord_(old);
      require_((old ? old.updated_at : '') === input.expected_updated_at, 'CONFLICT', '反映先が更新されました。読み直してください。');
      var row = { payment_method_id: input.payment_method_id, account_id: input.account_id, updated_at: timestamp_(old && old.updated_at) };
      if (old) writeRecord_(book, 'PaymentAccountLinks', old._row, row); else appendRecord_(book, 'PaymentAccountLinks', row);
      SpreadsheetApp.flush(); return row;
    });
  });
}

function createAssetMovement(input) {
  return api_(function () {
    return locked_(function () {
      var book = database_(), context = assetContext_(book);
      require_(context.snapshots.length, 'SETUP', '先に開始残高を登録してください。');
      require_(input && typeof input === 'object', 'VALIDATION', '移動内容を指定してください。');
      validId_(input.request_id, 'リクエストID'); validDate_(input.date);
      var old = context.movements.find(function (m) { return m.request_id === input.request_id; });
      if (old) { require_(['date', 'from_account_id', 'to_account_id', 'amount'].every(function (key) { return old[key] === input[key]; }), 'CONFLICT', '同じリクエストIDの移動内容が異なります。'); return { movement_id: old.movement_id, duplicate: true }; }
      require_(input.from_account_id !== input.to_account_id && context.accounts.some(function (a) { return a.account_id === input.from_account_id; }) && context.accounts.some(function (a) { return a.account_id === input.to_account_id; }), 'VALIDATION', '異なる移動元・移動先を選択してください。');
      require_(Number.isSafeInteger(input.amount) && input.amount >= 1, 'VALIDATION', '移動金額は1円以上の整数で入力してください。');
      var atDate = balancesAt_(context, input.date);
      require_(atDate.known && input.date <= today_(), 'VALIDATION', '開始残高の日付から今日までの移動日を指定してください。');
      var movement = { movement_id: Utilities.getUuid(), date: input.date, from_account_id: input.from_account_id, to_account_id: input.to_account_id, amount: input.amount, created_at: timestamp_(), deleted_at: '', request_id: input.request_id };
      var proposed = Object.assign({}, context, { movements: context.movements.concat([movement]) });
      balancesAt_(proposed, today_());
      appendRecord_(book, 'AssetMovements', movement); SpreadsheetApp.flush(); return { movement_id: movement.movement_id, duplicate: false };
    });
  });
}

function deleteAssetMovement(input) {
  return api_(function () {
    return locked_(function () {
      var book = database_();
      require_(input && typeof input === 'object', 'VALIDATION', '移動を指定してください。'); validId_(input.movement_id, '移動ID');
      var movement = records_(book, 'AssetMovements').find(function (m) { return m.movement_id === input.movement_id; });
      require_(movement, 'NOT_FOUND', '移動が見つかりません。');
      if (!movement.deleted_at) { movement.deleted_at = timestamp_(); writeRecord_(book, 'AssetMovements', movement._row, movement); SpreadsheetApp.flush(); }
      return { movement_id: movement.movement_id, deleted: true };
    });
  });
}

function validId_(id, label) { require_(typeof id === 'string' && /^[A-Za-z0-9_-]{8,100}$/.test(id), 'VALIDATION', label + 'を正しく指定してください。'); }
function validType_(type) { require_(type === 'income' || type === 'expense', 'VALIDATION', '収支区分が不正です。'); }
function validDate_(value) {
  require_(typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value), 'VALIDATION', '日付をYYYY-MM-DD形式で入力してください。');
  var parts = value.split('-').map(Number);
  var y = parts[0], m = parts[1], d = parts[2];
  var leap = y % 4 === 0 && (y % 100 !== 0 || y % 400 === 0);
  var days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  require_(y >= 1900 && y <= 9999 && m >= 1 && m <= 12 && d >= 1 && d <= days[m - 1], 'VALIDATION', '存在する日付を選択してください（1900〜9999年）。');
}
function validMonth_(value) { require_(typeof value === 'string' && /^\d{4}-\d{2}$/.test(value), 'VALIDATION', '対象年月が不正です。'); validDate_(value + '-01'); }
function safeAdd_(a, b) { require_(Number.isSafeInteger(b) && Number.isSafeInteger(a + b), 'DATA', '集計金額が安全に扱える整数の範囲を超えています。'); return a + b; }
function today_() { return Utilities.formatDate(new Date(), TZ_, 'yyyy-MM-dd'); }
function timestamp_(previous) {
  var now = new Date();
  if (previous && now.getTime() <= new Date(previous).getTime()) now = new Date(new Date(previous).getTime() + 1);
  return Utilities.formatDate(now, TZ_, "yyyy-MM-dd'T'HH:mm:ss.SSS") + '+09:00';
}
