// Local Apps Script service doubles. These are NOT Google integration tests.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { randomUUID } = require('node:crypto');

class Sheet {
  constructor(name) { this.name = name; this.rows = []; this.formats = new Map(); this.formulas = new Map(); }
  getLastRow() { let end = this.rows.length; while (end && !this.rows[end - 1]?.some(v => v !== '' && v !== undefined)) end--; return end; }
  getLastColumn() { return Math.max(0, ...this.rows.map(row => row.reduce((last, v, i) => v !== '' && v !== undefined ? i + 1 : last, 0))); }
  setFrozenRows(count) { this.frozenRows = count; }
  getRange(row, col, height = 1, width = 1) {
    const sheet = this;
    return {
      getValues() { return Array.from({ length: height }, (_, r) => Array.from({ length: width }, (_, c) => sheet.rows[row - 1 + r]?.[col - 1 + c] ?? '')); },
      setNumberFormat(format) {
        for (let r = 0; r < height; r++) for (let c = 0; c < width; c++) sheet.formats.set(`${row + r}:${col + c}`, format);
        return this;
      },
      setValues(values) {
        if (values.length !== height || values.some(line => line.length !== width)) throw new Error('Invalid range size');
        values.forEach((line, r) => {
          const target = sheet.rows[row - 1 + r] ||= [];
          line.forEach((value, c) => {
            const key = `${row + r}:${col + c}`;
            const escaped = typeof value === 'string' && value.startsWith("'");
            sheet.formulas.delete(key);
            if (typeof value === 'string' && value.startsWith('=') && !escaped) sheet.formulas.set(key, value);
            target[col - 1 + c] = escaped ? value.slice(1) : value;
          });
        });
        return this;
      },
      getFormulas() { return Array.from({ length: height }, (_, r) => Array.from({ length: width }, (_, c) => sheet.formulas.get(`${row + r}:${col + c}`) || '')); }
    };
  }
}

class Book {
  constructor(id) { this.id = id; this.sheets = new Map(); this.timeZone = null; }
  getId() { return this.id; }
  getSheetByName(name) { return this.sheets.get(name) || null; }
  insertSheet(name) { if (this.sheets.has(name)) throw new Error('Sheet exists'); const sheet = new Sheet(name); this.sheets.set(name, sheet); return sheet; }
  setSpreadsheetTimeZone(value) { this.timeZone = value; }
}

function formatDate(date, zone, format) {
  const parts = new Intl.DateTimeFormat('en-US', { timeZone: zone, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }).formatToParts(date);
  const get = type => parts.find(p => p.type === type).value;
  const day = `${get('year')}-${get('month')}-${get('day')}`;
  return format === 'yyyy-MM-dd' ? day : `${day}T${get('hour')}:${get('minute')}:${get('second')}.${String(date.getUTCMilliseconds()).padStart(3, '0')}`;
}

function createRuntime({ initialized = true } = {}) {
  const properties = new Map(), books = new Map();
  const runtime = { activeEmail: 'owner@example.test', effectiveEmail: 'owner@example.test', busy: false, lockHeld: false, lockCount: 0, flushFailure: false, errors: [] };
  const context = vm.createContext({
    Date,
    console: { error: value => runtime.errors.push(String(value)) },
    Utilities: { getUuid: randomUUID, formatDate },
    Session: { getActiveUser: () => ({ getEmail: () => runtime.activeEmail }), getEffectiveUser: () => ({ getEmail: () => runtime.effectiveEmail }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: key => properties.get(key) || null, setProperty: (key, value) => properties.set(key, String(value)) }) },
    LockService: { getScriptLock: () => ({ tryLock: () => { if (runtime.busy || runtime.lockHeld) return false; runtime.lockHeld = true; runtime.lockCount++; return true; }, releaseLock: () => { runtime.lockHeld = false; } }) },
    SpreadsheetApp: {
      create: () => { const book = new Book(randomUUID()); books.set(book.id, book); return book; },
      openById: id => { if (!books.has(id)) throw new Error('Spreadsheet not found'); return books.get(id); },
      flush: () => { if (runtime.flushFailure) { runtime.flushFailure = false; throw new Error('Simulated interruption after write'); } }
    }
  });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), context, { filename: 'Code.gs' });
  runtime.context = context; runtime.properties = properties; runtime.books = books;
  runtime.invoke = (method, argument) => {
    const result = context[method](argument);
    return result === undefined ? undefined : JSON.parse(JSON.stringify(result));
  };
  runtime.book = () => books.get(properties.get('SPREADSHEET_ID'));
  if (initialized) runtime.invoke('setupDatabase');
  return runtime;
}

module.exports = { createRuntime, Book, Sheet, formatDate };
