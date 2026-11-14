// Émulation minimale de Google Apps Script (SpreadsheetApp, Utilities, ...) pour tester
// apps-script/*.gs en local avec Node, sans compte Google.
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const crypto = require('crypto');

const toSigned = (buf) => Array.from(buf, (b) => (b > 127 ? b - 256 : b));
const toBuf = (v) => (Array.isArray(v) ? Buffer.from(v.map((b) => b & 0xff)) : Buffer.from(String(v), 'utf8'));

function colToNum(s) { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }

class Chain { constructor() { return new Proxy(this, { get: (t, p, recv) => (p in t ? t[p] : () => recv) }); } }

class Range extends Chain {
  constructor(sheet, r, c, nr, nc) { super(); Object.assign(this, { sheet, r, c, nr, nc }); }
  getRow() { return this.r; }
  getColumn() { return this.c; }
  getNumRows() { return this.nr; }
  getSheet() { return this.sheet; }
  getValues() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) row.push(this.sheet._value(this.r + i, this.c + j));
      out.push(row);
    }
    return out;
  }
  getFormulas() {
    const out = [];
    for (let i = 0; i < this.nr; i++) {
      const row = [];
      for (let j = 0; j < this.nc; j++) row.push(this.sheet._cell(this.r + i, this.c + j).f || '');
      out.push(row);
    }
    return out;
  }
  setValues(v) {
    if (v.length !== this.nr || v.some((row) => row.length !== this.nc)) throw new Error(`setValues: dimensions ${v.length}x${v[0] && v[0].length} ≠ ${this.nr}x${this.nc}`);
    v.forEach((row, i) => row.forEach((x, j) => this.sheet._set(this.r + i, this.c + j, { v: x === null || x === undefined ? '' : x })));
    return this;
  }
  setValue(x) { return this.setValues([[x]]); }
  setFormulas(f) { f.forEach((row, i) => row.forEach((x, j) => this.sheet._set(this.r + i, this.c + j, { v: '', f: x }))); return this; }
  clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sheet._set(this.r + i, this.c + j, { v: '' }); return this; }
}

class Sheet {
  constructor(ss, name, rows) { this.ss = ss; this.name = name; this.rows = (rows || []).map((r) => r.map((v) => ({ v }))); this.maxRows = 1000; }
  getName() { return this.name; }
  _cell(r, c) { return (this.rows[r - 1] && this.rows[r - 1][c - 1]) || { v: '' }; }
  _set(r, c, cell) {
    while (this.rows.length < r) this.rows.push([]);
    const row = this.rows[r - 1];
    while (row.length < c) row.push({ v: '' });
    row[c - 1] = cell;
  }
  _value(r, c) { const cell = this._cell(r, c); return cell.f ? this.ss._eval(this, r, cell.f) : cell.v; }
  getLastRow() {
    for (let i = this.rows.length; i > 0; i--) if (this.rows[i - 1].some((x) => x.f || (x.v !== '' && x.v !== null))) return i;
    return 0;
  }
  getLastColumn() { return Math.max(0, ...this.rows.map((r) => r.length)); }
  getMaxRows() { return this.maxRows; }
  getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const m = a.match(/^([A-Z]+)(\d*)(?::([A-Z]+)(\d*))?$/);
      const c1 = colToNum(m[1]), r1 = Number(m[2] || 1), c2 = m[3] ? colToNum(m[3]) : c1;
      return new Range(this, r1, c1, 1, c2 - c1 + 1); // utilisé seulement pour la mise en forme
    }
    return new Range(this, a, b, c || 1, d || 1);
  }
  _shift(at, delta) {
    for (const row of this.rows) for (const cell of row) {
      if (cell.f) cell.f = cell.f.replace(/(^|[^A-Z!$])(\$?)([A-Z]{1,2})(\$?)(\d+)/g, (m0, pre, d1, col, d2, n) => {
        n = Number(n);
        return pre + d1 + col + d2 + (n >= at ? n + delta : n);
      });
    }
  }
  insertRowBefore(r) { this._shift(r, 1); this.rows.splice(r - 1, 0, []); }
  deleteRow(r) { this.rows.splice(r - 1, 1); this._shift(r + 1, -1); }
  appendRow(vals) { this.getRange(this.getLastRow() + 1, 1, 1, vals.length).setValues([vals]); }
  setFrozenRows() {} setColumnWidth() {}
}

class Spreadsheet {
  constructor(data) { this.sheets = Object.entries(data).map(([n, rows]) => new Sheet(this, n, rows)); }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  insertSheet(n) { const s = new Sheet(this, n, []); this.sheets.push(s); return s; }
  getSheets() { return this.sheets; }
  // Évalue uniquement les deux formules posées par le script
  _eval(sheet, row, f) {
    const val = (col) => sheet._value(row, colToNum(col));
    let m = f.match(/^=SUMIFS\(DONS!\$E:\$E,DONS!\$C:\$C,\$A(\d+),DONS!\$D:\$D,\$H(\d+),DONS!\$K:\$K,"<>OUI"\)$/);
    if (m) {
      const id = sheet._value(Number(m[1]), 1), cyc = sheet._value(Number(m[2]), 8);
      const d = this.getSheetByName('DONS');
      let s = 0;
      for (let r = 2; r <= d.getLastRow(); r++) {
        const v = (c) => d._value(r, c);
        if (String(v(3)).toLowerCase() === String(id).toLowerCase() && String(v(4)) === String(cyc) && String(v(11)).toUpperCase() !== 'OUI') s += Number(v(5)) || 0;
      }
      return s;
    }
    m = f.match(/^=IF\(OR\(\$D(\d+)="",\$E\d+>=\$D\d+\),"-",\$D\d+-\$E\d+\)$/);
    if (m) {
      const r = Number(m[1]); const b = sheet._value(r, 4), e = sheet._value(r, 5);
      return b === '' || e >= b ? '-' : b - e;
    }
    throw new Error('Formule non gérée par le mock : ' + f);
  }
}

function createContext(data, opts = {}) {
  const ss = new Spreadsheet(data);
  const props = {}, cache = {};
  const user = { email: opts.ownerEmail || 'owner@example.org' };
  const ctx = {
    console: opts.quiet ? { log() {}, warn() {}, error() {} } : console,
    SpreadsheetApp: {
      getActive: () => ss, getActiveSpreadsheet: () => ss, flush() {},
      getUi() { throw new Error('Pas d\'interface dans les tests'); },
      newDataValidation: () => new Chain(),
    },
    Utilities: {
      DigestAlgorithm: { SHA_256: 'sha256' }, Charset: { UTF_8: 'utf8' },
      computeDigest: (alg, s) => toSigned(crypto.createHash(alg).update(toBuf(s)).digest()),
      computeHmacSha256Signature: (v, k) => toSigned(crypto.createHmac('sha256', toBuf(k)).update(toBuf(v)).digest()),
      base64EncodeWebSafe: (v) => toBuf(v).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
      base64DecodeWebSafe: (s) => toSigned(Buffer.from(String(s).replace(/-/g, '+').replace(/_/g, '/'), 'base64')),
      newBlob: (bytes) => ({ getDataAsString: () => toBuf(bytes).toString('utf8') }),
      getUuid: () => crypto.randomUUID(),
    },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] || null, setProperty: (k, v) => { props[k] = v; } }) },
    CacheService: { getScriptCache: () => ({ get: (k) => cache[k] || null, put: (k, v) => { cache[k] = v; }, remove: (k) => { delete cache[k]; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => user.email }), getEffectiveUser: () => ({ getEmail: () => user.email }) },
    ContentService: { MimeType: { JSON: 'json' }, createTextOutput: (s) => ({ setMimeType() { return this; }, getContent: () => s }) },
  };
  vm.createContext(ctx);
  const dir = path.join(__dirname, '..', 'apps-script');
  for (const f of fs.readdirSync(dir).filter((f) => f.endsWith('.gs')).sort()) {
    vm.runInContext(fs.readFileSync(path.join(dir, f), 'utf8'), ctx, { filename: f });
  }
  ctx.__ss = ss;
  ctx.api = (body) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify(body) } }).getContent());
  return ctx;
}

module.exports = { createContext };
