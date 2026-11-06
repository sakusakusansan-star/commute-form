// 実行: node apps-script/commute-history.test.js
// シートをモックにして、社員の稼働費GAS（commute-expense-sheet.gs）の「稼働費履歴」を確かめる
const fs = require('fs'), path = require('path'), assert = require('assert');

function chain(target) {
  return new Proxy(target, { get: (t, k) => (k in t ? t[k] : () => chain(t)) });
}
function Sheet(name, parent) {
  const g = [];
  const get = (r, c) => (g[r - 1] && g[r - 1][c - 1] !== undefined) ? g[r - 1][c - 1] : '';
  const set = (r, c, v) => { while (g.length < r) g.push([]); g[r - 1][c - 1] = v; };
  const range = (r, c, nr, nc) => chain({
    getValue: () => get(r, c),
    getValues: () => Array.from({ length: nr || 1 }, (_, i) => Array.from({ length: nc || 1 }, (_, j) => get(r + i, c + j))),
    setValues(v) { v.forEach((row, i) => row.forEach((x, j) => set(r + i, c + j, x))); return this; },
    setValue(v) { set(r, c, v); return this; }
  });
  return chain({
    _g: g, _set: set, getName: () => name, getParent: () => parent,
    getRange: (a, c, nr, nc) => range(a, c, nr, nc),
    getLastRow: () => { let n = g.length; while (n && !(g[n - 1] || []).some(x => x !== '' && x !== undefined)) n--; return n; },
    appendRow(row) { const r = this.getLastRow() + 1; row.forEach((x, j) => set(r, j + 1, x)); },
    deleteRow: r => g.splice(r - 1, 1)
  });
}
const sheets = {};
const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = Sheet(n, ss)) };
global.SpreadsheetApp = { getActiveSpreadsheet: () => ss, flush() {}, BorderStyle: { SOLID: 1 } };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
global.ContentService = { MimeType: { JSON: 'json' }, createTextOutput: t => ({ t, setMimeType() { return this; } }) };
global.PropertiesService = { getScriptProperties: () => ({ getProperty: () => null }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put() {} }) };
(0, eval)(fs.readFileSync(path.join(__dirname, 'commute-expense-sheet.gs'), 'utf8'));

// 福山さんのシート：10月のカレンダーに 10/1・10/3 とイレギュラー1件
const s = ss.insertSheet('福山康平');
s._set(2, 7, 10);
s._set(6, 1, new Date(2026, 9, 1));
s._set(6, 3, 'au・SB下通'); s._set(6, 7, 500); s._set(6, 8, 500); s._set(6, 9, '有'); s._set(6, 10, 1000);
s._set(8, 3, 'イオン'); s._set(8, 9, '無'); s._set(8, 10, 320);
s._set(38, 1, new Date(2026, 9, 12)); s._set(38, 3, '自宅→A→B→自宅'); s._set(38, 10, 500); s._set(38, 11, '応援');
const post = q => JSON.parse(doPost({ postData: { contents: JSON.stringify(q) } }).t);

// 履歴が無いうちは found: false
let r = post({ action: 'getHistory', staffName: '福山康平', ym: '2026-10' });
assert.strictEqual(r.ok, true); assert.strictEqual(r.data.found, false);
assert.strictEqual(post({ action: 'getHistory', staffName: '福山康平', ym: '10月' }).ok, false, '年月の形が違う');
assert.strictEqual(post({ action: 'getHistory', staffName: 'だれか', ym: '2026-10' }).ok, false, '名簿にいない人');

// 写す → 読める（getMonthInfo・getMyEntries と同じ形）
assert.strictEqual(archiveMonth_('福山康平', s), 3);
r = post({ action: 'getHistory', staffName: '福山康平', ym: '2026-10' }).data;
assert.strictEqual(r.found, true);
assert.deepStrictEqual([r.year, r.month, r.filledDays, r.irregularCount], [2026, 10, [1, 3], 1]);
assert.deepStrictEqual(JSON.parse(JSON.stringify(r.entries)), [
  { type: 'normal', day: 1, place: 'au・SB下通', highway: '有', tollGo: 500, tollBack: 500, amount: 1000 },
  { type: 'normal', day: 3, place: 'イオン', highway: '無', tollGo: 0, tollBack: 0, amount: 320 },
  { type: 'irregular', day: 12, content: '自宅→A→B→自宅', amount: 500, note: '応援' }]);

// もう一度写しても増えない（入れ直し）。金額を直してから写すと新しい値になる
s._set(8, 10, 400);
archiveMonth_('福山康平', s);
r = getHistoryWeb('福山康平', '2026-10');
assert.strictEqual(r.entries.length, 3);
assert.strictEqual(r.entries[1].amount, 400);
// 年月のセルが日付に変わっていても読める
sheets['稼働費履歴']._g[1][0] = new Date(2026, 9, 1);
assert.strictEqual(getHistoryWeb('福山康平', '2026-10').entries.length, 3);
assert.strictEqual(getHistoryWeb('福山康平', '2026-09').found, false, 'ほかの月は出ない');

// 月を切り替える（resetMonth）と、作り直す前の10月が履歴に残る。空の11月を写しても10月は消えない
let rebuilt = 0;
global.buildMonthRows = (sh, m, y) => { rebuilt++; for (let row = 6; row <= 40; row++) for (let c = 1; c <= 11; c++) sh._set(row, c, ''); sh._set(6, 1, new Date(y, m - 1, 1)); };
global.setupDropdowns = () => {};
sheets['稼働費履歴']._g.splice(1);   // 履歴を空に
s._set(8, 10, 320);
assert.strictEqual(post({ action: 'resetMonth', staffName: '福山康平', month: 11 }).ok, true);
assert.strictEqual(rebuilt, 1);
assert.strictEqual(getHistoryWeb('福山康平', '2026-10').entries.length, 3, '切り替える前の10月が残る');
assert.strictEqual(archiveMonth_('福山康平', s), 0, '空の月は写さない');
assert.strictEqual(getHistoryWeb('福山康平', '2026-10').entries.length, 3);
assert.strictEqual(getHistoryWeb('福山康平', '2026-11').found, false);
assert.strictEqual(getStaffInBook_().indexOf('稼働費履歴'), -1, '履歴シートは名前の一覧に出ない');
console.log('commute-history: all tests passed');
