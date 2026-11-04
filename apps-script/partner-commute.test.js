// 実行: node apps-script/partner-commute.test.js
// シート・Googleマップ・ドライブをモックにして、パートナー交通費のGASを確かめる
const fs = require('fs'), path = require('path'), assert = require('assert');

// ── シートのモック ──
function chain(target) {
  return new Proxy(target, { get: (t, k) => (k in t ? t[k] : () => chain(t)) });
}
function Sheet(name) {
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
    _g: g, getName: () => name, getSheetId: () => 7,
    getRange: (a, c, nr, nc) => typeof a === 'string' ? range(1, 1, 1, 1) : range(a, c, nr, nc),
    getLastRow: () => { let n = g.length; while (n && !(g[n - 1] || []).some(x => x !== '' && x !== undefined)) n--; return n; },
    appendRow(row) { const r = this.getLastRow() + 1; row.forEach((x, j) => set(r, j + 1, x)); },
    deleteRow: r => g.splice(r - 1, 1),
    clear() { g.length = 0; return this; }
  });
}
const sheets = {};
const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = Sheet(n)) };
global.SpreadsheetApp = { openById: () => ss, flush() {}, BorderStyle: { SOLID: 1 } };
global.Utilities = {
  getUuid: (() => { let i = 0; return () => 'uuid-' + String(++i).padStart(6, '0') + '-xxxx'; })(),
  formatDate: d => d.toISOString().slice(0, 10), sleep() {},
  base64Decode: b => Buffer.from(b, 'base64'), newBlob: (bytes, type, name) => ({ name, type, bytes })
};

// ── Googleマップのモック ──
const PLACES = {
  '飯塚市枝国': '日本、〒820-0011 福岡県飯塚市枝国1-2-3',
  'イオン大牟田': '日本、〒836-0037 福岡県大牟田市岬町3-4',
  'au春日': '日本、〒816-0846 福岡県春日市星見ヶ丘3丁目'
};
const DIST = {  // 片道（m）・高速あり/なし
  '福岡県飯塚市枝国1-2-3|福岡県大牟田市岬町3-4': [84987, 92011],
  '福岡県大牟田市岬町3-4|福岡県飯塚市枝国1-2-3': [85129, 91000],
  '福岡県大牟田市岬町3-4|福岡県春日市星見ヶ丘3丁目': [50050, 52030],
  '福岡県春日市星見ヶ丘3丁目|福岡県飯塚市枝国1-2-3': [36107, 36107]
};
let mapCalls = 0;
global.Maps = {
  newGeocoder: () => ({ setLanguage() { return this; }, setRegion() { return this; },
    geocode: q => { const hit = Object.keys(PLACES).filter(k => q.includes(k) || PLACES[k].includes(q) || q.includes(PLACES[k].replace(/^日本、〒\d{3}-\d{4} /, '')));
      return { status: hit.length ? 'OK' : 'ZERO_RESULTS', results: hit.map(k => ({ formatted_address: PLACES[k] })) }; } }),
  newDirectionFinder: () => { const f = { o: '', d: '', avoid: false,
    setOrigin(x) { this.o = x; return this; }, setDestination(x) { this.d = x; return this; },
    setLanguage() { return this; }, setRegion() { return this; }, setMode() { return this; }, setAvoid() { this.avoid = true; return this; },
    getDirections() { mapCalls++; const v = DIST[this.o + '|' + this.d]; if (!v) return { status: 'NOT_FOUND', routes: [] };
      return { status: 'OK', routes: [{ legs: [{ distance: { value: v[this.avoid ? 1 : 0] } }] }] }; } }; return f; },
  DirectionFinder: { Mode: { DRIVING: 'd' }, Avoid: { HIGHWAYS: 'h' } }
};

// ── ドライブ・PDF・Webアプリのモック ──
let fileSeq = 0; const allFiles = [];
function Folder(name) {
  const f = { name, folders: [], files: [] };
  const it = a => { let i = 0; return { hasNext: () => i < a.length, next: () => a[i++] }; };
  f.getFoldersByName = n => it(f.folders.filter(x => x.name === n));
  f.createFolder = n => { const c = Folder(n); f.folders.push(c); return c; };
  f.getFilesByName = n => it(f.files.filter(x => x.name === n && !x.trashed));
  f.getFiles = () => it(f.files.filter(x => !x.trashed));
  f.createFile = b => { const file = { id: 'file' + (++fileSeq), name: b.name || b._name, trashed: false, getId() { return this.id; }, getName() { return this.name; }, setTrashed(v) { this.trashed = v; } }; f.files.push(file); allFiles.push(file); return file; };
  return f;
}
const partnerRoot = Folder('パートナー');
global.DriveApp = { getFolderById: id => { if (id !== '1dsIbQIqUfDrw9p4tKHGweY_hPRIM5h5K') throw new Error('nf'); return partnerRoot; },
  getFileById: id => { const f = allFiles.find(x => x.id === id); if (!f) throw new Error('nf'); return f; } };
global.ScriptApp = { getOAuthToken: () => 't' };
let pdfType = 'application/pdf';
global.UrlFetchApp = { fetch: () => ({ getResponseCode: () => 200, getBlob: () => ({ getContentType: () => pdfType, setName(n) { this.name = n; return this; } }) }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
global.ContentService = { MimeType: { JSON: 'json' }, createTextOutput: s => ({ setMimeType() { return JSON.parse(s); } }) };

(0, eval)(fs.readFileSync(path.join(__dirname, 'partner-commute.gs'), 'utf8'));
const api = (action, o) => doPost({ postData: { contents: JSON.stringify(Object.assign({ action, name: '山口' }, o)) } });
const ok = (action, o) => { const r = api(action, o); assert.ok(r.ok, action + ': ' + r.error); return r.data; };
const ng = (action, o, re) => { const r = api(action, o); assert.ok(!r.ok && re.test(r.error), action + ' は ' + re + ' で断る（' + r.error + '）'); };

// 1) 名前：パートナーだけ
assert.strictEqual(doPost({ postData: { contents: JSON.stringify({ action: 'init', name: '福山' }) } }).ok, false);
let me = ok('init');
assert.strictEqual(me.homeSet, false);
assert.strictEqual(me.fullName, '山口真人');
assert.strictEqual(me.from, '2026-12');

// 2) 住所を探す → 自宅を登録（画面には市区町村までしか返さない）
assert.deepStrictEqual(ok('searchAddress', { query: '飯塚市枝国' }), ['福岡県飯塚市枝国1-2-3']);
assert.deepStrictEqual(ok('searchAddress', { query: 'どこにもない' }), []);
ng('saveDays', { dates: ['2026-12-01'], stops: ['home', 'x', 'home'] }, /自宅を登録/);
me = ok('saveProfile', { homeAddress: '福岡県飯塚市枝国1-2-3' });
assert.strictEqual(me.homeSet, true);
assert.strictEqual(me.homeMasked, '福岡県飯塚市…');
assert.ok(!JSON.stringify(me).includes('枝国'), '番地は返さない');
ng('saveProfile', { homeAddress: 'どこにもない' }, /見つかりませんでした/);

// 3) 場所を登録・名前の重複・「自宅」は使えない・直す・消す
let places = ok('savePlace', { placeName: 'イオン大牟田', address: '日本、〒836-0037 福岡県大牟田市岬町3-4' });
assert.deepStrictEqual(places.map(p => p.name + '|' + p.address), ['イオン大牟田|福岡県大牟田市岬町3-4']);
places = ok('savePlace', { placeName: 'au春日', address: '福岡県春日市星見ヶ丘3丁目' });
ng('savePlace', { placeName: 'イオン大牟田', address: '福岡県春日市星見ヶ丘3丁目' }, /もう登録/);
ng('savePlace', { placeName: '自宅', address: '福岡県春日市星見ヶ丘3丁目' }, /自宅/);
const aeon = places[0].id, kasuga = places[1].id;
places = ok('savePlace', { id: kasuga, placeName: 'au 春日店', address: '福岡県春日市星見ヶ丘3丁目' });
assert.strictEqual(places[1].name, 'au 春日店');
// ほかの人の場所は見えない・使えない
const other = doPost({ postData: { contents: JSON.stringify({ action: 'init', name: '鈴木' }) } }).data;
assert.strictEqual(other.places.length, 0);

// 4) 経路のプレビュー：区間ごとの距離（0.1km）と合計・ガソリン代
let pv = ok('previewRoute', { stops: ['home', aeon, 'home'], highway: true });
assert.deepStrictEqual(pv.legs.map(l => l.from + '→' + l.to + ':' + l.km), ['自宅→イオン大牟田:85', 'イオン大牟田→自宅:85.1']);
assert.strictEqual(pv.km, 170.1);
assert.strictEqual(pv.gas, 2552);   // 170.1 × 15 = 2551.5 → 2552
const callsAfter = mapCalls;
ok('previewRoute', { stops: ['home', aeon, 'home'], highway: true });
assert.strictEqual(mapCalls, callsAfter, '同じ区間は距離メモから（Googleマップに聞かない）');
pv = ok('previewRoute', { stops: ['home', aeon, 'home'], highway: false });
assert.strictEqual(pv.km, 183, '高速を使わないときは別の道で測る');
// 経由地：自宅 → イオン → au春日 → 自宅
pv = ok('previewRoute', { stops: ['home', aeon, kasuga, 'home'], highway: true });
assert.strictEqual(pv.text, '自宅 → イオン大牟田 → au 春日店 → 自宅');
assert.strictEqual(pv.km, 171.2);
ng('previewRoute', { stops: ['home', aeon, aeon, 'home'] }, /同じ場所が続いて/);
ng('previewRoute', { stops: ['home'] }, /2か所以上/);
ng('previewRoute', { stops: ['home', 'nai', 'home'] }, /登録されていない/);
ng('previewRoute', { stops: ['home', kasuga, aeon, 'home'], highway: true }, /距離を出せませんでした/);

// 5) 保存：いくつかの日をまとめて・高速代・駐車場代・上書き
let s = ok('saveDays', { dates: ['2026-12-01', '2026-12-02'], stops: ['home', aeon, 'home'], highway: true, toll: 2700, parking: 500, note: '' });
assert.strictEqual(s.saved, 2);
ok('saveDays', { dates: ['2026-12-05'], stops: ['home', aeon, kasuga, 'home'], highway: true, toll: 0, parking: 300, note: '春日で応援' });
ok('saveDays', { dates: ['2026-12-02'], stops: ['home', aeon, 'home'], highway: false });   // 2日を直す
let m = ok('listMonth', { ym: '2026-12' });
// 1日：170.1km → 2552 + 高速2700 + 駐車500 ／ 2日：183km（高速なしに直した）→ 2745 ／ 5日：171.2km → 2568 + 駐車300
assert.deepStrictEqual(m.entries.map(e => e.date + ' ' + e.km + ' ' + e.total), ['2026-12-01 170.1 5752', '2026-12-02 183 2745', '2026-12-05 171.2 2868']);
assert.deepStrictEqual(m.entries[2].stops, ['自宅', 'イオン大牟田', 'au 春日店', '自宅']);
assert.strictEqual(m.entries[2].note, '春日で応援');
assert.strictEqual(m.entries[1].highway, false);
assert.strictEqual(m.entries[1].toll, 0, '直した日は前の高速代を持ち越さない');
assert.deepStrictEqual(m.totals, { days: 3, routes: 1, km: 524.3, gas: 7865, toll: 2700, parking: 800, total: 11365 });
assert.strictEqual(sheets['記録']._g.length, 4, '同じ日は上書き（行が増えない）');
assert.ok(!JSON.stringify(m).includes('枝国'), '一覧にも自宅の住所は出さない');
// 場所の名前を変えたり消したりしても、前の記録の経路は変わらない
ok('deletePlace', { id: kasuga });
assert.strictEqual(ok('init').places.length, 1);
assert.deepStrictEqual(ok('listMonth', { ym: '2026-12' }).entries[2].stops, ['自宅', 'イオン大牟田', 'au 春日店', '自宅']);
ng('previewRoute', { stops: ['home', kasuga, 'home'] }, /登録されていない/);

// 6) 入力の確かめ
ng('saveDays', { dates: ['2026-11-30'], stops: ['home', aeon, 'home'] }, /12月分からこの画面.*11月分は前の画面/);
ng('listMonth', { ym: '2026-11' }, /前の画面/);
ng('saveDays', { dates: ['2026-12-31', '2027-01-01'], stops: ['home', aeon, 'home'] }, /違う月/);
ng('saveDays', { dates: ['2026-12-32'], stops: ['home', aeon, 'home'] }, /日付が正しくありません/);
ng('saveDays', { dates: [], stops: ['home', aeon, 'home'] }, /日付を選んで/);
ng('saveDays', { dates: ['2026-12-03'], stops: ['home', aeon, 'home'], highway: false, toll: 1000 }, /高速を使った/);
ng('saveDays', { dates: ['2026-12-03'], stops: ['home', aeon, 'home'], parking: -1 }, /0以上/);
ng('saveDays', { dates: ['2026-12-03'], stops: ['home', aeon, 'home'], parking: 1.5 }, /整数/);
ng('saveDays', { dates: ['2026-12-03'], stops: ['home', aeon, 'home'], parking: 100001 }, /大きすぎ/);
ng('nanika', {}, /不明なaction/);
// 鈴木さんは月の制限なし
assert.ok(doPost({ postData: { contents: JSON.stringify({ action: 'listMonth', name: '鈴木', ym: '2026-11' }) } }).ok);

// 7) PDF：パートナー/2026年/12月/2026-12_山口_稼働費.pdf、表は日付順＋合計行
let pdf = ok('exportPdf', { ym: '2026-12' });
assert.strictEqual(pdf.name, '2026-12_山口_稼働費.pdf');
ok('exportPdf', { ym: '2026-12' });
const dec = partnerRoot.folders[0].folders[0];
assert.strictEqual(partnerRoot.folders[0].name + '/' + dec.name, '2026年/12月');
assert.deepStrictEqual(dec.files.filter(f => !f.trashed).map(f => f.name), ['2026-12_山口_稼働費.pdf']);
const g = sheets['PDF作成用']._g;
assert.strictEqual(g[1][4], '山口真人');
assert.strictEqual(g[1][7], '3日');
assert.deepStrictEqual(g[5].slice(0, 8), ['12/1', '火', '自宅 → イオン大牟田 → 自宅（高速）', 170.1, 2552, 2700, 500, 5752]);
assert.deepStrictEqual(g[8].slice(0, 8), ['合計', '', '', 524.3, 7865, 2700, 800, 11365]);
ng('exportPdf', { ym: '2027-01' }, /入力がまだありません/);
pdfType = 'text/html';
ng('exportPdf', { ym: '2026-12' }, /HTML/);
pdfType = 'application/pdf';
// 氏名を変えるとPDFにも
ok('saveProfile', { fullName: '山口 真人' });
ok('exportPdf', { ym: '2026-12' });
assert.strictEqual(sheets['PDF作成用']._g[1][4], '山口 真人');
assert.strictEqual(ok('init').homeSet, true, '氏名だけ直しても自宅は残る');

// 8) 添付：その日の領収書・月のETC明細を足していく（前の分は消さない）
const b64 = Buffer.from('x').toString('base64');
const add = (o) => ok('addAttachment', Object.assign({ mimeType: 'image/jpeg', data: b64 }, o));
assert.deepStrictEqual(add({ kind: 'parking', date: '2026-12-01', uploadId: 'up-0000001' }), { id: 'up-0000001', file: '2026-12-01_山口_駐車場.jpg' });
add({ kind: 'parking', date: '2026-12-01', uploadId: 'up-0000002', mimeType: 'application/pdf' });
add({ kind: 'toll', date: '2026-12-01', uploadId: 'up-0000003' });
add({ kind: 'parking', date: '2026-12-05', uploadId: 'up-0000004' });
add({ kind: 'ETC', ym: '2026-12', uploadId: 'up-0000005', mimeType: 'application/pdf' });
add({ kind: 'ETC', ym: '2026-12', uploadId: 'up-0000006', mimeType: 'application/pdf' });
// 通信のやり直し（同じ uploadId）では増えない
assert.strictEqual(add({ kind: 'ETC', ym: '2026-12', uploadId: 'up-0000006', mimeType: 'application/pdf' }).again, true);
const live = () => dec.files.filter(f => !f.trashed).map(f => f.name).sort();
assert.deepStrictEqual(live(), ['2026-12-01_山口_駐車場.jpg', '2026-12-01_山口_駐車場_2.pdf', '2026-12-01_山口_高速.jpg', '2026-12-05_山口_駐車場.jpg',
  '2026-12_山口_ETC.pdf', '2026-12_山口_ETC_2.pdf', '2026-12_山口_稼働費.pdf']);
let files = ok('listMonth', { ym: '2026-12' }).files;
assert.deepStrictEqual(files.map(f => (f.date || '月') + ' ' + f.kind + ' ' + f.file), [
  '2026-12-01 parking 2026-12-01_山口_駐車場.jpg', '2026-12-01 parking 2026-12-01_山口_駐車場_2.pdf', '2026-12-01 toll 2026-12-01_山口_高速.jpg',
  '2026-12-05 parking 2026-12-05_山口_駐車場.jpg', '月 ETC 2026-12_山口_ETC.pdf', '月 ETC 2026-12_山口_ETC_2.pdf']);
// 消す：ドライブのファイルもゴミ箱へ。消したあとに足すと、空いた名前を使う
ok('deleteAttachment', { id: 'up-0000001' });
assert.ok(!live().includes('2026-12-01_山口_駐車場.jpg'));
assert.strictEqual(ok('listMonth', { ym: '2026-12' }).files.length, 5);
assert.strictEqual(add({ kind: 'parking', date: '2026-12-01', uploadId: 'up-0000007' }).file, '2026-12-01_山口_駐車場.jpg');
ng('deleteAttachment', { id: 'up-0000001' }, /見つかりません/);
// PDFに添付の数が載る
ok('exportPdf', { ym: '2026-12' });
assert.ok(sheets['PDF作成用']._g.some(r => String(r[0]).includes('ETCの利用明細 2件・領収書 4枚')));
// 確かめ
ng('addAttachment', { kind: 'x', date: '2026-12-01', uploadId: 'up-0000008', mimeType: 'image/png', data: b64 }, /種類/);
ng('addAttachment', { kind: 'parking', date: '2026-12-01', uploadId: 'up-0000008', mimeType: 'text/html', data: b64 }, /PDFか写真/);
ng('addAttachment', { kind: 'parking', date: '2026-11-30', uploadId: 'up-0000008', mimeType: 'image/png', data: b64 }, /前の画面/);
ng('addAttachment', { kind: 'parking', date: '2026-12-40', uploadId: 'up-0000008', mimeType: 'image/png', data: b64 }, /日付/);
ng('addAttachment', { kind: 'parking', date: '2026-12-01', uploadId: 'x', mimeType: 'image/png', data: b64 }, /送信ID/);
ng('addAttachment', { kind: 'parking', date: '2026-12-01', uploadId: 'up-0000008', mimeType: 'image/png', data: '' }, /空/);
ng('addAttachment', { kind: 'ETC', ym: '2026-1', uploadId: 'up-0000008', mimeType: 'image/png', data: b64 }, /月の指定/);
// ほかの人の添付は見えない・消せない
assert.strictEqual(doPost({ postData: { contents: JSON.stringify({ action: 'listMonth', name: '鈴木', ym: '2026-12' }) } }).data.files.length, 0);
assert.strictEqual(doPost({ postData: { contents: JSON.stringify({ action: 'deleteAttachment', name: '鈴木', id: 'up-0000002' }) } }).ok, false);

// 9) 消す・みんなの状況用の集計・/exec
ok('deleteEntry', { id: m.entries[0].id });
assert.deepStrictEqual(ok('summary', { ym: '2026-12' }), { ym: '2026-12', days: 2, routes: 1, amount: 2745 + 2868 });
ng('deleteEntry', { id: 'nai' }, /見つかりません/);
assert.strictEqual(doGet().data.version, CODE_VERSION);
assert.deepStrictEqual(doGet().data.partners, ['山口', '鈴木', '毛利', '山内']);
console.log('partner-commute: all tests passed');
