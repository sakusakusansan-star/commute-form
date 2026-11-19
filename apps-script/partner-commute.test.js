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
  '福岡県春日市星見ヶ丘3丁目|福岡県飯塚市枝国1-2-3': [36107, 36107],
  '福岡県春日市星見ヶ丘3丁目|福岡県大牟田市岬町3-4': [50100, 52000]
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
const partnerRoot = Folder('パートナー'), staffRoot = Folder('社員');
global.DriveApp = { getFolderById: id => { if (id === '1VreXxTQpiTeiAoc_uxsLrpj1KloJ7V2o') return staffRoot; if (id !== '1dsIbQIqUfDrw9p4tKHGweY_hPRIM5h5K') throw new Error('nf'); return partnerRoot; },
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

// 1) 名前：名簿が無いときは PARTNERS と STAFF（社員も12月分から使える）。ほかの人は断る
const fk = doPost({ postData: { contents: JSON.stringify({ action: 'init', name: '福山' }) } });
assert.ok(fk.ok, fk.error);
assert.deepStrictEqual([fk.data.group, fk.data.from, fk.data.fullName], ['staff', '2026-12', '福山康平']);
assert.strictEqual(doPost({ postData: { contents: JSON.stringify({ action: 'init', name: '佐藤' }) } }).ok, false);
assert.deepStrictEqual(doGet().data.partners, ['山口', '鈴木', '毛利', '山内']);
assert.deepStrictEqual(doGet().data.users, ['山口', '鈴木', '毛利', '山内', '福山', '古川', '清田']);
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

// 7) PDF：パートナー/山口/2026年/12月/2026-12_山口_稼働費.pdf、表は日付順＋合計行
let pdf = ok('exportPdf', { ym: '2026-12' });
assert.strictEqual(pdf.name, '2026-12_山口_稼働費.pdf');
ok('exportPdf', { ym: '2026-12' });
const dec = partnerRoot.folders[0].folders[0].folders[0];
assert.strictEqual(partnerRoot.folders[0].name + '/' + partnerRoot.folders[0].folders[0].name + '/' + dec.name, '山口/2026年/12月');
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

// 10) 自宅を2つ（3つまで）登録して、入力のたびに選ぶ
me = ok('init');
assert.deepStrictEqual(me.homes.map(h => h.label + '|' + h.main), ['自宅|true']);
ng('saveHome', { label: '自宅', address: '福岡県春日市星見ヶ丘3丁目' }, /もう使っています/);
ng('saveHome', { label: 'イオン大牟田', address: '福岡県春日市星見ヶ丘3丁目' }, /稼働場所の名前/);
ng('saveHome', { label: '', address: '福岡県春日市星見ヶ丘3丁目' }, /呼び名を入れて/);
ng('saveHome', { label: '実家', address: 'どこにもない' }, /見つかりませんでした/);
me = ok('saveHome', { label: '実家', address: '日本、〒816-0846 福岡県春日市星見ヶ丘3丁目' });
assert.deepStrictEqual(me.homes.map(h => h.label + '|' + h.masked + '|' + h.main), ['自宅|福岡県飯塚市…|true', '実家|福岡県春日市…|false']);
assert.ok(!JSON.stringify(me).includes('星見ヶ丘'), '番地は返さない');
const jikka = me.homes[1].id;
ng('savePlace', { placeName: '実家', address: '福岡県春日市星見ヶ丘3丁目' }, /自宅の呼び名/);
// 実家から出て実家に帰る／いつもの自宅から出て実家に帰る
pv = ok('previewRoute', { stops: ['home:' + jikka, aeon, 'home:' + jikka], highway: true });
assert.strictEqual(pv.text, '実家 → イオン大牟田 → 実家');
assert.strictEqual(pv.km, 50.1 + 50.1);
pv = ok('previewRoute', { stops: ['home', aeon, 'home:' + jikka], highway: true });
assert.strictEqual(pv.text, '自宅 → イオン大牟田 → 実家');
ng('previewRoute', { stops: ['home:nai', aeon, 'home'] }, /登録されていない自宅/);
s = ok('saveDays', { dates: ['2026-12-10'], stops: ['home:' + jikka, aeon, 'home:' + jikka], highway: true });
assert.deepStrictEqual(ok('listMonth', { ym: '2026-12' }).entries.filter(e => e.date === '2026-12-10')[0].stops, ['実家', 'イオン大牟田', '実家']);
// 「いつも」を実家にする → 'home' は実家
me = ok('saveHome', { id: jikka, main: true });
assert.deepStrictEqual(me.homes.map(h => h.label + '|' + h.main), ['自宅|false', '実家|true']);
assert.strictEqual(me.homeMasked, '福岡県春日市…');
assert.strictEqual(ok('previewRoute', { stops: ['home', aeon, 'home'], highway: true }).text, '実家 → イオン大牟田 → 実家');
// 呼び名を変えても前の記録はそのまま
me = ok('saveHome', { id: jikka, label: 'アパート' });
assert.strictEqual(me.homes[1].label, 'アパート');
assert.strictEqual(me.homes[1].main, true, '直しても「いつも」のまま');
assert.deepStrictEqual(ok('listMonth', { ym: '2026-12' }).entries.filter(e => e.date === '2026-12-10')[0].stops, ['実家', 'イオン大牟田', '実家']);
// 3つまで
ok('saveHome', { label: '寮', address: '福岡県大牟田市岬町3-4' });
ng('saveHome', { label: '別荘', address: '福岡県大牟田市岬町3-4' }, /3か所まで/);
// 消す：いつもの自宅を消すと、残りの最初が「いつも」に。最後の1つは消せない
me = ok('deleteHome', { id: jikka });
assert.deepStrictEqual(me.homes.map(h => h.label + '|' + h.main), ['自宅|true', '寮|false']);
me = ok('deleteHome', { id: me.homes[1].id });
ng('deleteHome', { id: me.homes[0].id }, /1つ残して/);
assert.strictEqual(sheets['人']._g.find(r => r[0] === '山口')[2], '福岡県飯塚市枝国1-2-3', '人シートの自宅住所はいつもの自宅');
// 前のやり方（人シートだけ）で登録した人も、そのまま使えて、2つ目を足せる
sheets['人']._g.push(['毛利', '', '福岡県飯塚市枝国1-2-3', new Date()]);
const mouri = o => doPost({ postData: { contents: JSON.stringify(Object.assign({ name: '毛利' }, o)) } });
assert.deepStrictEqual(mouri({ action: 'init' }).data.homes.map(h => h.label), ['自宅']);
assert.deepStrictEqual(mouri({ action: 'saveHome', label: '実家', address: '福岡県春日市星見ヶ丘3丁目' }).data.homes.map(h => h.label + '|' + h.main), ['自宅|true', '実家|false']);
// 名簿B：使える人は名簿のパートナーで「稼働費」に印のある人（外した人をのぞく）。from は PARTNERS から
const who = (name, o) => doPost({ postData: { contents: JSON.stringify(Object.assign({ name }, o)) } });
assert.strictEqual(who('田中', { action: 'init' }).ok, false, '名簿が無いときは PARTNERS');
const rs = ss.insertSheet('名簿');
rs._g.push(['名前', '区分', 'フルネーム', 'キャリア', '稼働費', '経費精算', 'みんなの状況', '外した日', '直した日時']);
rs._g.push(['山口', 'パートナー', '山口真人', 'au', '○', '', '', '', '']);
rs._g.push(['毛利', 'パートナー', '', '', '○', '', '', '2026-12-01', '']);
rs._g.push(['田中', 'パートナー', '田中一郎', '', '○', '', '', '', '']);
rs._g.push(['佐藤', 'パートナー', '', '', '', '', '', '', '']);
rs._g.push(['福山', '社員', '福山康平', '', '○', '○', '', '', '']);
let ng2, t = who('田中', { action: 'init' });
assert.ok(t.ok, t.error);
assert.deepStrictEqual([t.data.fullName, t.data.from], ['田中一郎', '']);
assert.strictEqual(who('山口', { action: 'init' }).data.from, '2026-12', 'from は PARTNERS のまま');
['毛利', '佐藤', '鈴木'].forEach(n => assert.strictEqual(who(n, { action: 'init' }).ok, false, n + ' は使えない'));
// 名簿の社員で「稼働費」に印のある人も使える（STAFF の人は12月分から）
t = who('福山', { action: 'init' });
assert.ok(t.ok, t.error);
assert.deepStrictEqual([t.data.group, t.data.from, t.data.fullName], ['staff', '2026-12', '福山康平']);
ng2 = who('福山', { action: 'saveDays', dates: ['2026-11-30'], stops: ['home', 'x', 'home'] });
assert.ok(!ng2.ok && /12月分から/.test(ng2.error), ng2.error);
// checkRoster はログに出すだけ
const logs = []; const oldLogger = global.Logger; global.Logger = { log: m => logs.push(String(m)) };
checkRoster();
global.Logger = oldLogger;
assert.ok(logs[0].startsWith('名簿 5人：') && logs[0].includes('毛利（外した）') && logs[0].endsWith('この画面を使える人：山口・田中・福山'), logs[0]);
// 前の画面（山口_稼働費計算_GAS版）から自宅と場所を写す。2回目は増えない
{
  const oldSh = Sheet('稼働場所');
  [['短縮名', '住所'], ['田中自宅', '福岡県飯塚市枝国1-2-3'], ['au春日', '福岡県春日市星見ヶ丘3丁目'], ['イオン大牟田', '〒836-0037 福岡県大牟田市岬町3-4'], ['', '']].forEach(r => oldSh.appendRow(r));
  const realOpen = SpreadsheetApp.openById;
  SpreadsheetApp.openById = id => id === 'OLD' ? { getSheetByName: n => n === '稼働場所' ? oldSh : null } : realOpen(id);
  const r1 = importOldPlaces_('田中', 'OLD');
  assert.strictEqual(r1, '田中：自宅 1件・場所 2件を写しました', r1);
  const me = doPost({ postData: { contents: JSON.stringify({ name: '田中', action: 'init' }) } });
  assert.ok(me.ok, me.error);
  assert.deepStrictEqual(me.data.homes.map(h => h.label + '|' + h.main), ['自宅|true']);
  assert.deepStrictEqual(me.data.places.map(p => p.name), ['au春日', 'イオン大牟田']);
  assert.strictEqual(sheets['場所']._g.find(r => r[2] === 'イオン大牟田')[3], '福岡県大牟田市岬町3-4', '〒は取る');
  assert.strictEqual(sheets['人']._g.find(r => r[0] === '田中')[2], '福岡県飯塚市枝国1-2-3');
  const r2 = importOldPlaces_('田中', 'OLD');
  assert.ok(r2.startsWith('田中：自宅 0件・場所 0件') && r2.includes('au春日'), r2);
  assert.strictEqual(sheets['場所']._g.filter(r => r[1] === '田中').length, 2);
  assert.ok(importOldPlaces_('田中', 'NONE').includes('見つかりません'));
  SpreadsheetApp.openById = realOpen;
}
// 社員：PDFは「社員」フォルダに。前のブック（全員の場所が1枚）からは、その人が使った場所と自宅だけ写す
{
  const fk = (o) => doPost({ postData: { contents: JSON.stringify(Object.assign({ name: '福山' }, o)) } });
  const book = {}, sh = (n, rows) => { const x = Sheet(n); rows.forEach(r => x.appendRow(r)); book[n] = x; };
  sh('稼働場所', [['短縮名', '住所'], ['福山自宅', '福岡県飯塚市枝国1-2-3'], ['古川自宅', '熊本県熊本市南区'], ['au春日', '福岡県春日市星見ヶ丘3丁目'], ['イオン大牟田', '福岡県大牟田市岬町3-4'], ['ドコモ近見', '熊本県熊本市南区近見']]);
  sh('福山康平', [['経費精算書'], ['氏名'], [''], [''], ['日付', '曜', '稼　働　場　所'], ['11/1', '日', 'au春日'], ['11/2', '月', '']]);
  sh('稼働費履歴', [['年月', 'スタッフ', '種類', '日', '場所・内容'], ['2026-10', '福山康平', '通常', 2, 'イオン大牟田'], ['2026-10', '古川泰成', '通常', 3, 'ドコモ近見'], ['2026-10', '福山康平', 'イレギュラー', 4, '出張']]);
  const realOpen = SpreadsheetApp.openById;
  SpreadsheetApp.openById = id => id === OLD_MAIN_BOOK ? { getSheetByName: n => book[n] || null } : realOpen(id);
  const r1 = importOldPlaces_('福山', OLD_BOOKS['福山']);
  assert.strictEqual(r1, '福山：自宅 1件・場所 2件を写しました', r1);
  const me = fk({ action: 'init' }).data;
  assert.deepStrictEqual(me.places.map(p => p.name), ['au春日', 'イオン大牟田'], 'ほかの人の場所（ドコモ近見）や自宅は写さない');
  assert.deepStrictEqual(me.homes.map(h => h.label + '|' + h.main), ['自宅|true']);
  assert.ok(importOldPlaces_('清田', { id: OLD_MAIN_BOOK, sheet: '清田和男' }).includes('見つかりません'), 'シートも履歴も無い人は止める');
  SpreadsheetApp.openById = realOpen;
  // 12月分を入れて PDF → 社員/福山/2026年/12月
  const au = me.places[1].id;
  let r = fk({ action: 'saveDays', dates: ['2026-12-01'], stops: ['home', au, 'home'], highway: false });
  assert.ok(r.ok, r.error);
  r = fk({ action: 'exportPdf', ym: '2026-12' });
  assert.ok(r.ok, r.error);
  const f = staffRoot.folders[0];
  assert.strictEqual(f.name + '/' + f.folders[0].name + '/' + f.folders[0].folders[0].name, '福山/2026年/12月');
  assert.strictEqual(f.folders[0].folders[0].files[0].name, '2026-12_福山_稼働費.pdf');
}
// 場所の一覧（日報のブックの「場所」シート）：本人が登録していない場所でも 'master:名前' で入れられる。使ったら本人の場所にも足す
{
  const master = Sheet('場所(一覧)');
  [['名前', '別名', '住所', '店舗稼働', '担当店舗', 'カッコなし', '外した日'], ['ゆめタウン大牟田', '', '福岡県大牟田市岬町3-4', '', '', '○', ''], ['外した場所', '', '福岡県大牟田市岬町3-4', '', '', '', '2026-11-01']].forEach(r => master.appendRow(r));
  const realOpen = SpreadsheetApp.openById;
  SpreadsheetApp.openById = id => id === ROSTER_BOOK_ID ? { getSheetByName: n => n === '場所' ? master : ss.getSheetByName(n) } : ss;
  const yg = (o) => doPost({ postData: { contents: JSON.stringify(Object.assign({ name: '山口' }, o)) } });
  const before = yg({ action: 'init' }).data.places.map(p => p.name);
  assert.ok(!before.includes('ゆめタウン大牟田'));
  let r = yg({ action: 'saveDays', dates: ['2026-12-20'], stops: ['home', 'master:ゆめタウン大牟田', 'home'], highway: false });
  assert.ok(r.ok, r.error);
  assert.strictEqual(r.data.route.text, '自宅 → ゆめタウン大牟田 → 自宅');
  const after = yg({ action: 'init' }).data.places.map(p => p.name);
  assert.ok(after.includes('ゆめタウン大牟田'), '使った場所は本人の場所に足す');
  // 2回目は足さない
  yg({ action: 'saveDays', dates: ['2026-12-21'], stops: ['home', 'master:ゆめタウン大牟田', 'home'], highway: false });
  assert.strictEqual(yg({ action: 'init' }).data.places.filter(p => p.name === 'ゆめタウン大牟田').length, 1);
  // 一覧に無い・外した場所は断る
  r = yg({ action: 'saveDays', dates: ['2026-12-22'], stops: ['home', 'master:外した場所', 'home'], highway: false });
  assert.ok(!r.ok && /一覧に無い/.test(r.error), r.error);
  SpreadsheetApp.openById = realOpen;
}
console.log('partner-commute: all tests passed');
