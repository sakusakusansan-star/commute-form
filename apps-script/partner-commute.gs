// ============================================================
//  パートナー交通費（稼働費）— Google Apps Script（Webアプリ）
//
//  パートナー（山口・鈴木・毛利・山内）が、稼働費入力の画面
//  （commute-form/partner.html）から使う。社員（福山・古川・清田）は
//  これまでどおり commute-expense-sheet.gs のほう。
//
//  ・自宅と稼働場所は、本人が画面から住所で登録する（Googleマップで場所を確かめる）
//  ・1日の経路は「自宅 → 場所 → 自宅」。イレギュラーの日は経由地を足して
//    「自宅 → A → B → 自宅」にできる。距離は区間ごとにGoogleマップで出して足す
//  ・ガソリン代は 1km 15円。高速代・駐車場代は実費を入れる
//  ・記録は「記録」シートに日ごとに1行ずつ残す（月を切り替えても消えない）
//  ・PDFは Googleドライブ「経費精算/パートナー/2026年/12月/2026-12_鈴木_稼働費.pdf」
//    ETCの明細・駐車場の領収書も同じフォルダに添付として保存する
//  ・同じ区間の距離は「距離メモ」シートに覚えておき、毎回同じ距離になるようにする
//
//  【デプロイ】このスプレッドシート（パートナー_交通費）の Apps Script に貼り、
//   ウェブアプリとしてデプロイ（実行するユーザー：自分／アクセス：全員）。
//   シートは最初に呼ばれたときに自動で作る。関数を手で実行する必要はない。
//
//  【パートナーが増えたら】PARTNERS に名字を足す（名簿 sieg-roster.js と同じ名前）。
// ============================================================

var SHEET_ID = '19TTixFCWjVUR9S4XvsMjdbQt_U3v4MlCdvVpgTMsK_U';   // パートナー_交通費（12月〜）
var CODE_VERSION = '2026-11-04';
var RATE_PER_KM = 15;

// 名前 → { full: PDFに載せる氏名（本人が画面で登録すればそちらを使う）, from: この画面で入力できる最初の月 }
var PARTNERS = {
  '山口': { full: '山口真人', from: '2026-12' },   // 11月分までは前の画面（山口_稼働費計算_GAS版）
  '鈴木': {},
  '毛利': {},
  '山内': {}
};

var PDF_FOLDER_ID = '1dsIbQIqUfDrw9p4tKHGweY_hPRIM5h5K';   // 経費精算/パートナー
var MAX_STOPS = 10;          // 1日の経路に入れられる地点の数（自宅を含む）
var MAX_AMOUNT = 100000;     // 高速代・駐車場代の1日の上限（打ち間違い防止）

var SH = { people: '人', places: '場所', records: '記録', memo: '距離メモ', pdf: 'PDF作成用' };
var HEAD = {
  '人':      ['名前', '氏名', '自宅住所', '更新日時'],
  '場所':    ['ID', '名前', '場所名', '住所', '作成日時', '削除'],
  '記録':    ['ID', '名前', '日付', '経路', '経路の表示', '区間km', '距離km', '高速', 'ガソリン代', '高速代', '駐車場代', '合計', '備考', '更新日時'],
  '距離メモ': ['キー', '出発', '到着', '高速', 'km', '日時']
};
var R = { ID: 0, NAME: 1, DATE: 2, STOPS: 3, TEXT: 4, LEGS: 5, KM: 6, HW: 7, GAS: 8, TOLL: 9, PARK: 10, TOTAL: 11, NOTE: 12, AT: 13 };
var WEEK = ['日', '月', '火', '水', '木', '金', '土'];

// ── シート ──────────────────────────────────────────────
function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function sheet_(name) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (HEAD[name]) {
      sh.getRange(1, 1, 1, HEAD[name].length).setValues([HEAD[name]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    // 日付は文字のまま持つ（「2026-12-01」を日付に変えられると読み書きがずれる）
    if (name === SH.records) sh.getRange('C:C').setNumberFormat('@');
  }
  return sh;
}

function rows_(name) {
  var sh = sheet_(name);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEAD[name].length).getValues();
}

function now_() { return new Date(); }
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

function dateText_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  return String(v || '');
}

// ── 人・月の確かめ ──────────────────────────────────────
function checkName_(name) {
  name = String(name || '');
  if (!Object.prototype.hasOwnProperty.call(PARTNERS, name)) throw new Error('この画面を使えない名前です: ' + name);
  return name;
}

function checkMonth_(name, ym) {
  ym = String(ym || '');
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(ym)) throw new Error('月の指定が正しくありません: ' + ym);
  var from = PARTNERS[name].from;
  if (from && ym < from) {
    throw new Error(Number(from.slice(5)) + '月分からこの画面で入力します。' + Number(ym.slice(5)) + '月分は前の画面で入力してください。');
  }
  return ym;
}

function checkDate_(d) {
  d = String(d || '');
  var m = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('日付が正しくありません: ' + d);
  var dt = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  if (dt.getMonth() !== Number(m[2]) - 1 || dt.getDate() !== Number(m[3])) throw new Error('日付が正しくありません: ' + d);
  return d;
}

function yen_(v, label) {
  if (v === '' || v === null || v === undefined) return 0;
  var n = Number(v);
  if (!isFinite(n) || n < 0 || Math.floor(n) !== n) throw new Error(label + 'は0以上の整数（円）で入れてください。');
  if (n > MAX_AMOUNT) throw new Error(label + 'が大きすぎます（1日 ' + MAX_AMOUNT.toLocaleString() + '円まで）。');
  return n;
}

// ── 住所 ────────────────────────────────────────────────
// Googleマップの住所「日本、〒820-0011 福岡県飯塚市…」から「日本、」と郵便番号を外す
function cleanAddress_(a) {
  return String(a || '').replace(/^日本、\s*/, '').replace(/〒?\d{3}-\d{4}\s*/, '').trim();
}

// 画面に出す自宅住所は市区町村まで（フォームは誰でも開けるので、番地までは返さない）
function maskAddress_(a) {
  a = cleanAddress_(a);
  if (!a) return '';
  var m = a.match(/^(.*?[都道府県])(.*?[市区町村])/) || a.match(/^(.*?[市区町村])/);
  return (m ? m[0] : a.slice(0, 6)) + '…';
}

function geocode_(query) {
  query = String(query || '').trim();
  if (!query) throw new Error('住所か場所の名前を入れてください。');
  if (query.length > 120) throw new Error('長すぎます（120文字まで）。');
  var res = Maps.newGeocoder().setLanguage('ja').setRegion('jp').geocode(query);
  if (res && res.status === 'OVER_QUERY_LIMIT') throw new Error('Googleマップの今日の回数を超えました。明日もう一度お試しください。');
  var seen = {};
  return ((res && res.results) || []).map(function (r) {
    return cleanAddress_(r.formatted_address);
  }).filter(function (a) {
    if (!a || seen[a]) return false;
    seen[a] = true;
    return true;
  }).slice(0, 5);
}

// ── 距離 ────────────────────────────────────────────────
// 同じ区間は「距離メモ」の値を使う（Googleマップの結果は日によって少し変わるため、
// 毎回同じ距離にそろえる。回数の節約にもなる）
function memo_() {
  var map = {};
  rows_(SH.memo).forEach(function (r) { map[r[0]] = Number(r[4]); });
  return map;
}

function legKm_(from, to, highway, memo) {
  var key = from + '→' + to + (highway ? '｜高速' : '｜一般');
  if (memo[key] > 0) return memo[key];
  var lastErr = null;
  for (var i = 0; i < 3; i++) {
    try {
      var finder = Maps.newDirectionFinder().setOrigin(from).setDestination(to)
        .setLanguage('ja').setRegion('jp').setMode(Maps.DirectionFinder.Mode.DRIVING);
      if (!highway) finder.setAvoid(Maps.DirectionFinder.Avoid.HIGHWAYS);
      var d = finder.getDirections();
      if (d && d.status === 'OVER_QUERY_LIMIT') throw new Error('Googleマップの今日の回数を超えました。明日もう一度お試しください。');
      if (!d || !d.routes || !d.routes.length) {
        lastErr = new Error('道が見つかりませんでした');
        break;
      }
      var meters = d.routes[0].legs.reduce(function (s, l) { return s + l.distance.value; }, 0);
      var km = Math.round(meters / 100) / 10;
      memo[key] = km;
      sheet_(SH.memo).appendRow([key, from, to, highway ? '有' : '無', km, now_()]);
      return km;
    } catch (err) {
      lastErr = err;
      if (/回数を超えました/.test(err.message)) throw err;
      Utilities.sleep(400 * (i + 1));
    }
  }
  throw new Error('距離を出せませんでした（' + (lastErr ? lastErr.message : '') + '）。場所の住所を確かめてください。');
}

// stops：'home' か場所のID の並び → 名前と住所
function resolveStops_(name, stops) {
  if (!stops || !stops.length || stops.length < 2) throw new Error('経路は2か所以上にしてください。');
  if (stops.length > MAX_STOPS) throw new Error('経路に入れられるのは ' + MAX_STOPS + ' か所までです。');
  var person = person_(name);
  var places = placesOf_(name);
  var out = stops.map(function (s) {
    if (s === 'home') {
      if (!person.home) throw new Error('先に「設定」で自宅を登録してください。');
      return { id: 'home', name: '自宅', address: person.home };
    }
    for (var i = 0; i < places.length; i++) if (places[i].id === s) return places[i];
    throw new Error('登録されていない場所が入っています。場所を選び直してください。');
  });
  for (var i = 1; i < out.length; i++) {
    if (out[i].id === out[i - 1].id) throw new Error('同じ場所が続いています（' + out[i].name + '）。');
  }
  return out;
}

function route_(name, stops, highway) {
  var list = resolveStops_(name, stops);
  var memo = memo_();
  var legs = [], total = 0;
  for (var i = 1; i < list.length; i++) {
    var km = legKm_(list[i - 1].address, list[i].address, highway, memo);
    legs.push({ from: list[i - 1].name, to: list[i].name, km: km });
    total += km;
  }
  total = Math.round(total * 10) / 10;
  return { stops: list, legs: legs, km: total, gas: Math.round(total * RATE_PER_KM) };
}

function routeText_(stops) { return stops.map(function (s) { return s.name; }).join(' → '); }

// ── 人・場所 ────────────────────────────────────────────
function person_(name) {
  var rs = rows_(SH.people);
  for (var i = 0; i < rs.length; i++) {
    if (rs[i][0] === name) return { row: i + 2, fullName: String(rs[i][1] || ''), home: String(rs[i][2] || '') };
  }
  return { row: 0, fullName: '', home: '' };
}

function placesOf_(name) {
  return rows_(SH.places).filter(function (r) { return r[1] === name && !r[5]; })
    .map(function (r) { return { id: String(r[0]), name: String(r[2]), address: String(r[3]) }; });
}

function init_(name) {
  var p = person_(name);
  return {
    name: name,
    fullName: p.fullName || PARTNERS[name].full || '',
    homeSet: !!p.home,
    homeMasked: maskAddress_(p.home),
    places: placesOf_(name),
    rate: RATE_PER_KM,
    from: PARTNERS[name].from || '',
    maxStops: MAX_STOPS
  };
}

function saveProfile_(name, fullName, homeAddress) {
  var p = person_(name);
  var sh = sheet_(SH.people);
  var full = fullName === undefined || fullName === null ? p.fullName : String(fullName).trim();
  if (full.length > 30) throw new Error('氏名が長すぎます（30文字まで）。');
  var home = p.home;
  if (homeAddress !== undefined && homeAddress !== null && String(homeAddress).trim() !== '') {
    home = cleanAddress_(homeAddress);
    if (!geocode_(home).length) throw new Error('この住所はGoogleマップで見つかりませんでした。');
  }
  var row = [name, full, home, now_()];
  if (p.row) sh.getRange(p.row, 1, 1, 4).setValues([row]);
  else sh.appendRow(row);
  return init_(name);
}

function savePlace_(name, id, placeName, address) {
  placeName = String(placeName || '').trim();
  address = cleanAddress_(address);
  if (!placeName) throw new Error('場所の名前を入れてください。');
  if (placeName.length > 30) throw new Error('場所の名前が長すぎます（30文字まで）。');
  if (placeName === '自宅') throw new Error('「自宅」は場所の名前に使えません（自宅は設定で登録します）。');
  if (!address) throw new Error('住所を選んでください。');
  var mine = placesOf_(name);
  mine.forEach(function (p) {
    if (p.name === placeName && p.id !== id) throw new Error('「' + placeName + '」はもう登録してあります。');
  });
  if (!geocode_(address).length) throw new Error('この住所はGoogleマップで見つかりませんでした。');
  var sh = sheet_(SH.places);
  if (id) {
    var rs = rows_(SH.places);
    for (var i = 0; i < rs.length; i++) {
      if (String(rs[i][0]) === id && rs[i][1] === name && !rs[i][5]) {
        sh.getRange(i + 2, 3, 1, 2).setValues([[placeName, address]]);
        return init_(name).places;
      }
    }
    throw new Error('直そうとした場所が見つかりません。画面を開き直してください。');
  }
  if (mine.length >= 100) throw new Error('登録できる場所は100か所までです。');
  sh.appendRow([newId_(), name, placeName, address, now_(), '']);
  return init_(name).places;
}

function deletePlace_(name, id) {
  var rs = rows_(SH.places);
  for (var i = 0; i < rs.length; i++) {
    if (String(rs[i][0]) === id && rs[i][1] === name && !rs[i][5]) {
      // 消しても、これまでの記録の経路はそのまま残る（記録には名前と住所を写してある）
      sheet_(SH.places).getRange(i + 2, 6).setValue('削除');
      return init_(name).places;
    }
  }
  throw new Error('消そうとした場所が見つかりません。画面を開き直してください。');
}

// ── 記録 ────────────────────────────────────────────────
function saveDays_(name, req) {
  var dates = (req.dates || []).map(checkDate_);
  if (!dates.length) throw new Error('日付を選んでください。');
  if (dates.length > 31) throw new Error('一度に入れられるのは31日分までです。');
  var ym = dates[0].slice(0, 7);
  dates.forEach(function (d) { if (d.slice(0, 7) !== ym) throw new Error('違う月の日付が混ざっています。'); });
  checkMonth_(name, ym);
  var highway = !!req.highway;
  var toll = yen_(req.toll, '高速代');
  var parking = yen_(req.parking, '駐車場代');
  if (toll && !highway) throw new Error('高速代を入れるときは「高速を使った」にしてください。');
  var note = String(req.note || '').trim();
  if (note.length > 200) throw new Error('備考が長すぎます（200文字まで）。');

  var r = route_(name, req.stops, highway);
  var stopsJson = JSON.stringify(r.stops.map(function (s) { return { n: s.name, a: s.address }; }));
  var legsJson = JSON.stringify(r.legs.map(function (l) { return l.km; }));
  var sh = sheet_(SH.records);
  var rs = rows_(SH.records);
  var at = now_();
  var byDate = {};
  rs.forEach(function (row, i) { if (row[R.NAME] === name) byDate[dateText_(row[R.DATE])] = i + 2; });

  dates.forEach(function (d) {
    var row = [newId_(), name, d, stopsJson, routeText_(r.stops), legsJson, r.km, highway ? '有' : '',
      r.gas, toll, parking, r.gas + toll + parking, note, at];
    if (byDate[d]) {
      row[0] = String(sh.getRange(byDate[d], 1).getValue() || row[0]);   // 同じ日は上書き（IDはそのまま）
      sh.getRange(byDate[d], 1, 1, row.length).setValues([row]);
    } else {
      sh.appendRow(row);
      byDate[d] = sh.getLastRow();
    }
  });
  return { saved: dates.length, route: { text: routeText_(r.stops), legs: r.legs, km: r.km, gas: r.gas } };
}

function entriesOf_(name, ym) {
  return rows_(SH.records).filter(function (r) {
    return r[R.NAME] === name && dateText_(r[R.DATE]).slice(0, 7) === ym;
  }).map(function (r) {
    var stops = [];
    try { stops = JSON.parse(r[R.STOPS]).map(function (s) { return s.n; }); } catch (e) {}
    var legs = [];
    try { legs = JSON.parse(r[R.LEGS]); } catch (e) {}
    return {
      id: String(r[R.ID]), date: dateText_(r[R.DATE]), stops: stops, text: String(r[R.TEXT]),
      legs: legs, km: Number(r[R.KM]) || 0, highway: r[R.HW] === '有',
      gas: Number(r[R.GAS]) || 0, toll: Number(r[R.TOLL]) || 0, parking: Number(r[R.PARK]) || 0,
      total: Number(r[R.TOTAL]) || 0, note: String(r[R.NOTE] || '')
    };
  }).sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
}

function totals_(list) {
  var t = { days: list.length, routes: 0, km: 0, gas: 0, toll: 0, parking: 0, total: 0 };
  list.forEach(function (e) {
    if (e.stops.length > 3) t.routes++;
    t.km += e.km; t.gas += e.gas; t.toll += e.toll; t.parking += e.parking; t.total += e.total;
  });
  t.km = Math.round(t.km * 10) / 10;
  return t;
}

function listMonth_(name, ym) {
  checkMonth_(name, ym);
  var list = entriesOf_(name, ym);
  return { ym: ym, entries: list, totals: totals_(list) };
}

function deleteEntry_(name, id) {
  var rs = rows_(SH.records);
  for (var i = 0; i < rs.length; i++) {
    if (String(rs[i][R.ID]) === String(id) && rs[i][R.NAME] === name) {
      checkMonth_(name, dateText_(rs[i][R.DATE]).slice(0, 7));
      sheet_(SH.records).deleteRow(i + 2);
      return true;
    }
  }
  throw new Error('消そうとした記録が見つかりません。一覧を読み直してください。');
}

// みんなの状況（team.html）用：その月の入力日数・経由地ありの日数・金額
function summary_(name, ym) {
  var t = totals_(entriesOf_(name, ym));
  return { ym: ym, days: t.days, routes: t.routes, amount: t.total };
}

// ── PDF ─────────────────────────────────────────────────
function pdfName_(name, ym, kind) { return ym + '_' + name + '_' + kind + '.pdf'; }

function monthFolder_(ym) {
  var root;
  try {
    root = DriveApp.getFolderById(PDF_FOLDER_ID);
  } catch (err) {
    throw new Error('PDFの保存先（パートナー）フォルダにアクセスできません。スクリプトを実行するアカウントに編集権限があるか確認してください。');
  }
  return folder_(folder_(root, ym.slice(0, 4) + '年'), Number(ym.slice(5)) + '月');
}

function folder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function buildPdfSheet_(name, ym) {
  var list = entriesOf_(name, ym);
  if (!list.length) throw new Error(Number(ym.slice(5)) + '月の入力がまだありません。');
  var t = totals_(list);
  var p = person_(name);
  var full = p.fullName || PARTNERS[name].full || name;
  var sh = sheet_(SH.pdf);
  sh.clear();
  var W = 9;
  var values = [
    ['交通費精算書（通勤・稼働）', '', '', '', '', '', '', '', ''],
    ['対象月', Number(ym.slice(0, 4)) + '年' + Number(ym.slice(5)) + '月', '', '氏名', full, '', '稼働日数', t.days + '日', ''],
    ['自宅と稼働場所の距離はGoogleマップで計算。ガソリン代は1kmあたり' + RATE_PER_KM + '円、高速代・駐車場代は実費。', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', ''],
    ['日付', '曜', '経路', '距離(km)', 'ガソリン代(円)', '高速代(円)', '駐車場代(円)', '合計(円)', '備考']
  ];
  list.forEach(function (e) {
    var d = new Date(Number(e.date.slice(0, 4)), Number(e.date.slice(5, 7)) - 1, Number(e.date.slice(8)));
    values.push([(d.getMonth() + 1) + '/' + d.getDate(), WEEK[d.getDay()], e.text + (e.highway ? '（高速）' : ''),
      e.km, e.gas, e.toll || '', e.parking || '', e.total, e.note]);
  });
  values.push(['合計', '', '', t.km, t.gas, t.toll, t.parking, t.total, '']);
  values.push(['', '', '', '', '', '', '', '', '']);
  values.push(['※月末締め、翌月3日までに申請。ETCの利用明細・駐車場の領収書は同じフォルダに保存。', '', '', '', '', '', '', '', '']);
  sh.getRange(1, 1, values.length, W).setValues(values);

  var first = 6, last = 5 + list.length, sum = last + 1;
  sh.getRange(1, 1, 1, W).merge().setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, 1, 1, W).setFontWeight('bold');
  sh.getRange(3, 1, 1, W).merge().setFontSize(9).setFontColor('#555555');
  sh.getRange(values.length, 1, 1, W).merge().setFontSize(9).setFontColor('#555555');
  sh.getRange(5, 1, 1, W).setFontWeight('bold').setBackground('#EEF1F5').setHorizontalAlignment('center');
  sh.getRange(5, 1, sum - 4, W).setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(first, 3, list.length, 1).setWrap(true);
  sh.getRange(first, 9, list.length, 1).setWrap(true);
  sh.getRange(first, 4, sum - first + 1, 1).setNumberFormat('0.0');
  sh.getRange(first, 5, sum - first + 1, 4).setNumberFormat('#,##0');
  sh.getRange(first, 1, sum - first + 1, 2).setHorizontalAlignment('center');
  sh.getRange(sum, 1, 1, W).setFontWeight('bold').setBackground('#F6F7F9');
  [52, 30, 330, 70, 90, 80, 85, 85, 150].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  SpreadsheetApp.flush();
  return { sheet: sh, lastRow: values.length, lastCol: W };
}

function exportPdf_(name, ym) {
  checkMonth_(name, ym);
  var b = buildPdfSheet_(name, ym);
  var params = {
    exportFormat: 'pdf', format: 'pdf', gid: b.sheet.getSheetId(), size: 'A4', portrait: 'false', scale: 4,
    r1: 0, c1: 0, r2: b.lastRow, c2: b.lastCol,
    top_margin: 0.3, bottom_margin: 0.3, left_margin: 0.3, right_margin: 0.3,
    horizontal_alignment: 'CENTER', vertical_alignment: 'TOP',
    sheetnames: 'false', printtitle: 'false', pagenumbers: 'false', gridlines: 'false', fzr: 'false'
  };
  var query = Object.keys(params).map(function (k) { return k + '=' + encodeURIComponent(params[k]); }).join('&');
  var res = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/export?' + query, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('PDFの書き出しに失敗しました（HTTP ' + res.getResponseCode() + '）。');
  var blob = res.getBlob();
  if (String(blob.getContentType() || '').indexOf('pdf') === -1) throw new Error('PDFではなくHTMLが返されました。デプロイの「実行するユーザー」を確認してください。');
  var fileName = pdfName_(name, ym, '稼働費');
  var folder = monthFolder_(ym);
  var old = folder.getFilesByName(fileName);
  while (old.hasNext()) old.next().setTrashed(true);
  folder.createFile(blob.setName(fileName));
  return { name: fileName };
}

// ── 添付（ETCの利用明細・駐車場の領収書） ─────────────────
// 名前は 2026-12_鈴木_ETC.pdf、2件目からは _2。1件目を受け取ったときに、
// その月・その人の同じ種類の古い添付をゴミ箱へ移す（送り直しても混ざらない）。
var ATTACH_KINDS = { ETC: 'ETC', parking: '駐車場' };
var ATTACH_TYPES = { 'application/pdf': 'pdf', 'image/jpeg': 'jpg', 'image/png': 'png' };
var ATTACH_MAX_BYTES = 10 * 1024 * 1024;

function saveAttachment_(name, ym, kind, index, mimeType, base64) {
  checkMonth_(name, ym);
  var label = ATTACH_KINDS[kind];
  if (!label) throw new Error('添付の種類が分かりません: ' + kind);
  var ext = ATTACH_TYPES[mimeType];
  if (!ext) throw new Error('PDFか写真（JPEG・PNG）を添付してください。');
  index = Number(index);
  if (!(index >= 1 && index <= 20 && Math.floor(index) === index)) throw new Error('添付の番号が正しくありません。');
  var bytes = Utilities.base64Decode(String(base64 || ''));
  if (!bytes.length) throw new Error('添付ファイルが空です。');
  if (bytes.length > ATTACH_MAX_BYTES) throw new Error('添付ファイルが大きすぎます（10MBまで）。');

  var folder = monthFolder_(ym);
  var base = ym + '_' + name + '_' + label;
  var fileName = base + (index > 1 ? '_' + index : '') + '.' + ext;
  var stem = fileName.replace(/\.[^.]+$/, '');
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var n = f.getName();
    var sameKind = n.indexOf(base + '.') === 0 || n.indexOf(base + '_') === 0;
    if (sameKind && (index === 1 || n.replace(/\.[^.]+$/, '') === stem)) f.setTrashed(true);
  }
  folder.createFile(Utilities.newBlob(bytes, mimeType, fileName));
  return { name: fileName };
}

// ── Webアプリ ───────────────────────────────────────────
function doGet() {
  return json_({ ok: true, data: { status: 'ready', version: CODE_VERSION, partners: Object.keys(PARTNERS), time: new Date().toISOString() } });
}

// 書き込まない操作（ロックをかけない）。previewRoute は距離メモに足すことはあるが、
// 同じ行が2つできても困らないのでロックしない。
var READ_ONLY = { init: true, searchAddress: true, previewRoute: true, listMonth: true, summary: true };

function doPost(e) {
  var out, lock = null;
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('リクエスト本文が空です');
    var q = JSON.parse(e.postData.contents);
    var name = checkName_(q.name);
    if (!READ_ONLY[q.action]) {
      lock = LockService.getScriptLock();
      if (!lock.tryLock(25000)) {
        lock = null;
        throw new Error('ほかの人の処理が終わっていません。少し待ってからもう一度お試しください。');
      }
    }
    var data;
    switch (q.action) {
      case 'init':          data = init_(name); break;
      case 'searchAddress': data = geocode_(q.query); break;
      case 'saveProfile':   data = saveProfile_(name, q.fullName, q.homeAddress); break;
      case 'savePlace':     data = savePlace_(name, q.id || '', q.placeName, q.address); break;
      case 'deletePlace':   data = deletePlace_(name, String(q.id || '')); break;
      case 'previewRoute':  var r = route_(name, q.stops, !!q.highway);
                            data = { text: routeText_(r.stops), legs: r.legs, km: r.km, gas: r.gas }; break;
      case 'saveDays':      data = saveDays_(name, q); break;
      case 'listMonth':     data = listMonth_(name, String(q.ym || '')); break;
      case 'deleteEntry':   data = deleteEntry_(name, q.id); break;
      case 'exportPdf':     data = exportPdf_(name, String(q.ym || '')); break;
      case 'saveAttachment': data = saveAttachment_(name, String(q.ym || ''), q.kind, q.index, q.mimeType, q.data); break;
      case 'summary':       data = summary_(name, String(q.ym || '')); break;
      default: throw new Error('不明なaction: ' + q.action);
    }
    if (lock) SpreadsheetApp.flush();
    out = { ok: true, data: data };
  } catch (err) {
    out = { ok: false, error: err.message };
  } finally {
    if (lock) lock.releaseLock();
  }
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
