// ============================================================
//  稼働費計算スプレッドシート - Google Apps Script
//
//  【セットアップ手順】
//  1. スプレッドシートをGoogleドライブにアップロード
//     → 右クリック →「Googleスプレッドシートで開く」
//  2. 拡張機能 → Apps Script → このコードを貼り付けて保存
//  3. スプレッドシートを再読み込みすると「📋 稼働費管理」メニューが出る
//  4. 「⚙️ 初期設定（初回のみ）」を実行 → Maps API の権限を許可
//     ※このとき自動計算用の編集トリガー（インストール型）も登録されます。
//     　これが無いと C列/I列を選んでも距離が自動で入らなかったり、
//     　権限エラーで距離取得に失敗したりします。
//  5. 以降は月を変えたら「🗓️ 月を更新」を実行
//
//  【日々の使い方】
//  ・C列（稼働場所）をプルダウンで選択
//    → I列（高速 有/無）をプルダウンで選択
//    → 距離が自動で入ります
//  ・G列・H列は高速代（円）を手入力
//
//  【Web入力フォーム（GitHub Pages）対応】
//  ・doPost経由でスプレッドシートを直接触らずに入力できる関数群を追加
//  ・PDF保存先フォルダIDは PDF_ROOT_FOLDER_ID に設定済み
//  ・稼働場所の追加・編集・削除もWebフォームから可能
//
//  【Webフォームで「Unexpected token '<'」が出る場合】
//  JSONではなくHTML（Googleのログイン画面や権限エラー画面）が返っています。
//  doPostの中で起きた例外はJSONで返るため、これはコードではなくデプロイ設定側の問題です。
//  ・デプロイを管理 → 実行するユーザー: 自分
//  ・デプロイを管理 → アクセスできるユーザー: 全員
//  ・コードに新しいサービス（DriveApp/UrlFetchApp等）を足したら必ず再デプロイ
//  デプロイURL（末尾 /exec）をブラウザで直接開くと doGet が状態JSONを返すので、
//  そこでログイン画面が出るなら上記のアクセス設定が原因です。
// ============================================================

const CONFIG = {
  year: 2026,
  ratePerKm: 15,       // 円/km
  dataStartRow: 6,     // データ開始行
  MONTH_ROW: 2,
  MONTH_COL: 7,        // G2: 月入力セル
  IRREGULAR_ROWS: 3,   // イレギュラー（稼働場所以外）の自由記述行数

  // 列番号（1始まり）
  COL: {
    DATE:       1,   // A: 日付
    WEEKDAY:    2,   // B: 曜日
    PLACE:      3,   // C: 稼働場所（プルダウン）
    DIST_GO:    4,   // D: 行き距離(km)
    DIST_BACK:  5,   // E: 戻り距離(km)
    DIST_TOTAL: 6,   // F: 合計距離(km)
    TOLL_GO:    7,   // G: 高速代 行き(円)
    TOLL_BACK:  8,   // H: 高速代 戻り(円)
    HIGHWAY:    9,   // I: 高速 有/無（プルダウン）
    TOTAL_AMT:  10,  // J: 合計金額(円)
    NOTE:       11,  // K: 備考
  },

  // スタッフ名 → 自宅の短縮名（稼働場所リストと一致させる）
  STAFF_HOME: {
    '福山康平': '福山自宅',
    '古川泰成': '古川自宅',
    '清田和男': '清田自宅',
  },

  PLACE_SHEET: '稼働場所',  // 場所マスタシート名

  // PDF出力設定
  PDF: {
    size: 'A4',
    portrait: false,  // false = 横向き
    scale: 4,         // 1:100% / 2:幅に合わせる / 3:高さに合わせる / 4:1ページに収める
    margin: 0.25,     // 余白（インチ）。小さいほど本文を大きく印刷できる
  },

  EDIT_TRIGGER_HANDLER: 'handleEdit', // インストール型onEditトリガーのハンドラ名
  ROUTE_CACHE_SEC: 21600,             // 同一経路の距離キャッシュ保持時間（秒・最大6時間）
};

// ▼▼▼ PDF保存先のGoogle DriveフォルダID ▼▼▼
const PDF_ROOT_FOLDER_ID = "1VreXxTQpiTeiAoc_uxsLrpj1KloJ7V2o";

const WEEKDAY_JA = ['日','月','火','水','木','金','土'];

// 距離取得に失敗した時に備考(K列)へ書くメッセージの接頭辞。
// 手入力の備考を消さずに、スクリプトが書いたものだけをクリアするための目印。
const DIST_NOTE_PREFIXES = ['距離取得エラー', '自宅未登録', '住所未登録'];

// ── カスタムメニュー ─────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 稼働費管理')
    .addItem('🗓️ 月を更新（日付を再生成）', 'updateMonth')
    .addSeparator()
    .addItem('📍 距離を再取得（全行）', 'refreshAllDistances')
    .addSeparator()
    .addItem('⚙️ 初期設定（初回のみ）', 'initialSetup')
    .addItem('🔌 自動計算トリガーを再設定', 'installEditTriggerFromMenu')
    .addSeparator()
    .addItem('🩺 Drive権限をチェック', 'checkDriveAccess')
    .addItem('🩺 PDF保存をテスト', 'testExportPdf')
    .addToUi();
}

// ── 初期設定 ─────────────────────────────────────────────────
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (!ss.getSheetByName(CONFIG.PLACE_SHEET)) {
    ui.alert('「稼働場所」シートが見つかりません。');
    return;
  }

  Object.keys(CONFIG.STAFF_HOME).forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;
    setupDropdowns(sheet);
  });

  // C列/I列を編集した時にMapsで距離を自動計算するための
  // インストール型トリガーを登録（単純トリガーのonEditだとMapsサービスが権限エラーになるため）
  installEditTrigger();

  ui.alert('✅ 初期設定完了！\n\nG2セルの月を確認して「月を更新」を実行してください。\n\n以降はC列（場所）→ I列（高速有無）の順に選ぶと距離が自動入力されます。');
}

function installEditTriggerFromMenu() {
  installEditTrigger();
  SpreadsheetApp.getUi().alert('✅ 自動計算トリガーを再設定しました。');
}

// C列/I列の編集を検知するインストール型トリガーを（再）登録する。
// 同名ハンドラの重複トリガーは一旦削除してから作り直すため、何度実行しても安全。
function installEditTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === CONFIG.EDIT_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger(CONFIG.EDIT_TRIGGER_HANDLER)
    .forSpreadsheet(ss)
    .onEdit()
    .create();
}

// メッセージをUIに出す。エディタから実行した時はUIが無く getUi() が例外を投げるので、
// その場合はログだけに残す。診断用の関数をメニューからでもエディタからでも
// 同じように実行できるようにするための小道具。
function report_(msg) {
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    // エディタ実行時はUIが使えないのでログのみ。ここで失敗させない。
  }
  return msg;
}

// Drive権限だけを切り分ける診断用。引数なしなのでエディタから直接実行できる。
// 「DriveApp.getFolderById を呼び出す権限がありません」が出るかどうかを、
// PDF書き出しやスタッフ名の指定と無関係に単体で確認できる。
function checkDriveAccess() {
  try {
    const folder = DriveApp.getFolderById(PDF_ROOT_FOLDER_ID);
    return report_('✅ Drive OK: フォルダにアクセスできました →「' + folder.getName() + '」');
  } catch (err) {
    return report_('❌ Drive NG: ' + err.message);
  }
}

// PDF保存だけを試す診断用。Webフォーム経由だと通信エラーに埋もれて原因が見えないため、
// 権限やフォルダIDの問題をGAS側だけで切り分けられるようにしている。
// 引数なしで実行できるので、スプレッドシートのメニューからでも
// スクリプトエディタの「実行」からでも動く。
function testExportPdf() {
  let name = null;
  try {
    name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  } catch (e) {
    // エディタ実行ではアクティブシートを取れないことがあるので後段でフォールバック
  }
  // アクティブシートがスタッフシートでなければ、設定済みスタッフの先頭を使う
  if (!CONFIG.STAFF_HOME[name]) {
    name = Object.keys(CONFIG.STAFF_HOME)[0];
  }

  try {
    const res = exportCommutePdf(name);
    return report_('✅ ' + name + ': ' + res);
  } catch (err) {
    return report_('❌ ' + name + ': ' + err.message);
  }
}

// ── 月更新（スプレッドシートのメニューから実行する場合） ───────
function updateMonth() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ui  = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  const name  = sheet.getName();

  if (!CONFIG.STAFF_HOME[name]) {
    ui.alert('スタッフシートを選択してから実行してください。');
    return;
  }

  const month = Number(sheet.getRange(CONFIG.MONTH_ROW, CONFIG.MONTH_COL).getValue());
  if (!month || month < 1 || month > 12) {
    ui.alert('G2セルに月（1〜12）を入力してください。');
    return;
  }

  const res = ui.alert(`${month}月に更新`, `「${name}」を${month}月で更新します。\n既存データは上書きされます。よろしいですか？`, ui.ButtonSet.YES_NO);
  if (res !== ui.Button.YES) return;

  buildMonthRows(sheet, month);
  setupDropdowns(sheet);
  ui.alert(`✅ ${month}月のカレンダーを生成しました。\nC列で場所を選ぶと距離が自動で入ります。`);
}

// ── 月のカレンダー行を生成 ───────────────────────────────────
function buildMonthRows(sheet, month) {
  const year = CONFIG.year;
  const C    = CONFIG.COL;
  const startRow = CONFIG.dataStartRow;
  const daysInMonth = new Date(year, month, 0).getDate();
  const IRREG_ROWS = CONFIG.IRREGULAR_ROWS;

  const FIXED_DAYS    = 31;
  const headRow       = startRow + FIXED_DAYS;
  const irregStart    = headRow + 1;
  const irregEnd      = irregStart + IRREG_ROWS - 1;
  const totalRow       = irregEnd + 1;

  sheet.getRange(startRow, 1, FIXED_DAYS, 11).clearContent().clearFormat();

  const dataRange = sheet.getRange(startRow, 1, FIXED_DAYS, 11);
  dataRange.setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  for (let day = 1; day <= FIXED_DAYS; day++) {
    const row = startRow + day - 1;

    if (day > daysInMonth) {
      sheet.hideRows(row);
      continue;
    }
    sheet.showRows(row);

    sheet.getRange(row, C.DATE)
      .setFormula(`=DATE(${year},${month},${day})`)
      .setNumberFormat('m/d')
      .setHorizontalAlignment('center');
    const wd = new Date(year, month - 1, day).getDay();
    const isWe = wd === 0 || wd === 6;

    const wdColor = wd === 0 ? '#CC0000' : (wd === 6 ? '#0070C0' : '#333333');
    sheet.getRange(row, C.WEEKDAY)
      .setValue(WEEKDAY_JA[wd])
      .setFontColor(wdColor)
      .setHorizontalAlignment('center');

    if (isWe) {
      sheet.getRange(row, 1, 1, 11).setBackground('#F0F0F0');
    }

    sheet.getRange(row, C.TOLL_GO).setNumberFormat('#,##0').setBackground(isWe ? '#F0F0F0' : '#EBF5FB');
    sheet.getRange(row, C.TOLL_BACK).setNumberFormat('#,##0').setBackground(isWe ? '#F0F0F0' : '#EBF5FB');
  }

  ensureIrregularSection(sheet, headRow, irregStart, irregEnd);

  const lastCalRow = startRow + daysInMonth - 1;
  sheet.getRange(totalRow, 1, 1, 11).setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(totalRow, 1).setValue('合　　計').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(totalRow, 1, 1, 11).setBackground('#FFF2CC');
  sheet.getRange(totalRow, C.DIST_TOTAL).setFormula(`=SUM(F${startRow}:F${lastCalRow})`).setNumberFormat('#,##0.000"km"').setFontWeight('bold');
  sheet.getRange(totalRow, C.TOLL_GO).setFormula(`=SUM(G${startRow}:G${lastCalRow})`).setNumberFormat('#,##0"円"').setFontWeight('bold');
  sheet.getRange(totalRow, C.TOLL_BACK).setFormula(`=SUM(H${startRow}:H${lastCalRow})`).setNumberFormat('#,##0"円"').setFontWeight('bold');
  sheet.getRange(totalRow, C.TOTAL_AMT)
    .setFormula(`=SUM(J${startRow}:J${lastCalRow})+SUM(J${irregStart}:J${irregEnd})`)
    .setNumberFormat('#,##0"円"').setFontWeight('bold').setFontSize(13);

  sheet.getRange(2, 9).setFormula(`=COUNTA(C${startRow}:C${lastCalRow})`);
}

// ── イレギュラー枠（見出し+3行）の書式を保証 ─────────────────
function ensureIrregularSection(sheet, headRow, irregStart, irregEnd) {
  const C = CONFIG.COL;

  sheet.getRange(headRow, 1, 1, 11).breakApart();
  sheet.getRange(headRow, 1, 1, 11).merge();
  sheet.getRange(headRow, 1)
    .setValue('◆ イレギュラー（稼働場所以外の特別対応・出張等）　※自由記述')
    .setFontWeight('bold')
    .setFontColor('#FFFFFF')
    .setBackground('#C0504D')
    .setHorizontalAlignment('center')
    .setFontSize(9);
  sheet.getRange(headRow, 1, 1, 11).setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);

  for (let r = irregStart; r <= irregEnd; r++) {
    sheet.getRange(r, 1, 1, 2).breakApart();
    sheet.getRange(r, 1, 1, 2).merge();
    sheet.getRange(r, 1, 1, 2).setBackground('#FCE4D6').setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(r, 1).setNumberFormat('m/d');

    sheet.getRange(r, 3, 1, 4).breakApart();
    sheet.getRange(r, 3, 1, 4).merge();
    sheet.getRange(r, 3).setBackground('#FCE4D6').setHorizontalAlignment('left').setFontSize(10);
    sheet.getRange(r, 3, 1, 4).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(r, 7, 1, 2).breakApart();
    sheet.getRange(r, 7, 1, 2).merge();
    sheet.getRange(r, 7, 1, 2).setBackground('#FCE4D6').setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(r, 9).setBackground('#FCE4D6').setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(r, 10).setBackground('#FCE4D6').setNumberFormat('#,##0').setHorizontalAlignment('right').setFontSize(10);
    sheet.getRange(r, 10).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

    sheet.getRange(r, 11).setBackground('#FCE4D6').setHorizontalAlignment('left').setFontSize(10);
    sheet.getRange(r, 11).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  }
}

// ── プルダウン設定 ───────────────────────────────────────────
function setupDropdowns(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const placeSheet = ss.getSheetByName(CONFIG.PLACE_SHEET);
  if (!placeSheet) return;

  const startRow = CONFIG.dataStartRow;
  const C = CONFIG.COL;
  const FIXED_DAYS = 31;

  const lastPlaceRow = placeSheet.getLastRow();
  const placeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(placeSheet.getRange(`A2:A${lastPlaceRow}`), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(startRow, C.PLACE, FIXED_DAYS, 1).setDataValidation(placeRule);

  const hwRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['無', '有'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(startRow, C.HIGHWAY, FIXED_DAYS, 1).setDataValidation(hwRule);
}

// ── handleEdit: C列(場所)またはI列(高速有無)の変更を検知 ────
// ※インストール型トリガー（installEditTrigger）から呼ばれる。
//   関数名を"onEdit"にすると単純トリガーとしても自動実行され、
//   Mapsサービスが権限エラーで失敗する／二重実行される原因になるため、
//   あえて別名にしてインストール型トリガー経由でのみ実行する。
function handleEdit(e) {
  const sheet    = e.range.getSheet();
  const staffName = sheet.getName();
  if (!CONFIG.STAFF_HOME[staffName]) return;

  const editedCol = e.range.getColumn();
  const editedRow = e.range.getRow();
  const C = CONFIG.COL;
  const startRow     = CONFIG.dataStartRow;
  const lastDataRow  = startRow + 31 - 1;

  if ((editedCol === C.PLACE || editedCol === C.HIGHWAY)
      && editedRow >= startRow && editedRow <= lastDataRow) {
    updateRowDistance(sheet, staffName, editedRow);
  }

  if ((editedCol === C.TOLL_GO || editedCol === C.TOLL_BACK)
      && editedRow >= startRow && editedRow <= lastDataRow) {
    setTotalFormula(sheet, editedRow);
  }
}

// ── 1行分の距離を取得・セット ────────────────────────────────
// placeMap は呼び出し側から渡せる。複数日をまとめて処理する時に
// 1行ごとに「稼働場所」シート全体を読み直すと、日数分だけ往復が増えて
// 送信が極端に遅くなるため。
function updateRowDistance(sheet, staffName, row, placeMap) {
  const C = CONFIG.COL;
  // C列（場所）とI列（高速有無）を1回でまとめて読む
  const rowVals    = sheet.getRange(row, C.PLACE, 1, C.HIGHWAY - C.PLACE + 1).getValues()[0];
  const placeShort = rowVals[0];
  const highway    = rowVals[C.HIGHWAY - C.PLACE];

  if (!placeShort) {
    sheet.getRange(row, C.DIST_GO, 1, 3).clearContent();
    clearDistanceNote(sheet, row);
    return;
  }

  if (!placeMap) {
    placeMap = getPlaceMap(SpreadsheetApp.getActiveSpreadsheet());
  }

  const homeShort = CONFIG.STAFF_HOME[staffName];
  const homeAddr  = placeMap[homeShort];
  const destAddr  = placeMap[placeShort];

  if (!homeAddr) {
    clearDistanceCells(sheet, row);
    sheet.getRange(row, C.NOTE).setValue('自宅未登録');
    return;
  }
  if (!destAddr) {
    clearDistanceCells(sheet, row);
    sheet.getRange(row, C.NOTE).setValue('住所未登録');
    return;
  }

  const routeType = (highway === '有') ? 'highway' : 'drive';

  try {
    const distGo   = getRouteDistanceKm(homeAddr, destAddr, routeType);
    const distBack = getRouteDistanceKm(destAddr, homeAddr, routeType);

    // D・E・F はまとめて1回で書く（"=" で始まる文字列は数式として入る）
    sheet.getRange(row, C.DIST_GO, 1, 3)
      .setValues([[
        distGo,
        distBack,
        `=IF(AND(ISNUMBER(D${row}),ISNUMBER(E${row})),D${row}+E${row},"")`
      ]])
      .setNumberFormat('#,##0.000');
    clearDistanceNote(sheet, row);
    setTotalFormula(sheet, row);
  } catch(err) {
    // D列（数値列）にエラー文字列を入れると、合計行の SUM(F...) が
    // #VALUE! になり月全体の合計まで壊れてしまうため、
    // 距離セルは空にして、エラー内容は備考（K列）にだけ残す。
    clearDistanceCells(sheet, row);
    sheet.getRange(row, C.NOTE).setValue('距離取得エラー: ' + err.message);
    setTotalFormula(sheet, row);
    Logger.log(`Row ${row} エラー: ${err.message}`);
  }
}

function clearDistanceCells(sheet, row) {
  const C = CONFIG.COL;
  sheet.getRange(row, C.DIST_GO, 1, 3).clearContent(); // D,E,F
}

// 備考(K列)のうち、スクリプトが書いたエラーメッセージだけを消す。
// 手入力の備考を巻き込んで消さないよう、既知の接頭辞に一致する場合のみクリアする。
function clearDistanceNote(sheet, row) {
  const cell = sheet.getRange(row, CONFIG.COL.NOTE);
  const val  = String(cell.getValue() || '');
  if (!val) return;
  const isOurs = DIST_NOTE_PREFIXES.some(p => val.indexOf(p) === 0);
  if (isOurs) cell.clearContent();
}

// ── 合計金額の数式をセット ───────────────────────────────────
function setTotalFormula(sheet, row) {
  const C = CONFIG.COL;
  sheet.getRange(row, C.TOTAL_AMT)
    .setFormula(
      `=IF(C${row}="","",` +
      `IF(ISNUMBER(F${row}),F${row}*${CONFIG.ratePerKm},0)` +
      `+IF(ISNUMBER(G${row}),G${row},0)` +
      `+IF(ISNUMBER(H${row}),H${row},0))`
    )
    .setNumberFormat('#,##0');
}

// ── 全行の距離を一括再取得 ──────────────────────────────────
function refreshAllDistances() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ui   = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();
  const name  = sheet.getName();

  if (!CONFIG.STAFF_HOME[name]) {
    ui.alert('スタッフシートを選択してから実行してください。');
    return;
  }

  const month = Number(sheet.getRange(CONFIG.MONTH_ROW, CONFIG.MONTH_COL).getValue());
  const daysInMonth = new Date(CONFIG.year, month, 0).getDate();
  const startRow    = CONFIG.dataStartRow;
  const C = CONFIG.COL;

  // 場所マスタとC列は1回ずつまとめて読む
  const placeMap = getPlaceMap(ss);
  const placeCol = sheet.getRange(startRow, C.PLACE, daysInMonth, 1).getValues();

  let updated = 0;
  for (let day = 1; day <= daysInMonth; day++) {
    const row = startRow + day - 1;
    if (!placeCol[day - 1][0]) continue;

    updateRowDistance(sheet, name, row, placeMap);
    updated++;

    // API レート制限対策（キャッシュヒット時はmapQuery自体が呼ばれないので実質スキップされる）
    if (updated % 5 === 0) Utilities.sleep(1000);
  }

  ui.alert(`✅ ${updated}件の距離を更新しました。`);
}

// ── 場所マスタ読み込み（短縮名 → 住所） ─────────────────────
function getPlaceMap(ss) {
  const sheet = ss.getSheetByName(CONFIG.PLACE_SHEET);
  if (!sheet) return {};
  const rows = sheet.getDataRange().getValues();
  const map  = {};
  for (let i = 1; i < rows.length; i++) {
    const short = String(rows[i][0]).trim();
    const addr  = String(rows[i][1]).trim();
    if (short && addr) map[short] = addr;
  }
  return map;
}

// ── 経路距離の取得（キャッシュ付き） ─────────────────────────
// 同じスタッフは同じ経路（自宅⇔現場）を毎日往復するため、
// 一度取得した距離をキャッシュして無駄なMaps API呼び出し・待ち時間を減らす。
// 一時的なエラー（レート制限等）は短い間隔でリトライする。
function getRouteDistanceKm(origin, dest, routeType) {
  const cache = CacheService.getScriptCache();
  const key   = routeCacheKey(origin, dest, routeType);
  const cached = cache.get(key);
  if (cached !== null) return Number(cached);

  let lastErr;
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const km = mapQuery(origin, dest, routeType, 'distance');
      cache.put(key, String(km), CONFIG.ROUTE_CACHE_SEC);
      return km;
    } catch (err) {
      lastErr = err;
      Utilities.sleep(500 * (attempt + 1));
    }
  }
  throw lastErr;
}

function routeCacheKey(origin, dest, routeType) {
  const raw = `${origin}::${dest}::${routeType}`;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8);
  const hex = digest.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0')).join('');
  return 'route_' + hex;
}

// ── mapQuery: Google Maps で距離・時間を取得 ─────────────────
function mapQuery(src, dest, type, result) {
  const finder = Maps.newDirectionFinder()
    .setOrigin(src)
    .setDestination(dest)
    .setLanguage('ja');

  switch (type) {
    case 'highway':
      finder.setMode(Maps.DirectionFinder.Mode.DRIVING);
      break;
    case 'drive':
      finder.setMode(Maps.DirectionFinder.Mode.DRIVING)
            .setAvoid(Maps.DirectionFinder.Avoid.TOLLS);
      break;
    case 'toll':
      finder.setMode(Maps.DirectionFinder.Mode.DRIVING)
            .setAvoid(Maps.DirectionFinder.Avoid.HIGHWAYS);
      break;
    case 'transit':
      finder.setMode(Maps.DirectionFinder.Mode.TRANSIT);
      break;
    case 'bicycle':
      finder.setMode(Maps.DirectionFinder.Mode.BICYCLING);
      break;
    case 'walk':
      finder.setMode(Maps.DirectionFinder.Mode.WALKING);
      break;
    default:
      finder.setMode(Maps.DirectionFinder.Mode.DRIVING);
  }

  const directions = finder.getDirections();
  if (!directions || !directions.routes || !directions.routes.length) {
    throw new Error('経路が見つかりません（住所を確認してください）');
  }
  const route = directions.routes[0];
  const leg   = route.legs[0];

  if (result === 'distance') return leg.distance.value / 1000;
  if (result === 'duration') return leg.duration.value / 60;
  return leg;
}


// ============================================================
//  ここから下：GitHub Pages版フロントエンド（fetch通信）対応
// ============================================================

// ── デプロイ確認用（ブラウザで /exec を直接開いた時に返る） ────
// ここでログイン画面やエラーHTMLが出る場合、原因はコードではなく
// デプロイ設定（アクセスできるユーザー）です。
function doGet(e) {
  var info = { status: 'ready', time: new Date().toISOString() };
  try {
    info.effectiveUser = Session.getEffectiveUser().getEmail();
  } catch (err) {
    info.effectiveUser = '(取得不可)';
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, data: info }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Webアプリのエントリーポイント ───────────────────────────
function doPost(e) {
  var out;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('リクエスト本文が空です');
    }
    var req = JSON.parse(e.postData.contents);
    var action = req.action;
    var data;

    if (action === 'getStaffList') {
      data = Object.keys(CONFIG.STAFF_HOME);
    } else if (action === 'getInitialData') {
      data = getInitialData(req.staffName);
    } else if (action === 'getMonthInfo') {
      data = getMonthInfoWeb(req.staffName);
    } else if (action === 'getPlaceList') {
      data = getPlaceListWeb();
    } else if (action === 'getPlaceListManage') {
      data = getPlaceListManage();
    } else if (action === 'addPlace') {
      data = addPlaceWeb(req.name, req.address);
    } else if (action === 'updatePlace') {
      data = updatePlaceWeb(req.row, req.name, req.address);
    } else if (action === 'deletePlace') {
      data = deletePlaceWeb(req.row);
    } else if (action === 'submitCommute') {
      data = submitCommute(req.staffName, req.data);
    } else if (action === 'submitIrregular') {
      data = submitIrregular(req.staffName, req.data);
    } else if (action === 'getMyEntries') {
      data = getMyEntries(req.staffName);
    } else if (action === 'deleteNormalEntry') {
      data = deleteNormalEntry(req.staffName, req.day);
    } else if (action === 'deleteIrregularEntry') {
      data = deleteIrregularEntry(req.staffName, req.row);
    } else if (action === 'resetMonth') {
      data = resetMonthWeb(req.staffName, req.month);
    } else if (action === 'exportPdf') {
      data = exportCommutePdf(req.staffName);
    } else {
      throw new Error('不明なaction: ' + action);
    }

    out = { ok: true, data: data };
  } catch (err) {
    out = { ok: false, error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── スタッフ名からシートを取得 ────────────────────────────
function getCommuteSheet(staffName) {
  if (!CONFIG.STAFF_HOME[staffName]) {
    throw new Error('不明なスタッフです: ' + staffName);
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(staffName);
  if (!sheet) {
    throw new Error('シートが見つかりません: ' + staffName);
  }
  return sheet;
}

// ── イレギュラー行の範囲（固定計算） ─────────────────────────
function getIrregularRowRange() {
  var startRow = CONFIG.dataStartRow;      // 6
  var FIXED_DAYS = 31;
  var headRow = startRow + FIXED_DAYS;     // 37
  var irregStart = headRow + 1;            // 38
  var irregEnd = irregStart + CONFIG.IRREGULAR_ROWS - 1; // 40
  return { irregStart: irregStart, irregEnd: irregEnd };
}

// ── 起動時に必要な情報をまとめて返す ─────────────────────────
// 月情報と場所一覧を別々のリクエストで取ると、GASの実行が2本並行して走り
// 初回読み込みがタイムアウトしやすい。1回の呼び出しにまとめる。
function getInitialData(staffName) {
  return {
    monthInfo: getMonthInfoWeb(staffName),
    placeList: getPlaceListWeb()
  };
}

// ── 月情報＋入力済みの日を返す（フォーム表示用） ─────────────
function getMonthInfoWeb(staffName) {
  var sheet = getCommuteSheet(staffName);
  var month = Number(sheet.getRange(CONFIG.MONTH_ROW, CONFIG.MONTH_COL).getValue());
  var daysInMonth = new Date(CONFIG.year, month, 0).getDate();
  var C = CONFIG.COL;
  var range = getIrregularRowRange();

  // C列（通常31日分＋イレギュラー枠）を1回でまとめて読む。
  // 1セルずつ getValue() すると往復が数十回発生し、初回読み込みが
  // タイムアウトする原因になる。
  var firstRow  = CONFIG.dataStartRow;
  var rowCount  = range.irregEnd - firstRow + 1;
  var placeCol  = sheet.getRange(firstRow, C.PLACE, rowCount, 1).getValues();

  var filledDays = [];
  for (var d = 1; d <= daysInMonth; d++) {
    if (placeCol[d - 1][0]) filledDays.push(d);
  }

  var irregularCount = 0;
  for (var r = range.irregStart; r <= range.irregEnd; r++) {
    if (placeCol[r - firstRow][0]) irregularCount++;
  }

  return {
    year: CONFIG.year,
    month: month,
    daysInMonth: daysInMonth,
    filledDays: filledDays,
    irregularCount: irregularCount,
    irregularCapacity: CONFIG.IRREGULAR_ROWS
  };
}

// ── 稼働場所の一覧を返す（自宅は除外、Web入力タブ用） ────────
function getPlaceListWeb() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.PLACE_SHEET);
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    var name = String(rows[i][0]).trim();
    if (name && name.indexOf('自宅') === -1) list.push(name);
  }
  return list;
}

// ── 稼働場所マスタの管理（一覧・追加・編集・削除） ─────────────
function getPlaceListManage() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.PLACE_SHEET);
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < rows.length; i++) {
    var name = String(rows[i][0]).trim();
    var addr = String(rows[i][1]).trim();
    if (!name) continue;
    if (name.indexOf('自宅') !== -1) continue; // 自宅はスタッフ設定に連動するため管理対象外
    list.push({ row: i + 1, name: name, address: addr });
  }
  return list;
}

function addPlaceWeb(name, address) {
  name = String(name || '').trim();
  address = String(address || '').trim();
  if (!name) throw new Error('場所の名称を入力してください');
  if (!address) throw new Error('住所を入力してください');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.PLACE_SHEET);
  if (!sheet) throw new Error('「稼働場所」シートが見つかりません');

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === name) {
      throw new Error('「' + name + '」は既に登録されています');
    }
  }

  sheet.appendRow([name, address]);
  refreshAllStaffDropdowns();
  return 'OK';
}

function updatePlaceWeb(row, name, address) {
  row = Number(row);
  name = String(name || '').trim();
  address = String(address || '').trim();
  if (!name) throw new Error('場所の名称を入力してください');
  if (!address) throw new Error('住所を入力してください');

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.PLACE_SHEET);
  if (!sheet) throw new Error('「稼働場所」シートが見つかりません');
  if (row < 2 || row > sheet.getLastRow()) throw new Error('対象の行が不正です');

  var oldName = String(sheet.getRange(row, 1).getValue()).trim();

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (i + 1 !== row && String(rows[i][0]).trim() === name) {
      throw new Error('「' + name + '」は既に登録されています');
    }
  }

  sheet.getRange(row, 1).setValue(name);
  sheet.getRange(row, 2).setValue(address);

  // 名称が変わった場合、既存の稼働実績（C列）の表記も追随させる
  if (oldName && oldName !== name) {
    renamePlaceInStaffSheets(oldName, name);
  }

  refreshAllStaffDropdowns();
  return 'OK';
}

function deletePlaceWeb(row) {
  row = Number(row);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.PLACE_SHEET);
  if (!sheet) throw new Error('「稼働場所」シートが見つかりません');
  if (row < 2 || row > sheet.getLastRow()) throw new Error('対象の行が不正です');

  var name = String(sheet.getRange(row, 1).getValue()).trim();
  if (name.indexOf('自宅') !== -1) {
    throw new Error('自宅の設定はここからは削除できません');
  }

  sheet.deleteRow(row);
  refreshAllStaffDropdowns();
  return 'OK';
}

// ── 全スタッフシートのプルダウンを再設定 ─────────────────────
function refreshAllStaffDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(CONFIG.STAFF_HOME).forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) setupDropdowns(sheet);
  });
}

// ── 稼働場所名の変更を各スタッフシートのC列に反映 ─────────────
function renamePlaceInStaffSheets(oldName, newName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var C = CONFIG.COL;
  var startRow = CONFIG.dataStartRow;
  var FIXED_DAYS = 31;

  Object.keys(CONFIG.STAFF_HOME).forEach(function(staffName) {
    var sheet = ss.getSheetByName(staffName);
    if (!sheet) return;
    var range = sheet.getRange(startRow, C.PLACE, FIXED_DAYS, 1);
    var values = range.getValues();
    var changed = false;
    for (var i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim() === oldName) {
        values[i][0] = newName;
        changed = true;
      }
    }
    if (changed) range.setValues(values);
  });
}

// ── 通常入力：開始日〜終了日の範囲にまとめて書き込み ─────────
// data = { startDay, endDay, place, highway, tollGo, tollBack }
function submitCommute(staffName, data) {
  var sheet = getCommuteSheet(staffName);
  var C = CONFIG.COL;
  var startDay = Number(data.startDay);
  var endDay = Number(data.endDay) || startDay;
  if (endDay < startDay) { var tmp = startDay; startDay = endDay; endDay = tmp; }

  // 場所マスタは範囲全体で1回だけ読む。日ごとに読み直すと
  // 連続勤務をまとめて送った時に往復が日数分だけ増える。
  var placeMap = getPlaceMap(SpreadsheetApp.getActiveSpreadsheet());
  var isHighway = (data.highway === '有');
  var tollGo    = isHighway && data.tollGo   ? Number(data.tollGo)   : '';
  var tollBack  = isHighway && data.tollBack ? Number(data.tollBack) : '';

  for (var day = startDay; day <= endDay; day++) {
    var row = CONFIG.dataStartRow + day - 1;
    sheet.getRange(row, C.PLACE).setValue(data.place);
    sheet.getRange(row, C.HIGHWAY).setValue(isHighway ? '有' : '無');
    // G・H は隣り合っているのでまとめて1回で書く
    sheet.getRange(row, C.TOLL_GO, 1, 2).setValues([[tollGo, tollBack]]);

    updateRowDistance(sheet, staffName, row, placeMap);
  }

  return { updatedDays: endDay - startDay + 1 };
}

// ── イレギュラー入力：空いている枠に書き込み ─────────────────
// data = { day, content, distanceKm }
function submitIrregular(staffName, data) {
  var sheet = getCommuteSheet(staffName);
  var C = CONFIG.COL;
  var month = Number(sheet.getRange(CONFIG.MONTH_ROW, CONFIG.MONTH_COL).getValue());
  var range = getIrregularRowRange();

  var targetRow = null;
  for (var r = range.irregStart; r <= range.irregEnd; r++) {
    var content = sheet.getRange(r, C.PLACE).getValue();
    if (!content) { targetRow = r; break; }
  }
  if (!targetRow) {
    throw new Error('イレギュラー枠が上限（' + CONFIG.IRREGULAR_ROWS + '件）に達しています。確認・修正タブから既存の項目を削除してください。');
  }

  var distanceKm = Number(data.distanceKm) || 0;
  var amount = Math.round(distanceKm * CONFIG.ratePerKm);

  sheet.getRange(targetRow, 1).setValue(new Date(CONFIG.year, month - 1, Number(data.day))).setNumberFormat('m/d');
  sheet.getRange(targetRow, C.PLACE).setValue(data.content);
  sheet.getRange(targetRow, C.TOTAL_AMT).setValue(amount);
  sheet.getRange(targetRow, C.NOTE).setValue(distanceKm ? (distanceKm + 'km') : '');

  return { row: targetRow, amount: amount };
}

// ── その月の入力済み一覧（通常＋イレギュラー） ───────────────
function getMyEntries(staffName) {
  var sheet = getCommuteSheet(staffName);
  var C = CONFIG.COL;
  var month = Number(sheet.getRange(CONFIG.MONTH_ROW, CONFIG.MONTH_COL).getValue());
  var daysInMonth = new Date(CONFIG.year, month, 0).getDate();

  var range    = getIrregularRowRange();
  var firstRow = CONFIG.dataStartRow;
  var rowCount = range.irregEnd - firstRow + 1;

  // A〜K列を1回でまとめて読む。1セルずつ getValue() すると
  // 1行あたり5〜6回、31行で150回以上の往復になり非常に遅い。
  var values = sheet.getRange(firstRow, 1, rowCount, 11).getValues();

  var results = [];

  for (var d = 1; d <= daysInMonth; d++) {
    var v = values[d - 1];
    if (!v[C.PLACE - 1]) continue;
    results.push({
      type: 'normal',
      day: d,
      place: v[C.PLACE - 1],
      highway: v[C.HIGHWAY - 1],
      tollGo: v[C.TOLL_GO - 1],
      tollBack: v[C.TOLL_BACK - 1],
      amount: v[C.TOTAL_AMT - 1],
      row: firstRow + d - 1
    });
  }

  for (var r = range.irregStart; r <= range.irregEnd; r++) {
    var iv = values[r - firstRow];
    if (!iv[C.PLACE - 1]) continue;
    var dateVal = iv[C.DATE - 1];
    results.push({
      type: 'irregular',
      day: (dateVal instanceof Date) ? dateVal.getDate() : null,
      content: iv[C.PLACE - 1],
      amount: iv[C.TOTAL_AMT - 1],
      note: iv[C.NOTE - 1],
      row: r
    });
  }

  results.sort(function(a, b) { return (a.day || 99) - (b.day || 99); });
  return results;
}

// ── 通常入力の1日分を削除 ────────────────────────────────────
function deleteNormalEntry(staffName, day) {
  var sheet = getCommuteSheet(staffName);
  var C = CONFIG.COL;
  var row = CONFIG.dataStartRow + Number(day) - 1;

  sheet.getRange(row, C.PLACE).clearContent();
  sheet.getRange(row, C.DIST_GO).clearContent();
  sheet.getRange(row, C.DIST_BACK).clearContent();
  sheet.getRange(row, C.TOLL_GO).clearContent();
  sheet.getRange(row, C.TOLL_BACK).clearContent();
  sheet.getRange(row, C.HIGHWAY).clearContent();
  clearDistanceNote(sheet, row);
  // DIST_TOTAL(F)・TOTAL_AMT(J)は数式なのでそのまま。C・D・Eが空になれば自動で空欄評価される

  return 'OK';
}

// ── イレギュラー入力の1件を削除 ──────────────────────────────
function deleteIrregularEntry(staffName, row) {
  var sheet = getCommuteSheet(staffName);
  var C = CONFIG.COL;
  row = Number(row);
  var range = getIrregularRowRange();
  if (row < range.irregStart || row > range.irregEnd) {
    throw new Error('削除対象の行が不正です');
  }

  sheet.getRange(row, 1).clearContent();
  sheet.getRange(row, C.PLACE).clearContent();
  sheet.getRange(row, C.TOTAL_AMT).clearContent();
  sheet.getRange(row, C.NOTE).clearContent();

  return 'OK';
}

// ── 対象月の変更（カレンダー行を再生成） ─────────────────────
function resetMonthWeb(staffName, month) {
  var sheet = getCommuteSheet(staffName);
  month = Number(month);
  if (!month || month < 1 || month > 12) {
    throw new Error('月の指定が不正です');
  }
  sheet.getRange(CONFIG.MONTH_ROW, CONFIG.MONTH_COL).setValue(month);
  buildMonthRows(sheet, month);
  setupDropdowns(sheet);
  return { year: CONFIG.year, month: month };
}

// ── Google DriveへPDF保存（指定フォルダへ直接保存） ──────────
function exportCommutePdf(staffName) {
  if (!PDF_ROOT_FOLDER_ID) {
    throw new Error('PDF保存先フォルダが未設定です。PDF_ROOT_FOLDER_IDを設定してください。');
  }
  var sheet = getCommuteSheet(staffName);
  var ss = sheet.getParent();
  var month = Number(sheet.getRange(CONFIG.MONTH_ROW, CONFIG.MONTH_COL).getValue());

  // フォルダIDの誤りとアクセス権不足を区別できるようにメッセージを付ける
  var targetFolder;
  try {
    targetFolder = DriveApp.getFolderById(PDF_ROOT_FOLDER_ID);
  } catch (err) {
    throw new Error('PDF保存先フォルダにアクセスできません（ID: ' + PDF_ROOT_FOLDER_ID + '）。'
      + 'スクリプトの実行アカウントにこのフォルダの編集権限があるか確認してください。');
  }

  var fileName = CONFIG.year + '年' + month + '月_' + staffName + '_稼働費.pdf';
  var pdfBlob = exportCommuteSheetAsPdf(ss, sheet).setName(fileName);

  var existingFiles = targetFolder.getFilesByName(fileName);
  while (existingFiles.hasNext()) {
    existingFiles.next().setTrashed(true);
  }
  targetFolder.createFile(pdfBlob);

  return '保存完了';
}

function exportCommuteSheetAsPdf(ss, sheet) {
  var P = CONFIG.PDF;

  // 出力範囲を合計行・K列までに限定する。範囲を指定せずシート全体を対象にすると、
  // 右側・下側の未使用セルまで含めて1ページに押し込もうとして文字が極端に小さくなる。
  // r1/c1 は0始まりの開始位置、r2/c2 は終了行・終了列（1始まりの番号）。
  var lastRow = getIrregularRowRange().irregEnd + 1;  // 合計行
  var lastCol = 11;                                   // K列

  var params = {
    exportFormat: 'pdf',
    format: 'pdf',
    gid: sheet.getSheetId(),
    size: P.size,
    portrait: P.portrait ? 'true' : 'false',
    scale: P.scale,          // 4 = 1ページに収める
    r1: 0,
    c1: 0,
    r2: lastRow,
    c2: lastCol,
    top_margin: P.margin,
    bottom_margin: P.margin,
    left_margin: P.margin,
    right_margin: P.margin,
    horizontal_alignment: 'CENTER',
    vertical_alignment: 'TOP',
    sheetnames: 'false',
    printtitle: 'false',
    pagenumbers: 'false',
    gridlines: 'false',
    fzr: 'false'
  };

  // ss.getUrl() は末尾が "/edit" とは限らず（#gid=... が付く等）、
  // replace(/edit$/, "") だと ".../editexport" のような壊れたURLになり、
  // PDFではなくHTMLのエラーページを保存してしまう。IDから組み立てる方が確実。
  var query = Object.keys(params).map(function(k) {
    return k + '=' + encodeURIComponent(params[k]);
  }).join('&');
  var exportUrl = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' + query;

  var response = UrlFetchApp.fetch(exportUrl, {
    headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    throw new Error('PDFの書き出しに失敗しました（HTTP ' + code + '）。'
      + 'スクリプトの実行アカウントにこのスプレッドシートの閲覧権限があるか確認してください。');
  }

  var blob = response.getBlob();
  // 権限が無い場合はPDFではなくHTMLのログイン画面が200で返ることがある。
  // そのまま保存すると中身がHTMLの「.pdf」ができてしまうので弾く。
  if (String(blob.getContentType() || '').indexOf('pdf') === -1) {
    throw new Error('PDFではなくHTMLが返されました。Webアプリの「実行するユーザー」設定と、'
      + 'そのアカウントのスプレッドシート閲覧権限を確認してください。');
  }
  return blob;
}
