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
//     　権限エラーで「取得エラー」になったりします。
//  5. 以降は月を変えたら「🗓️ 月を更新」を実行
//
//  【日々の使い方】
//  ・C列（稼働場所）をプルダウンで選択
//    → I列（高速 有/無）をプルダウンで選択
//    → 距離が自動で入ります
//  ・G列・H列は高速代（円）を手入力
// ============================================================

const CONFIG = {
  year: 2026,
  ratePerKm: 15,       // 円/km
  dataStartRow: 6,     // データ開始行
  MONTH_ROW: 2,
  MONTH_COL: 7,        // G2: 月入力セル

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

  EDIT_TRIGGER_HANDLER: 'handleEdit', // インストール型onEditトリガーのハンドラ名
  ROUTE_CACHE_SEC: 21600,             // 同一経路の距離キャッシュ保持時間（秒・最大6時間）
};

const WEEKDAY_JA = ['日','月','火','水','木','金','土'];

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

  // 全スタッフシートにプルダウンを設定
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

// ── 月更新 ───────────────────────────────────────────────────
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

  // 既存データをクリア（最大31日分 + 旧合計行を含めて余裕を持ってクリア）
  // 31日の月から30日の月に変えた際、旧合計行（startRow+31）が残らないよう
  // 32行分（データ31日 + 合計行1）をクリア対象にする
  sheet.getRange(startRow, 1, 32, 11).clearContent().clearFormat();

  // ベースのボーダー・書式を再設定
  const dataRange = sheet.getRange(startRow, 1, daysInMonth, 11);
  dataRange.setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  for (let day = 1; day <= daysInMonth; day++) {
    const row = startRow + day - 1;
    // タイムゾーンズレ完全回避: DATE関数で日付をセット
    sheet.getRange(row, C.DATE)
      .setFormula(`=DATE(${year},${month},${day})`)
      .setNumberFormat('m/d')
      .setHorizontalAlignment('center');
    const wd = new Date(year, month - 1, day).getDay();  // 0=日, 6=土
    const isWe = wd === 0 || wd === 6;

    // B: 曜日
    const wdColor = wd === 0 ? '#CC0000' : (wd === 6 ? '#0070C0' : '#333333');
    sheet.getRange(row, C.WEEKDAY)
      .setValue(WEEKDAY_JA[wd])
      .setFontColor(wdColor)
      .setHorizontalAlignment('center');

    // 土日の背景色
    if (isWe) {
      sheet.getRange(row, 1, 1, 11).setBackground('#F0F0F0');
    }

    // G・H列（高速代）の書式
    sheet.getRange(row, C.TOLL_GO).setNumberFormat('#,##0').setBackground(isWe ? '#F0F0F0' : '#EBF5FB');
    sheet.getRange(row, C.TOLL_BACK).setNumberFormat('#,##0').setBackground(isWe ? '#F0F0F0' : '#EBF5FB');
  }

  // 合計行（データ日数の直後1行のみ。前月分の余り行は上のclearContentで消去済み）
  const totalRow = startRow + daysInMonth;
  const lastRow  = startRow + daysInMonth - 1;
  sheet.getRange(totalRow, 1, 1, 11).setBorder(true, true, true, true, true, true, '#999999', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(totalRow, 1).setValue('合　　計').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#FFF2CC');
  sheet.getRange(totalRow, 1, 1, 11).setBackground('#FFF2CC');
  sheet.getRange(totalRow, C.DIST_TOTAL).setFormula(`=SUM(F${startRow}:F${lastRow})`).setNumberFormat('#,##0.000"km"').setFontWeight('bold');
  sheet.getRange(totalRow, C.TOLL_GO).setFormula(`=SUM(G${startRow}:G${lastRow})`).setNumberFormat('#,##0"円"').setFontWeight('bold');
  sheet.getRange(totalRow, C.TOLL_BACK).setFormula(`=SUM(H${startRow}:H${lastRow})`).setNumberFormat('#,##0"円"').setFontWeight('bold');
  sheet.getRange(totalRow, C.TOTAL_AMT).setFormula(`=SUM(J${startRow}:J${lastRow})`).setNumberFormat('#,##0"円"').setFontWeight('bold').setFontSize(13);

  // 稼働日数（C列の非空白カウント）
  sheet.getRange(2, 9).setFormula(`=COUNTA(C${startRow}:C${lastRow})`);
}

// ── プルダウン設定 ───────────────────────────────────────────
function setupDropdowns(sheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const placeSheet = ss.getSheetByName(CONFIG.PLACE_SHEET);
  if (!placeSheet) return;

  const month = Number(sheet.getRange(CONFIG.MONTH_ROW, CONFIG.MONTH_COL).getValue());
  const daysInMonth = new Date(CONFIG.year, month, 0).getDate();
  const startRow = CONFIG.dataStartRow;
  const C = CONFIG.COL;

  // C列: 稼働場所プルダウン（稼働場所シートのA列全件）
  const lastPlaceRow = placeSheet.getLastRow();
  const placeRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(placeSheet.getRange(`A2:A${lastPlaceRow}`), true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(startRow, C.PLACE, daysInMonth, 1).setDataValidation(placeRule);

  // I列: 高速 有/無 プルダウン
  const hwRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['無', '有'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(startRow, C.HIGHWAY, daysInMonth, 1).setDataValidation(hwRule);
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
  const month = Number(sheet.getRange(CONFIG.MONTH_ROW, CONFIG.MONTH_COL).getValue());
  if (!month) return;
  const daysInMonth  = new Date(CONFIG.year, month, 0).getDate();
  const startRow     = CONFIG.dataStartRow;
  const lastDataRow  = startRow + daysInMonth - 1;

  // C列（場所）またはI列（高速有無）の変更
  if ((editedCol === C.PLACE || editedCol === C.HIGHWAY)
      && editedRow >= startRow && editedRow <= lastDataRow) {
    updateRowDistance(sheet, staffName, editedRow);
  }

  // G・H列（高速代）の変更 → 合計金額を再計算
  if ((editedCol === C.TOLL_GO || editedCol === C.TOLL_BACK)
      && editedRow >= startRow && editedRow <= lastDataRow) {
    setTotalFormula(sheet, editedRow);
  }
}

// ── 1行分の距離を取得・セット ────────────────────────────────
function updateRowDistance(sheet, staffName, row) {
  const C          = CONFIG.COL;
  const placeShort = sheet.getRange(row, C.PLACE).getValue();
  const highway    = sheet.getRange(row, C.HIGHWAY).getValue();  // '有' or '無'

  // 場所が空なら距離もクリア
  if (!placeShort) {
    sheet.getRange(row, C.DIST_GO, 1, 4).clearContent();
    sheet.getRange(row, C.NOTE).clearContent();
    return;
  }

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const placeMap = getPlaceMap(ss);

  // 自宅住所（稼働場所リストの「○○自宅」を参照）
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

  // 高速有無でルートタイプを切替
  // '有' → highway（高速利用）, '無' または未選択 → drive（一般道）
  const routeType = (highway === '有') ? 'highway' : 'drive';

  try {
    const distGo   = getRouteDistanceKm(homeAddr, destAddr, routeType);
    const distBack = getRouteDistanceKm(destAddr, homeAddr, routeType);

    sheet.getRange(row, C.DIST_GO).setValue(distGo).setNumberFormat('#,##0.000');
    sheet.getRange(row, C.DIST_BACK).setValue(distBack).setNumberFormat('#,##0.000');
    sheet.getRange(row, C.DIST_TOTAL)
      .setFormula(`=IF(AND(ISNUMBER(D${row}),ISNUMBER(E${row})),D${row}+E${row},"")`)
      .setNumberFormat('#,##0.000');
    sheet.getRange(row, C.NOTE).clearContent();
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

  let updated = 0;
  for (let day = 1; day <= daysInMonth; day++) {
    const row        = startRow + day - 1;
    const placeShort = sheet.getRange(row, C.PLACE).getValue();
    if (!placeShort) continue;

    updateRowDistance(sheet, name, row);
    updated++;

    // API レート制限対策（キャッシュヒット時はmapQuery自体が呼ばれないので実質スキップされる）
    if (updated % 5 === 0) Utilities.sleep(1000);
  }

  ui.alert(`✅ ${updated}件の距離を更新しました。`);
}

// ── 場所マスタ読み込み（短縮名 → 住所） ─────────────────────
// 稼働場所シート: A列=短縮名, B列=住所
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
      // 高速道路を使う（制限なし）
      finder.setMode(Maps.DirectionFinder.Mode.DRIVING);
      break;
    case 'drive':
      // 一般道のみ（有料道路・高速を回避）
      finder.setMode(Maps.DirectionFinder.Mode.DRIVING)
            .setAvoid(Maps.DirectionFinder.Avoid.TOLLS);
      break;
    case 'toll':
      // 高速は使わない（toll道路回避）
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

  const route = finder.getDirections().routes[0];
  const leg   = route.legs[0];

  if (result === 'distance') return leg.distance.value / 1000;   // km
  if (result === 'duration') return leg.duration.value / 60;     // 分
  return leg;
}
