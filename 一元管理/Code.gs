/**
 * 経堂リニューアル　アイデアプール — 統合 GAS
 *
 * ★ Apps Script にはこの Code.gs だけ貼れば OK（他ファイル不要）
 * Git 開発用の分割版: _modules/ フォルダ
 *
 * https://docs.google.com/spreadsheets/d/1rayTPZSL9X0UsI9Gv2SJthlBviT2WUvvHCJvC6zroIw/edit
 * ID: 1rayTPZSL9X0UsI9Gv2SJthlBviT2WUvvHCJvC6zroIw
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("共同データ")
    .addItem("リフォーマーデータを更新", "pullReformerAllData")
    .addItem("（確認）deta列を表示", "inspectReformerDetaHeaders")
    .addSeparator()
    .addItem("利用状況をリアルタイム連携", "setupUsageStatusLiveLink")
    .addItem("利用状況: FW ⇔ TM 切替", "toggleUsageStatusMode")
    .addToUi();
}

// =============================================================================
// リフォーマー連携
// =============================================================================
const REFORMER_SYNC = {
  REFORMER_SPREADSHEET_ID: "1Qc1l2pqkoMAgGae_7epfCcvj62a3O_BPbL-ezWDR5pc",
  SOURCE_SHEET_NAME: "会員数",
  DEST_SHEET_NAME: "リフォーマー",
  DETA_SHEET_NAME: "deta"
};

/** メニュー用: 会員数コピー ＋（将来）ダッシュボード追記 */
function pullReformerAllData() {
  const ui = SpreadsheetApp.getUi();
  try {
    const copiedRows = pullReformerMemberCount_();
    const dstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REFORMER_SYNC.DEST_SHEET_NAME);
    integrateDashboardRows789_(dstSheet);
    ui.alert(
      "更新完了",
      "会員数を反映しました。\n" +
        "・8行目から【2025年度】表（上部サマリー直後）\n" +
        "・⑥の下に ⑦⑧⑨ を追加\n\n" +
        "⑧⑨は数式です。ずらすときは ⑦の数値、または\n" +
        "非表示行の枠数・キャンセル数（各月 B〜M）を編集してください。",
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert("更新エラー", String(e.message || e), ui.ButtonSet.OK);
    throw e;
  }
}

/** @deprecated メニュー名変更前の互換 */
function pullReformerMemberCount() {
  return pullReformerAllData();
}

/**
 * 会員数シートをリフォーマーへコピー（値＋書式）
 * @return {number} コピーした行数
 */
function pullReformerMemberCount_() {
  const cfg = REFORMER_SYNC;
  const dstSs = SpreadsheetApp.getActiveSpreadsheet();

  const srcSs = SpreadsheetApp.openById(cfg.REFORMER_SPREADSHEET_ID);
  const srcSheet = srcSs.getSheetByName(cfg.SOURCE_SHEET_NAME);
  if (!srcSheet) {
    throw new Error("チェック表に「" + cfg.SOURCE_SHEET_NAME + "」がありません。");
  }

  let dstSheet = dstSs.getSheetByName(cfg.DEST_SHEET_NAME);
  if (!dstSheet) {
    dstSheet = dstSs.insertSheet(cfg.DEST_SHEET_NAME);
  }

  const lastRow = srcSheet.getLastRow();
  const lastCol = srcSheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    dstSheet.clear();
    return 0;
  }

  const srcRange = srcSheet.getRange(1, 1, lastRow, lastCol);
  dstSheet.clear();
  const dstRange = dstSheet.getRange(1, 1, lastRow, lastCol);
  // 別ブック間は copyTo 不可 → 値と書式を個別にコピー
  copyRangeAcrossWorkbooks_(srcRange, dstRange);

  for (let c = 1; c <= lastCol; c++) {
    dstSheet.setColumnWidth(c, srcSheet.getColumnWidth(c));
  }

  return lastRow;
}

/**
 * 別スプレッドシート間の範囲コピー（copyTo は同一ブック専用のため）
 */
function copyRangeAcrossWorkbooks_(srcRange, dstRange) {
  dstRange.setValues(srcRange.getValues());
  dstRange.setNumberFormats(srcRange.getNumberFormats());
  dstRange.setBackgrounds(srcRange.getBackgrounds());
  dstRange.setFontWeights(srcRange.getFontWeights());
  dstRange.setFontColors(srcRange.getFontColors());
  dstRange.setFontFamilies(srcRange.getFontFamilies());
  dstRange.setFontSizes(srcRange.getFontSizes());
  dstRange.setHorizontalAlignments(srcRange.getHorizontalAlignments());
  dstRange.setVerticalAlignments(srcRange.getVerticalAlignments());
  dstRange.setWraps(srcRange.getWraps());
}
// =============================================================================
// リフォーマー deta 集計（⑦⑧⑨）
// =============================================================================
const DASHBOARD_AGG = {
  REFORMER_SPREADSHEET_ID: "1Qc1l2pqkoMAgGae_7epfCcvj62a3O_BPbL-ezWDR5pc",
  DETA_SHEET_NAME: "deta",
  MAX_CAPACITY: 4,
  FISCAL_MONTHS: [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3],
  MONTH_COL_START: 2,
  MONTH_COL_END: 13,
  YEAR_TOTAL_COL: 14,
  FISCAL_SECTION_START_ROW: 8,
  DATE_HEADERS: ["予約日", "予約日時", "レッスン日", "日付", "日時"],
  TIME_HEADERS: ["時間", "開始", "開始時刻", "レッスン時間"],
  INSTRUCTOR_HEADERS: ["インストラクター", "担当", "スタッフ", "講師"],
  STATUS_HEADERS: ["区分", "ステータス", "状態", "種別", "ラベル", "Gmailラベル"]
};

function aggregateReformerMonthAll_(calendarYear, calendarMonth) {
  const parsed = loadDetaRows_();
  if (!parsed.rows.length) {
    return { reservations: 0, cancels: 0, slots: 0 };
  }

  const slotKeys = {};
  let reservations = 0;
  let cancels = 0;

  for (let i = 0; i < parsed.rows.length; i++) {
    const row = parsed.rows[i];
    const d = parseDashboardDate_(row[parsed.colDate], row[parsed.colTime]);
    if (!d || d.getFullYear() !== calendarYear || d.getMonth() + 1 !== calendarMonth) continue;

    const instructor =
      parsed.colInstructor >= 0 ? String(row[parsed.colInstructor] || "").trim() : "";
    slotKeys[dashboardSlotKey_(d, row[parsed.colTime], instructor)] = true;
    if (isCancelRow_(row, parsed.colStatus)) cancels++;
    else reservations++;
  }

  return {
    reservations: reservations,
    cancels: cancels,
    slots: Object.keys(slotKeys).length
  };
}

function aggregateFiscalYearRaw_(fiscalYear) {
  const months = DASHBOARD_AGG.FISCAL_MONTHS;
  const res = [];
  const can = [];
  const slots = [];
  for (let i = 0; i < months.length; i++) {
    const m = months[i];
    const calYear = m >= 4 ? fiscalYear : fiscalYear + 1;
    const s = aggregateReformerMonthAll_(calYear, m);
    res.push(s.reservations);
    can.push(s.cancels);
    slots.push(s.slots);
  }
  return { reservations: res, cancels: can, slots: slots };
}

/**
 * 会員数コピー後: レイアウト整理 → ⑥の下に ⑦⑧⑨ を統合
 */
function integrateDashboardRows789_(dstSheet) {
  removeOldSeparateDashboardBlock_(dstSheet);
  compactReformerSheetLayout_(dstSheet);

  integrateOneFiscalYear_(dstSheet, 2025, aggregateFiscalYearRaw_(2025));
  integrateOneFiscalYear_(dstSheet, 2026, aggregateFiscalYearRaw_(2026));
}

/**
 * 上部サマリー(1〜7行)の直後に【2025年度】が来るよう詰める。説明文・図形も削除。
 */
function compactReformerSheetLayout_(sheet) {
  removeReformerDashboardFluff_(sheet);

  const targetRow = DASHBOARD_AGG.FISCAL_SECTION_START_ROW;
  const headerRow = findFiscalYearHeaderRow_(sheet, 2025);
  if (headerRow < 0) return;

  if (headerRow > targetRow) {
    sheet.deleteRows(targetRow, headerRow - targetRow);
  }
}

function removeReformerDashboardFluff_(sheet) {
  const drawings = sheet.getDrawings();
  for (let i = drawings.length - 1; i >= 0; i--) {
    try {
      drawings[i].remove();
    } catch (e) {
      /* 図形削除失敗は無視 */
    }
  }

  const lastRow = Math.min(sheet.getLastRow(), 40);
  for (let r = 1; r <= lastRow; r++) {
    const v = String(sheet.getRange(r, 1).getValue() || "");
    if (
      v.indexOf("Pilates Reformer") >= 0 ||
      v.indexOf("ダッシュボードの計算式") >= 0 ||
      v.indexOf("▼") === 0
    ) {
      sheet.getRange(r, 1, 1, sheet.getLastColumn()).clearContent();
    }
  }
}

function findFiscalYearHeaderRow_(sheet, fiscalYear) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return -1;
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    const t = String(colA[i][0] || "");
    if (t.indexOf(String(fiscalYear)) >= 0 && t.indexOf("年度") >= 0) return i + 1;
  }
  return -1;
}

function integrateOneFiscalYear_(sheet, fiscalYear, raw) {
  const row6 = findRow6InFiscalSection_(sheet, fiscalYear);
  if (row6 < 0) {
    throw new Error("【" + fiscalYear + "年度】付近の ⑥ 行が見つかりません。");
  }

  ensureMetricRowsAfterRow6_(sheet, row6);

  const row7 = row6 + 1;
  const row8 = row6 + 2;
  const row9 = row6 + 3;
  const rowSlots = row6 + 4;
  const rowCancels = row6 + 5;
  const c0 = DASHBOARD_AGG.MONTH_COL_START;
  const c1 = DASHBOARD_AGG.MONTH_COL_END;
  const nCol = DASHBOARD_AGG.YEAR_TOTAL_COL;
  const cap = DASHBOARD_AGG.MAX_CAPACITY;

  sheet.getRange(row7, 1).setValue("⑦月間予約数（有効）");
  sheet.getRange(row8, 1).setValue("⑧平均稼働率");
  sheet.getRange(row9, 1).setValue("⑨キャンセル率");
  sheet.getRange(rowSlots, 1).setValue("");
  sheet.getRange(rowCancels, 1).setValue("");

  rangeRow_(sheet, row7, c0, c1).setValues([raw.reservations]);
  rangeRow_(sheet, row7, c0, c1).setNumberFormat("0");
  rangeRow_(sheet, rowSlots, c0, c1).setValues([raw.slots]);
  rangeRow_(sheet, rowSlots, c0, c1).setNumberFormat("0");
  rangeRow_(sheet, rowCancels, c0, c1).setValues([raw.cancels]);
  rangeRow_(sheet, rowCancels, c0, c1).setNumberFormat("0");

  for (let c = c0; c <= c1; c++) {
    const col = columnToLetter_(c);
    sheet.getRange(row8, c).setFormula(
      "=IF(" +
        col +
        rowSlots +
        "*" +
        cap +
        ">0," +
        col +
        row7 +
        "/(" +
        col +
        rowSlots +
        "*" +
        cap +
        "),0)"
    );
    sheet.getRange(row9, c).setFormula(
      "=IF(" +
        col +
        row7 +
        "+" +
        col +
        rowCancels +
        ">0," +
        col +
        rowCancels +
        "/(" +
        col +
        row7 +
        "+" +
        col +
        rowCancels +
        "),0)"
    );
  }

  const bL = columnToLetter_(c0);
  const mL = columnToLetter_(c1);
  sheet.getRange(row7, nCol).setFormula("=SUM(" + bL + row7 + ":" + mL + row7 + ")");
  sheet.getRange(row7, nCol).setNumberFormat("#,##0");
  sheet.getRange(row8, nCol).setFormula(
    "=IF(SUM(" +
      bL +
      rowSlots +
      ":" +
      mL +
      rowSlots +
      ")*" +
      cap +
      ">0,SUM(" +
      bL +
      row7 +
      ":" +
      mL +
      row7 +
      ")/(SUM(" +
      bL +
      rowSlots +
      ":" +
      mL +
      rowSlots +
      ")*" +
      cap +
      "),0)"
  );
  sheet.getRange(row9, nCol).setFormula(
    "=IF(SUM(" +
      bL +
      row7 +
      ":" +
      mL +
      row7 +
      ")+SUM(" +
      bL +
      rowCancels +
      ":" +
      mL +
      rowCancels +
      ")>0,SUM(" +
      bL +
      rowCancels +
      ":" +
      mL +
      rowCancels +
      ")/(SUM(" +
      bL +
      row7 +
      ":" +
      mL +
      row7 +
      ")+SUM(" +
      bL +
      rowCancels +
      ":" +
      mL +
      rowCancels +
      ")),0)"
  );

  rangeRow_(sheet, row8, c0, c1).setNumberFormat("0.0%");
  sheet.getRange(row8, nCol).setNumberFormat("0.0%");
  rangeRow_(sheet, row9, c0, c1).setNumberFormat("0.0%");
  sheet.getRange(row9, nCol).setNumberFormat("0.0%");

  const bg = fiscalYear === 2025 ? "#fdf2f8" : "#fff7ed";
  sheet.getRange(row7, 1, 3, nCol).setBackground(bg);

  sheet.hideRows(rowSlots, 2);
}

/** ⑥の直下: ⑦⑧⑨ ＋ 枠数/キャンセル非表示行（計5行） */
function ensureMetricRowsAfterRow6_(sheet, row6) {
  const next = String(sheet.getRange(row6 + 1, 1).getValue() || "");
  if (next.indexOf("⑦") < 0) {
    sheet.insertRowsAfter(row6, 5);
    return;
  }

  const rowAfter9 = String(sheet.getRange(row6 + 4, 1).getValue() || "");
  if (rowAfter9.indexOf("枠数") >= 0) return;

  const maybeNextYear = rowAfter9.indexOf("年度") >= 0;
  if (maybeNextYear) {
    sheet.insertRowsBefore(row6 + 4, 2);
  } else {
    sheet.insertRowsAfter(row6 + 3, 2);
  }
}

function findRow6InFiscalSection_(sheet, fiscalYear) {
  const lastRow = sheet.getLastRow();
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  let sectionStart = -1;
  let sectionEnd = lastRow + 1;

  for (let i = 0; i < colA.length; i++) {
    const t = String(colA[i][0] || "");
    if (sectionStart < 0 && t.indexOf(String(fiscalYear)) >= 0 && t.indexOf("年度") >= 0) {
      sectionStart = i + 1;
      continue;
    }
    if (sectionStart > 0 && i + 1 > sectionStart && t.indexOf("年度") >= 0 && t.indexOf(String(fiscalYear)) < 0) {
      sectionEnd = i + 1;
      break;
    }
  }
  if (sectionStart < 0) return -1;

  for (let r = sectionStart; r < sectionEnd; r++) {
    const t = String(colA[r - 1][0] || "");
    if (t.indexOf("⑥") >= 0) return r;
  }
  return -1;
}

function removeOldSeparateDashboardBlock_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = colA.length - 1; i >= 0; i--) {
    const t = String(colA[i][0] || "");
    if (
      t.indexOf("全インストラクター合計") >= 0 ||
      t.indexOf("【2025】月間予約数") >= 0 ||
      t.indexOf("【2026】月間予約数") >= 0
    ) {
      sheet.deleteRow(i + 1);
    }
  }
}

function columnToLetter_(col) {
  let n = col;
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** 1行分の範囲（A1表記で getRange の行数/終端行の取り違えを防ぐ） */
function rangeRow_(sheet, row, colStart, colEnd) {
  return sheet.getRange(columnToLetter_(colStart) + row + ":" + columnToLetter_(colEnd) + row);
}

function loadDetaRows_() {
  const ss = SpreadsheetApp.openById(DASHBOARD_AGG.REFORMER_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(DASHBOARD_AGG.DETA_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) {
    return { rows: [], colDate: 0, colTime: -1, colStatus: -1, colInstructor: -1, headers: [] };
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(function (h) {
    return String(h || "").trim();
  });
  const colDate = findHeaderIndex_(headers, DASHBOARD_AGG.DATE_HEADERS);
  const colTime = findHeaderIndex_(headers, DASHBOARD_AGG.TIME_HEADERS);
  const colStatus = findHeaderIndex_(headers, DASHBOARD_AGG.STATUS_HEADERS);
  const colInstructor = findHeaderIndex_(headers, DASHBOARD_AGG.INSTRUCTOR_HEADERS);

  if (colDate < 0) {
    throw new Error("deta シートに日付列が見つかりません。1行目: " + headers.join(" / "));
  }

  const rows = [];
  for (let r = 1; r < data.length; r++) rows.push(data[r]);
  return {
    rows: rows,
    colDate: colDate,
    colTime: colTime,
    colStatus: colStatus,
    colInstructor: colInstructor,
    headers: headers
  };
}

function findHeaderIndex_(headers, candidates) {
  for (let c = 0; c < candidates.length; c++) {
    for (let h = 0; h < headers.length; h++) {
      if (headers[h].indexOf(candidates[c]) >= 0) return h;
    }
  }
  return -1;
}

function parseDashboardDate_(dateCell, timeCell) {
  if (dateCell instanceof Date && !isNaN(dateCell.getTime())) {
    const d = new Date(dateCell.getTime());
    const t = parseTimePart_(timeCell);
    if (t) d.setHours(t.h, t.m, 0, 0);
    return d;
  }
  const s = String(dateCell || "").trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  const d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

function parseTimePart_(timeCell) {
  if (timeCell instanceof Date) return { h: timeCell.getHours(), m: timeCell.getMinutes() };
  const m = String(timeCell || "").trim().match(/(\d{1,2}):(\d{2})/);
  if (m) return { h: parseInt(m[1], 10), m: parseInt(m[2], 10) };
  return null;
}

function dashboardSlotKey_(dateObj, timeCell, instructor) {
  const t = parseTimePart_(timeCell);
  const th = t ? t.h : dateObj.getHours();
  const tm = t ? t.m : dateObj.getMinutes();
  return (
    dateObj.getFullYear() +
    "-" +
    (dateObj.getMonth() + 1) +
    "-" +
    dateObj.getDate() +
    " " +
    th +
    ":" +
    tm +
    "|" +
    (instructor || "")
  );
}

function isCancelRow_(row, colStatus) {
  if (colStatus < 0) return false;
  const s = String(row[colStatus] || "").toLowerCase();
  return s.indexOf("キャンセル") >= 0 || s.indexOf("cancel") >= 0;
}

/** @deprecated 旧: 下部に別表を作る */
function appendReformerDashboardFromDeta_(dstSheet, afterRow) {
  integrateDashboardRows789_(dstSheet);
}

function inspectReformerDetaHeaders() {
  const parsed = loadDetaRows_();
  SpreadsheetApp.getUi().alert(
    "deta 列確認",
    "ヘッダー: " + parsed.headers.join(" | ") + "\n\n" +
      "日付列: " + (parsed.colDate + 1) + "列目\n" +
      "時間列: " + (parsed.colTime >= 0 ? parsed.colTime + 1 : "なし") + "列目\n" +
      "区分列: " + (parsed.colStatus >= 0 ? parsed.colStatus + 1 : "なし") + "列目\n" +
      "担当列: " + (parsed.colInstructor >= 0 ? parsed.colInstructor + 1 : "なし") + "列目\n\n" +
      "⑦=件数（N列は合計）。⑧⑨=率（N列は年間の加重計算）。\n" +
      "枠数/キャンセルは ⑨ の下2行（非表示）の B〜M を編集すると ⑧⑨ が変わります。",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
// =============================================================================
// 利用状況（FW / TM・IMPORTRANGE）
// =============================================================================
const USAGE_STATUS_SYNC = {
  SOURCE_SPREADSHEET_ID: "1qWOf5apYTULT3YWw8oF2ZQ8p73jBM4-7Y7NY8adcsXg",
  DEST_SHEET_NAME: "利用状況",
  LEGACY_SHEET_NAME: "ジム",
  SWITCH_CELL: "B2",
  PROP_DIMS: "usage_status_dims_json",
  PROP_LAYOUT: "usage_status_layout_json",
  MODES: [
    {
      key: "FW",
      switchLabel: "FW フリーウェイト",
      label: "FW フリーウェイト",
      maxCapacity: 18,
      dashSheet: "FW_分析ダッシュボード",
      dataSheet: "FW_データ入力"
    },
    {
      key: "TM",
      switchLabel: "TM トレッドミル",
      label: "TM トレッドミル",
      maxCapacity: 12,
      dashSheet: "TM_分析ダッシュボード",
      dataSheet: "TM_データ入力"
    }
  ],
  DASH_IMPORT_COLS: 11,
  DASH_HEADER_ROWS: 3,
  COL_SUM_AVG: 12,
  COL_UTIL_RATE: 13
};

/** 初回: シート骨組み + デフォルト FW 表示 */
function setupUsageStatusLiveLink() {
  const ui = SpreadsheetApp.getUi();
  const cfg = USAGE_STATUS_SYNC;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = SpreadsheetApp.openById(cfg.SOURCE_SPREADSHEET_ID);
  const sheet = ensureUsageStatusSheet_(ss, cfg);

  cacheUsageSourceDimensions_(src, cfg);
  const layout = getUsageLayout_(cfg);

  sheet.clear();
  buildUsageStatusShell_(sheet, cfg, layout);
  sheet.getRange(cfg.SWITCH_CELL).setValue(cfg.MODES[0].switchLabel);
  applyUsageStatusMode_(sheet, cfg.MODES[0].key, false);

  ui.alert(
    "利用状況シートを設定しました",
    "・上段 = 分析ダッシュボード\n" +
      "・L列=総合平均 M列=利用率（数式・手修正可）\n\n" +
      "B2 のプルダウン、または\n" +
      "共同データ → 「利用状況: FW ⇔ TM 切替」\n" +
      "でフリーウェイト / トレッドミルを切り替えます。\n\n" +
      "※ 初回 #REF! → セルを開き「アクセスを許可」",
    ui.ButtonSet.OK
  );
}

/** メニュー1つで FW ⇔ TM 切替 */
function toggleUsageStatusMode() {
  const cfg = USAGE_STATUS_SYNC;
  const sheet = getUsageStatusSheet_(cfg);
  const current = parseUsageModeKey_(sheet.getRange(cfg.SWITCH_CELL).getValue());
  const nextKey = current === "TM" ? "FW" : "TM";
  const nextMode = findUsageMode_(cfg, nextKey);
  sheet.getRange(cfg.SWITCH_CELL).setValue(nextMode.switchLabel);
  applyUsageStatusMode_(sheet, nextKey, false);
}

/** B2 変更時（プルダウン） */
function onEdit(e) {
  if (!e || !e.range) return;
  const cfg = USAGE_STATUS_SYNC;
  if (e.source.getActiveSheet().getName() !== cfg.DEST_SHEET_NAME) return;
  if (e.range.getA1Notation() !== cfg.SWITCH_CELL) return;
  applyUsageStatusModeFromCell_(e.source);
}

function applyUsageStatusModeFromCell_(ss) {
  const cfg = USAGE_STATUS_SYNC;
  const sheet = ss.getSheetByName(cfg.DEST_SHEET_NAME);
  if (!sheet) return;
  const modeKey = parseUsageModeKey_(sheet.getRange(cfg.SWITCH_CELL).getValue());
  applyUsageStatusMode_(sheet, modeKey, false);
}

function applyUsageStatusMode_(sheet, modeKey, skipSwitchCell) {
  const cfg = USAGE_STATUS_SYNC;
  const mode = findUsageMode_(cfg, modeKey);
  const dims = loadUsageDimensions_(cfg)[modeKey];
  if (!dims) {
    throw new Error("表示モード「" + modeKey + "」のサイズ情報がありません。メニューから再セットアップしてください。");
  }
  const layout = getUsageLayout_(cfg);

  if (!skipSwitchCell) {
    sheet.getRange(cfg.SWITCH_CELL).setValue(mode.switchLabel);
  }

  sheet
    .getRange(layout.rowDashHdr, 1, 1, 13)
    .merge()
    .setValue("📊 " + mode.label + " — 分析ダッシュボード")
    .setFontWeight("bold")
    .setBackground("#e8f0fe")
    .setFontColor("#1a73e8")
    .setHorizontalAlignment("left");

  sheet
    .getRange(layout.rowDataHdr, 1, 1, 13)
    .merge()
    .setValue("📋 " + mode.label + " — データ入力（MAX人数）")
    .setFontWeight("bold")
    .setBackground("#e2e3e5")
    .setHorizontalAlignment("left");

  sheet
    .getRange(layout.rowDash, 1)
    .setFormula(
      buildImportRangeFormula_(
        cfg.SOURCE_SPREADSHEET_ID,
        mode.dashSheet,
        1,
        1,
        dims.dashRows,
        cfg.DASH_IMPORT_COLS
      )
    );

  applyDashboardFormulaColumns_(sheet, layout, mode, dims.dashRows);
  applyUsageStatusVisuals_(sheet, layout, mode, dims);

  sheet
    .getRange(layout.rowData, 1)
    .setFormula(
      buildImportRangeFormula_(
        cfg.SOURCE_SPREADSHEET_ID,
        mode.dataSheet,
        1,
        1,
        dims.dataRows,
        dims.dataCols
      )
    );
}

/** ダッシュボード L=総合平均 M=利用率（A〜K は IMPORTRANGE） */
function applyDashboardFormulaColumns_(sheet, layout, mode, dashRows) {
  const cfg = USAGE_STATUS_SYNC;
  const cap = mode.maxCapacity;
  const hdrRow = layout.rowDash + cfg.DASH_HEADER_ROWS - 1;
  const firstData = layout.rowDash + cfg.DASH_HEADER_ROWS;
  const lastData = layout.rowDash + dashRows - 1;
  const cL = cfg.COL_SUM_AVG;
  const cM = cfg.COL_UTIL_RATE;

  sheet
    .getRange(hdrRow, cL)
    .setValue("総合平均")
    .setBackground("#34a853")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
  sheet
    .getRange(hdrRow, cM)
    .setValue("平均利用率(" + cap + "名)")
    .setBackground("#34a853")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  if (lastData < firstData) return;

  for (let r = firstData; r <= lastData; r++) {
    sheet
      .getRange(r, cL)
      .setFormula("=IFERROR(ROUND(AVERAGE(B" + r + ":K" + r + "),2),\"\")");
    sheet
      .getRange(r, cM)
      .setFormula("=IF(L" + r + ">0,L" + r + "/" + cap + ",0)");
  }

  sheet.getRange(firstData, cL, lastData - firstData + 1, 1).setNumberFormat("0.00");
  sheet.getRange(firstData, cM, lastData - firstData + 1, 1).setNumberFormat("0.0%");
  sheet
    .getRange(firstData, cM, lastData - firstData + 1, 1)
    .setBackground("#fff8e1")
    .setFontColor("#b06000")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}

function applyUsageStatusVisuals_(sheet, layout, mode, dims) {
  const cfg = USAGE_STATUS_SYNC;
  const dashEnd = layout.rowDash + dims.dashRows - 1;
  const firstData = layout.rowDash + cfg.DASH_HEADER_ROWS;
  const dataHdrRow = layout.rowData + 1;

  sheet.setConditionalFormatRules([]);

  sheet.getRange(layout.rowDashHdr, 1, 1, cfg.COL_UTIL_RATE).setBorder(true, false, false, false, false, false, "#1a73e8", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(layout.rowDataHdr, 1, 1, cfg.COL_UTIL_RATE).setBorder(true, false, false, false, false, false, "#666666", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  if (dashEnd >= layout.rowDash) {
    sheet
      .getRange(layout.rowDash, 1, dashEnd - layout.rowDash + 1, cfg.COL_UTIL_RATE)
      .setBorder(true, true, true, true, true, true, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);
  }

  if (dashEnd >= firstData) {
    sheet
      .getRange(firstData, 1, dashEnd - firstData + 1, 1)
      .setBackground("#f8f9fa")
      .setFontWeight("bold");
    applyUsageDashConditionalFormat_(sheet, firstData, dashEnd);
  }

  if (dataHdrRow <= layout.rowData + dims.dataRows - 1) {
    sheet
      .getRange(dataHdrRow, 1, 1, dims.dataCols)
      .setBackground("#f3f3f3")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
  }

  sheet.setColumnWidth(1, 140);
  sheet.setFrozenRows(4);
}

function applyUsageDashConditionalFormat_(sheet, firstRow, lastRow) {
  const range = sheet.getRange(firstRow, 2, lastRow - firstRow + 1, 10);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(10)
      .setFontColor("#d32f2f")
      .setBackground("#fce8e6")
      .setBold(true)
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(9)
      .setFontColor("#137333")
      .setBackground("#e6f4ea")
      .setBold(true)
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(7)
      .setFontColor("#b06000")
      .setBackground("#fef7e0")
      .setBold(true)
      .setRanges([range])
      .build()
  ];
  sheet.setConditionalFormatRules(rules);
}

function buildUsageStatusShell_(sheet, cfg, layout) {
  sheet
    .getRange(1, 1, 1, 13)
    .merge()
    .setValue("施設利用状況")
    .setFontWeight("bold")
    .setFontSize(14)
    .setBackground("#1a73e8")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  sheet.getRange(2, 1).setValue("表示").setFontWeight("bold").setBackground("#e8f0fe").setHorizontalAlignment("center");

  const labels = cfg.MODES.map(function (m) {
    return m.switchLabel;
  });
  sheet.getRange(cfg.SWITCH_CELL).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(labels, true).setAllowInvalid(false).build()
  );
  sheet.getRange(cfg.SWITCH_CELL).setBackground("#fff9c4").setFontWeight("bold").setHorizontalAlignment("center");

  sheet
    .getRange(2, 3, 1, 11)
    .merge()
    .setValue("FW/TM 切替 ｜ 上=ダッシュボード（L・Mは数式） ｜ 下=MAX人数 ｜ 利用人数ブックとリアルタイム連動")
    .setFontColor("#444")
    .setFontSize(9)
    .setVerticalAlignment("middle")
    .setBackground("#f8f9fa");

  sheet.getRange(1, 1, 2, 13).setBorder(true, true, true, true, false, false, "#dadce0", SpreadsheetApp.BorderStyle.SOLID);

  sheet.setColumnWidth(2, 160);
  for (let c = 4; c <= 13; c++) sheet.setColumnWidth(c, 82);

  sheet.getRange(layout.rowDashHdr, 1).setValue("");
  sheet.getRange(layout.rowDataHdr, 1).setValue("");
}

function cacheUsageSourceDimensions_(srcSs, cfg) {
  const byKey = {};
  let maxDashRows = 1;
  let maxDataRows = 1;

  for (let i = 0; i < cfg.MODES.length; i++) {
    const mode = cfg.MODES[i];
    const dash = srcSs.getSheetByName(mode.dashSheet);
    const data = srcSs.getSheetByName(mode.dataSheet);
    if (!dash) throw new Error("利用人数ブックに「" + mode.dashSheet + "」がありません。");
    if (!data) throw new Error("利用人数ブックに「" + mode.dataSheet + "」がありません。");

    const dashRows = Math.max(dash.getLastRow(), 1);
    const dashCols = Math.max(dash.getLastColumn(), 1);
    const dataRows = Math.max(data.getLastRow(), 1);
    const dataCols = Math.max(data.getLastColumn(), 1);

    byKey[mode.key] = { dashRows: dashRows, dashCols: dashCols, dataRows: dataRows, dataCols: dataCols };
    maxDashRows = Math.max(maxDashRows, dashRows);
    maxDataRows = Math.max(maxDataRows, dataRows);
  }

  const layout = {
    rowDashHdr: 4,
    rowDash: 5,
    rowDataHdr: 5 + maxDashRows + 2,
    rowData: 5 + maxDashRows + 3,
    maxDashRows: maxDashRows,
    maxDataRows: maxDataRows
  };

  const props = PropertiesService.getDocumentProperties();
  props.setProperty(cfg.PROP_DIMS, JSON.stringify(byKey));
  props.setProperty(cfg.PROP_LAYOUT, JSON.stringify(layout));
}

function getUsageLayout_(cfg) {
  const raw = PropertiesService.getDocumentProperties().getProperty(cfg.PROP_LAYOUT);
  if (raw) return JSON.parse(raw);
  return { rowDashHdr: 4, rowDash: 5, rowDataHdr: 32, rowData: 33, maxDashRows: 25, maxDataRows: 500 };
}

function loadUsageDimensions_(cfg) {
  const raw = PropertiesService.getDocumentProperties().getProperty(cfg.PROP_DIMS);
  if (!raw) throw new Error("利用状況のレイアウト未設定です。「利用状況をリアルタイム連携」を実行してください。");
  return JSON.parse(raw);
}

function parseUsageModeKey_(displayValue) {
  const s = String(displayValue || "");
  if (s.indexOf("TM") >= 0) return "TM";
  return "FW";
}

function findUsageMode_(cfg, key) {
  for (let i = 0; i < cfg.MODES.length; i++) {
    if (cfg.MODES[i].key === key) return cfg.MODES[i];
  }
  return cfg.MODES[0];
}

function getUsageStatusSheet_(cfg) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.DEST_SHEET_NAME);
  if (!sheet) {
    throw new Error("「" + cfg.DEST_SHEET_NAME + "」シートがありません。先に「利用状況をリアルタイム連携」を実行してください。");
  }
  return sheet;
}

function buildImportRangeFormula_(spreadsheetId, sheetName, r0, c0, numRows, numCols) {
  const a1 =
    columnToLetter_(c0) +
    r0 +
    ":" +
    columnToLetter_(c0 + numCols - 1) +
    (r0 + numRows - 1);
  return '=IMPORTRANGE("' + spreadsheetId + '","' + sheetName + "!" + a1 + '")';
}

function ensureUsageStatusSheet_(ss, cfg) {
  let sheet = ss.getSheetByName(cfg.DEST_SHEET_NAME);
  if (sheet) return sheet;

  const legacy = ss.getSheetByName(cfg.LEGACY_SHEET_NAME);
  if (legacy) {
    legacy.setName(cfg.DEST_SHEET_NAME);
    return legacy;
  }
  return ss.insertSheet(cfg.DEST_SHEET_NAME);
}