/**
 * リフォーマーダッシュボード集計
 * - deta から ⑦予約数・枠数/キャンセル（非表示行）を月列 B〜M に書込
 * - ⑧稼働率・⑨キャンセル率は月ごとの数式（手修正しやすい）
 * - ⑥の直下に ⑦⑧⑨ を挿入し、既存の年度表と1つにまとめる
 */

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
