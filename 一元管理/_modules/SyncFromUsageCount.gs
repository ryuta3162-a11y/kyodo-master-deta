/**
 * 開発用 — 利用状況モジュール（Apps Script には Code.gs 統合版を貼る）
 */
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