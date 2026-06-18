/**
 * アイデアプール側 — リフォーマーデータ取り込み
 *
 * Phase 1: チェック表「会員数」→ シート「リフォーマー」（値＋書式コピー）
 * Phase 2: 年度表の ⑥ 直下に ⑦⑧⑨ を統合（ReformerDashboardAggregate.gs）
 *
 * Pilates Code.gs は変更しません。
 */

const REFORMER_SYNC = {
  REFORMER_SPREADSHEET_ID: "1Qc1l2pqkoMAgGae_7epfCcvj62a3O_BPbL-ezWDR5pc",
  SOURCE_SHEET_NAME: "会員数",
  DEST_SHEET_NAME: "リフォーマー",
  DETA_SHEET_NAME: "deta"
};

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

/** 集計本体は ReformerDashboardAggregate.gs */
