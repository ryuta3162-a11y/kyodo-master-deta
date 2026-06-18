/**
 * アイデアプール連携（Pilates / 会員数）
 *
 * 【重要】Code.gs（AirReserve本体）は変更しません。
 * このファイルだけ Apps Script に追加してください。
 *
 * 元: 経堂リフォーマーチェック表 → シート「会員数」
 * 先: 経堂リニューアル　アイデアプール → シート「リフォーマー」（内容は全く同じコピー）
 */

const IDEA_POOL_SYNC = {
  REFORMER_SPREADSHEET_ID: "1Qc1l2pqkoMAgGae_7epfCcvj62a3O_BPbL-ezWDR5pc",
  IDEA_POOL_SPREADSHEET_ID: "1rayTPZSL9X0UsI9Gv2SJthlBviT2WUvvHCJvC6zroIw",
  SOURCE_SHEET_NAME: "会員数",
  DEST_SHEET_NAME: "リフォーマー"
};

/**
 * ※ onOpen は Code.gs と競合するため、このファイルには置きません。
 * 通常はアイデアプール側 SyncFromReformer.gs を使います（プッシュ用の参考実装）。
 *
 * 会員数シートの内容を、アイデアプールの「リフォーマー」にそのままコピー
 */
function syncMemberCountToIdeaPool() {
  const cfg = IDEA_POOL_SYNC;
  const ui = SpreadsheetApp.getUi();

  const srcSs = SpreadsheetApp.openById(cfg.REFORMER_SPREADSHEET_ID);
  const srcSheet = srcSs.getSheetByName(cfg.SOURCE_SHEET_NAME);
  if (!srcSheet) {
    ui.alert("同期エラー", "元シート「" + cfg.SOURCE_SHEET_NAME + "」が見つかりません。", ui.ButtonSet.OK);
    return;
  }

  const dstSs = SpreadsheetApp.openById(cfg.IDEA_POOL_SPREADSHEET_ID);
  let dstSheet = dstSs.getSheetByName(cfg.DEST_SHEET_NAME);
  if (!dstSheet) {
    dstSheet = dstSs.insertSheet(cfg.DEST_SHEET_NAME);
  }

  const lastRow = srcSheet.getLastRow();
  const lastCol = srcSheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    dstSheet.clear();
    ui.alert("同期完了", "元の会員数シートが空のため、リフォーマーをクリアしました。", ui.ButtonSet.OK);
    return;
  }

  const range = srcSheet.getRange(1, 1, lastRow, lastCol);
  const values = range.getValues();
  const displays = range.getDisplayValues();

  dstSheet.clear();
  dstSheet.getRange(1, 1, lastRow, lastCol).setValues(values);
  // 見た目も近づける（日付・数値の表示）
  dstSheet.getRange(1, 1, lastRow, lastCol).setNumberFormats(
    srcSheet.getRange(1, 1, lastRow, lastCol).getNumberFormats()
  );

  ui.alert(
    "同期完了",
    "会員数（" + lastRow + "行×" + lastCol + "列）を\n" +
      "アイデアプール「" + cfg.DEST_SHEET_NAME + "」に反映しました。",
    ui.ButtonSet.OK
  );
}
