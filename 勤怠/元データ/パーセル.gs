/**
 * 【経堂店】シフト管理システム API (Code.gs)
 * * 機能:
 * 1. 「2026年〇月シフト」などのマスタシートを直接読み取る
 * 2. 勤務時間・公休・有給の集計はスプレッドシート側で行うため、本APIでは集計しない
 * 3. みとさんのイレギュラー対応（空欄維持）などの表示整形を行う
 * 4. Vercel向けにJSONデータを配信する
 */

const APP_CONFIG = {
  INPUT: {
    DATE_ROW: 3,       // 日付は3行目
    DOW_ROW: 4,        // 曜日は4行目
    START_ROW: 5,      // スタッフデータは5行目から
    NAME_COL: 2,       // 名前はB列(2列目)
    DATA_START_COL: 3  // データはC列(3列目)から
  }
};

function doGet(e) {
  const data = getShiftDataForWeb();
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getShiftDataForWeb() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 「シフト」を含み、「印刷用」を含まないシートを探す
    const sheets = ss.getSheets();
    const sheet = sheets.find(s => s.getName().includes("シフト") && !s.getName().includes("印刷用"));
    
    if (!sheet) {
      return { error: "「〇月シフト」のシートが見つかりませんでした。" };
    }

    const lastCol = sheet.getLastColumn();
    // 共同エリアの終了行（23行目あたりまで取得しておけば安心）
    const dataRange = sheet.getRange(1, 1, 30, lastCol).getValues();

    // 1. タイトル（月）の取得
    const sheetName = sheet.getName();
    const monthMatch = sheetName.match(/(\d{1,2})月/);
    const monthTitle = monthMatch ? `${monthMatch[1]}月 経堂` : sheetName;

    // 2. 日付ヘッダーの生成
    const rawDates = dataRange[APP_CONFIG.INPUT.DATE_ROW - 1];
    let dates = [];
    
    // C列からループ
    for (let c = APP_CONFIG.INPUT.DATA_START_COL - 1; c < lastCol; c++) {
      // ★C3セル(最初の列)は強制的に「1日」とする
      if (c === APP_CONFIG.INPUT.DATA_START_COL - 1) {
        dates.push("1日");
      } else {
        const val = rawDates[c];
        if (val instanceof Date) {
          dates.push(val.getDate() + "日");
        } else if (val && String(val).match(/\d/)) {
          // 数字が含まれていれば日付とみなす
          dates.push(String(val).replace("日", "") + "日");
        } else {
          // 空欄などが続いたら終了とみなすことも可能だが、今回は継続
          if (val) dates.push(String(val));
        }
      }
    }
    const dateCount = dates.length; // 取得できた日付の数だけループする

    // 3. 曜日ヘッダー
    const rawDows = dataRange[APP_CONFIG.INPUT.DOW_ROW - 1];
    const dows = rawDows.slice(APP_CONFIG.INPUT.DATA_START_COL - 1, APP_CONFIG.INPUT.DATA_START_COL - 1 + dateCount)
                        .map(d => String(d).trim().charAt(0));

    // 4. スタッフデータの抽出
    let staffs = [];
    
    // 5行目から2行刻みでループ
    for (let r = APP_CONFIG.INPUT.START_ROW - 1; r < dataRange.length - 1; r += 2) {
      const rowShift = dataRange[r];
      const rowMemo = dataRange[r+1]; // 次の行をメモとする
      
      const name = String(rowShift[APP_CONFIG.INPUT.NAME_COL - 1]).trim();
      // 名前がない、または「メモ」「MEMO」などはスキップ
      if (!name || name === "null" || name === "MEMO" || name === "Memo") continue;

      const isMito = name.includes("みと");
      
      let shifts = [];
      let memos = [];

      for (let c = 0; c < dateCount; c++) {
        const colIdx = APP_CONFIG.INPUT.DATA_START_COL - 1 + c;
        let shiftText = String(rowShift[colIdx] || "").trim();
        let memoText = rowMemo ? String(rowMemo[colIdx] || "").trim() : "";

        // --- データ整形ロジック（表示の統一） ---
        // 休み判定
        if (shiftText === "×" || shiftText === "休み" || shiftText === "公休" || shiftText === "休") {
          shiftText = "×"; 
        } 
        // 空欄判定
        else if (shiftText === "") {
          if (!isMito) { // みとさん以外は空欄＝公休扱い
            shiftText = "×"; 
          }
        }
        // 有給判定
        else if (shiftText.includes("有給") || shiftText.includes("有休")) {
          shiftText = "有給"; 
        }

        shifts.push(shiftText);
        memos.push(memoText);
      }

      staffs.push({
        name: name,
        shifts: shifts,
        memos: memos,
        total: "",
        holidays: "",
        paidLeave: ""
      });
    }

    return {
      month: monthTitle,
      dates: dates,
      dows: dows,
      staffs: staffs
    };

  } catch (e) {
    return { error: "GASエラー: " + e.toString() };
  }
}