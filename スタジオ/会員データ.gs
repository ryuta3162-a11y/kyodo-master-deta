/**
 * ホットスタジオの複数月のシートを読み込み、
 * サマリー（全体の参加率・年齢 ＋ 時間帯別の参加率・年齢）と、
 * 会員ごとの参加回数・氏名・年齢を出力するスクリプトです。
 *
 * ※ F列（6列目）・D6・E6 など手入力セルは上書きしません。
 * ※ G6 … 352番台の名簿ユニーク人数（自動）
 * ※ H6〜 … 352番台会員の月別参加人数（4月→H, 5月→I … 各月ユニーク人数）
 */

const MANUAL_COL = 6; // F列（手入力・GAS非更新）
const G6_ROW = 6;
const G6_COL = 7;
const LIST_MONTH_START_COL = 8; // H列から月別（H=4月, I=5月, J=6月…）

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('ホットスタジオ')
    .addItem('集計を更新', 'aggregateAttendance')
    .addToUi();
}

/** 名簿・6行目の月別列（H列から連続） */
function listMonthCol_(monthIndex) {
  return LIST_MONTH_START_COL + monthIndex;
}

function normalizeId_(id) {
  id = String(id).trim().replace(/-/g, "");
  if (id.length >= 10) return id;

  if (id.indexOf("352") === 0) {
    const suffix = id.substring(3);
    return "352" + suffix.padStart(7, "0");
  }
  if (id.length < 10 && /^\d+$/.test(id)) {
    return "1304" + id.padStart(6, "0");
  }
  return id;
}

/** 閉店以降（352番台）会員か — 正規化後に 352 で始まる10桁 */
function is352SeriesMember_(id) {
  const n = normalizeId_(String(id || "").trim());
  return n.length >= 10 && n.indexOf("352") === 0;
}

/** 352番台の同一人物判定用キー（正規化済み10桁） */
function canonical352Id_(rawOrNormalized) {
  const n = normalizeId_(rawOrNormalized);
  return is352SeriesMember_(n) ? n : "";
}

/**
 * 6行目 H6, I6, J6… に、各月シートで参加した352番台会員のユニーク人数を書き込む
 * （D6の名前リストではなく、出席表の会員番号が352で始まるかで判定）
 */
function writeRow6Monthly352Counts_(sheet, targetMonths, monthlyUniqueUsers) {
  for (let i = 0; i < targetMonths.length; i++) {
    const month = targetMonths[i];
    let count = 0;
    const users = monthlyUniqueUsers[month];
    if (users) {
      users.forEach(function (id) {
        if (is352SeriesMember_(id)) count++;
      });
    }
    sheet.getRange(G6_ROW, listMonthCol_(i)).setValue(count);
  }
}

/** シート名から年（2026年4月 → 2026）。なければ 0 */
function extractYearFromSheetName_(sheetName) {
  const m = String(sheetName).match(/(\d{4})年/);
  return m ? parseInt(m[1], 10) : 0;
}

/** サマリー行を書き込み（F列はスキップ） */
function writeSummaryRow_(sheet, row, rowData) {
  sheet.getRange(row, 1, 1, 5).setValues([rowData.slice(0, 5)]);
  sheet.getRange(row, 7, 1, rowData.length - 6).setValues([rowData.slice(6)]);
}

/** F列と6行目をスナップショット（D6/E6 手入力・H6〜は後で上書き） */
function snapshotManualCells_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastCol = Math.max(sheet.getLastColumn(), 18);
  const colF = sheet.getRange(1, MANUAL_COL, lastRow, 1).getValues();
  const row6 = sheet.getRange(G6_ROW, 1, 1, lastCol).getValues()[0];
  return { colF: colF, row6: row6, lastRow: lastRow, lastCol: lastCol };
}

/** F列・6行目の手入力（G6・H6〜月別は除く）を復元し G6 を書き込む */
function restoreManualCells_(sheet, snapshot, unique352Count, monthCount) {
  if (snapshot.colF.length > 0) {
    sheet.getRange(1, MANUAL_COL, snapshot.colF.length, 1).setValues(snapshot.colF);
  }
  const skipMonthCols = {};
  skipMonthCols[G6_COL] = true;
  for (let i = 0; i < monthCount; i++) {
    skipMonthCols[listMonthCol_(i)] = true;
  }
  if (snapshot.row6.length > 0) {
    for (let c = 0; c < snapshot.row6.length; c++) {
      const colNum = c + 1;
      if (skipMonthCols[colNum]) continue;
      sheet.getRange(G6_ROW, colNum).setValue(snapshot.row6[c]);
    }
  }
  sheet.getRange(G6_ROW, G6_COL).setValue(unique352Count);
}

/** 指定範囲をクリア（F列は除外） */
function clearRangeExcludingColF_(sheet, startRow, numRows, startCol, endCol) {
  if (numRows <= 0) return;
  if (startCol <= MANUAL_COL - 1) {
    const leftEnd = Math.min(endCol, MANUAL_COL - 1);
    if (startCol <= leftEnd) {
      sheet.getRange(startRow, startCol, numRows, leftEnd - startCol + 1).clearContent();
    }
  }
  if (endCol > MANUAL_COL) {
    const rightStart = Math.max(startCol, MANUAL_COL + 1);
    if (rightStart <= endCol) {
      sheet.getRange(startRow, rightStart, numRows, endCol - rightStart + 1).clearContent();
    }
  }
}

function aggregateAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheetName = "集計";

  // =========================================
  // 【設定①】会員情報が入力されているシート名
  // =========================================
  const memberDataSheetName = "会員データ";

  // =========================================
  // 【設定②】月ごとの全体契約者数（月初会員数）
  // =========================================
  const monthlyTotalMembers = {
    "4月": 207,
    "5月": 230,
    "6月": 217
  };

  // --- ユーティリティ：年齢・年代の集計を計算する関数 ---
  function calcStats(agesArray, uniqueCount, totalMembers) {
    let rate = "-";
    if (typeof totalMembers === "number" && totalMembers > 0) {
      rate = (uniqueCount / totalMembers * 100).toFixed(1) + "%";
    }

    if (agesArray.length === 0) return { rate: rate, avg: "-", pct: "-" };

    const sum = agesArray.reduce(function (a, b) { return a + b; }, 0);
    const avg = (sum / agesArray.length).toFixed(1) + "歳";

    const gens = { "20代以下": 0, "30代": 0, "40代": 0, "50代": 0, "60代": 0, "70代以上": 0 };
    agesArray.forEach(function (age) {
      if (age < 30) gens["20代以下"]++;
      else if (age < 40) gens["30代"]++;
      else if (age < 50) gens["40代"]++;
      else if (age < 60) gens["50代"]++;
      else if (age < 70) gens["60代"]++;
      else gens["70代以上"]++;
    });

    const genStr = [];
    for (let g in gens) {
      if (gens[g] > 0) {
        const pct = (gens[g] / agesArray.length * 100).toFixed(1);
        genStr.push(g + ":" + pct + "%");
      }
    }
    return { rate: rate, avg: avg, pct: genStr.join(", ") };
  }

  // --- 会員データの取得と格納 ---
  const memberInfo = {};
  const memberSheet = ss.getSheetByName(memberDataSheetName);

  if (memberSheet) {
    const mData = memberSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < mData.length; i++) {
      const rawMId = String(mData[i][0]).trim();
      const normalized = normalizeId_(rawMId);
      const mId = canonical352Id_(normalized) || normalized;
      const mName = String(mData[i][1]).trim();
      const mAge = parseInt(mData[i][2]);

      if (mId !== "") {
        memberInfo[mId] = {
          name: mName !== "" ? mName : "名前なし",
          age: isNaN(mAge) ? null : mAge
        };
      }
    }
  } else {
    SpreadsheetApp.getUi().alert("「" + memberDataSheetName + "」シートが見つかりません。シート名を変更するか、シートを作成してください。");
    return;
  }

  let targetSheet = ss.getSheetByName(targetSheetName);
  const isNewSheet = !targetSheet;
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
  }
  const manualSnapshot = isNewSheet
    ? { colF: [], row6: [], lastRow: 0, lastCol: 0 }
    : snapshotManualCells_(targetSheet);
  const oldLastRow = manualSnapshot.lastRow;

  const allSheets = ss.getSheets();
  const attendanceCount = {};

  const monthlyUniqueUsers = {};
  const monthlyAges = {};
  const monthlyUniqueUsersByPeriod = {};
  const monthlyAgesByPeriod = {};

  const targetMonths = [];
  const monthSheetByKey = {};

  // --- 対象月シートを選定（同月が複数ある場合は年が新しい方を優先: 2026年4月 など） ---
  for (let s = 0; s < allSheets.length; s++) {
    const sheet = allSheets[s];
    const sheetName = sheet.getName();

    if (!/\d+月/.test(sheetName) || sheetName.includes(targetSheetName) || sheetName.includes("原本") ||
        sheetName.includes("休講") || sheetName === memberDataSheetName) {
      continue;
    }

    const monthMatch = sheetName.match(/(\d+月)/);
    if (!monthMatch) continue;
    const monthKey = monthMatch[1];
    const year = extractYearFromSheetName_(sheetName);
    const prev = monthSheetByKey[monthKey];
    if (!prev || year > prev.year) {
      monthSheetByKey[monthKey] = { sheet: sheet, year: year };
    }
  }

  const monthKeys = Object.keys(monthSheetByKey);
  if (monthKeys.length === 0) {
    SpreadsheetApp.getUi().alert("対象となる月のシートが見つかりませんでした。");
    return;
  }

  monthKeys.sort(function (a, b) { return parseInt(a) - parseInt(b); });

  // --- データ抽出 ---
  for (let m = 0; m < monthKeys.length; m++) {
    const monthKey = monthKeys[m];
    const sheet = monthSheetByKey[monthKey].sheet;

    targetMonths.push(monthKey);

    monthlyUniqueUsers[monthKey] = new Set();
    monthlyAges[monthKey] = [];
    monthlyUniqueUsersByPeriod[monthKey] = { "朝": new Set(), "昼": new Set(), "夜": new Set(), "不明": new Set() };
    monthlyAgesByPeriod[monthKey] = { "朝": [], "昼": [], "夜": [], "不明": [] };

    const displayData = sheet.getDataRange().getDisplayValues();
    const maxRow = Math.min(displayData.length, 170);
    const maxCol = displayData[0].length;

    for (let r = 3; r < maxRow; r++) {

      const rowNum = r + 1;
      let currentPeriod = "不明";

      if (rowNum >= 4 && rowNum <= 57) {
        currentPeriod = "朝";
      } else if (rowNum >= 58 && rowNum <= 128) {
        currentPeriod = "昼";
      } else if (rowNum >= 129 && rowNum <= 170) {
        currentPeriod = "夜";
      }

      for (let c = 1; c < maxCol; c += 4) {
        let participantColIndex = -1;
        if (r + 1 < maxRow) {
          for (let i = 0; i < 4; i++) {
            if (c + i < maxCol && displayData[r + 1][c + i] === "参加者") {
              participantColIndex = c + i;
              break;
            }
          }
        }

        if (participantColIndex !== -1 && displayData[r][c] !== "" && displayData[r][c] !== undefined) {
          for (let j = 0; j < 12; j++) {
            if (r + 2 + j < maxRow) {
              const rawMemberId = String(displayData[r + 2 + j][participantColIndex]).trim();

              if (rawMemberId !== "" && /^[A-Za-z0-9\-]+$/.test(rawMemberId)) {
                const memberId = normalizeId_(rawMemberId);
                const idFor352Set = canonical352Id_(memberId) || memberId;

                if (!monthlyUniqueUsers[monthKey].has(idFor352Set)) {
                  monthlyUniqueUsers[monthKey].add(idFor352Set);
                  if (memberInfo[memberId] && memberInfo[memberId].age !== null) {
                    monthlyAges[monthKey].push(memberInfo[memberId].age);
                  } else if (memberInfo[idFor352Set] && memberInfo[idFor352Set].age !== null) {
                    monthlyAges[monthKey].push(memberInfo[idFor352Set].age);
                  }
                }

                if (currentPeriod !== "不明") {
                  if (!monthlyUniqueUsersByPeriod[monthKey][currentPeriod].has(idFor352Set)) {
                    monthlyUniqueUsersByPeriod[monthKey][currentPeriod].add(idFor352Set);
                    if (memberInfo[memberId] && memberInfo[memberId].age !== null) {
                      monthlyAgesByPeriod[monthKey][currentPeriod].push(memberInfo[memberId].age);
                    }
                  }
                }

                if (!attendanceCount[idFor352Set]) {
                  attendanceCount[idFor352Set] = { total: 0 };
                }
                if (!attendanceCount[idFor352Set][monthKey]) {
                  attendanceCount[idFor352Set][monthKey] = 0;
                }
                attendanceCount[idFor352Set][monthKey]++;
                attendanceCount[idFor352Set].total++;
              }
            }
          }
        }
      }
    }
  }

  if (targetMonths.length === 0) {
    SpreadsheetApp.getUi().alert("対象となる月のシートが見つかりませんでした。");
    return;
  }

  let currentRow = 1;

  const summaryHeader = [
    "対象月", "月初会員数", "利用者数(全体)", "利用率(全体)", "全体平均年齢", "全体年代比率",
    "【朝】利用者数", "【朝】参加率", "【朝】平均年齢", "【朝】年代比率",
    "【昼】利用者数", "【昼】参加率", "【昼】平均年齢", "【昼】年代比率",
    "【夜】利用者数", "【夜】参加率", "【夜】平均年齢", "【夜】年代比率"
  ];

  writeSummaryRow_(targetSheet, currentRow, summaryHeader);
  targetSheet.getRange(currentRow, 1, 1, 5).setBackground("#f3f3f3").setFontWeight("bold");
  targetSheet.getRange(currentRow, 7, 1, summaryHeader.length - 6).setBackground("#f3f3f3").setFontWeight("bold");
  currentRow++;

  for (let m = 0; m < targetMonths.length; m++) {
    const month = targetMonths[m];
    const totalMembers = monthlyTotalMembers[month] || "未設定";
    const uniqueUsersCount = monthlyUniqueUsers[month].size;

    const overallStats = calcStats(monthlyAges[month], uniqueUsersCount, totalMembers);
    const mCount = monthlyUniqueUsersByPeriod[month]["朝"].size;
    const morningStats = calcStats(monthlyAgesByPeriod[month]["朝"], mCount, totalMembers);
    const aCount = monthlyUniqueUsersByPeriod[month]["昼"].size;
    const afternoonStats = calcStats(monthlyAgesByPeriod[month]["昼"], aCount, totalMembers);
    const nCount = monthlyUniqueUsersByPeriod[month]["夜"].size;
    const nightStats = calcStats(monthlyAgesByPeriod[month]["夜"], nCount, totalMembers);

    const rowData = [
      month, totalMembers, uniqueUsersCount, overallStats.rate, overallStats.avg, overallStats.pct,
      mCount, morningStats.rate, morningStats.avg, morningStats.pct,
      aCount, afternoonStats.rate, afternoonStats.avg, afternoonStats.pct,
      nCount, nightStats.rate, nightStats.avg, nightStats.pct
    ];

    writeSummaryRow_(targetSheet, currentRow, rowData);
    currentRow++;
  }

  const summaryDataEndRow = currentRow - 1;
  currentRow += 2;

  const listStartRow = currentRow;
  const listHeaderLeft = ["会員番号", "会員氏名", "年齢", "合計参加回数"];
  targetSheet.getRange(listStartRow, 1, 1, 4).setValues([listHeaderLeft])
    .setBackground("#e2efda").setFontWeight("bold");
  for (let i = 0; i < targetMonths.length; i++) {
    targetSheet.getRange(listStartRow, listMonthCol_(i)).setValue(targetMonths[i] + " 参加回数");
  }

  const sortedMemberIds = Object.keys(attendanceCount).sort(function (a, b) {
    return attendanceCount[b].total - attendanceCount[a].total;
  });

  for (let i = 0; i < sortedMemberIds.length; i++) {
    const row = listStartRow + 1 + i;
    const memberId = sortedMemberIds[i];
    const name = memberInfo[memberId] ? memberInfo[memberId].name : "名簿に未登録";
    const age = (memberInfo[memberId] && memberInfo[memberId].age !== null) ? memberInfo[memberId].age : "不明";

    targetSheet.getRange(row, 1, 1, 4).setValues([[memberId, name, age, attendanceCount[memberId].total]]);
    for (let m = 0; m < targetMonths.length; m++) {
      targetSheet.getRange(row, listMonthCol_(m)).setValue(attendanceCount[memberId][targetMonths[m]] || 0);
    }
  }

  const listEndRow = listStartRow + sortedMemberIds.length;

  if (!isNewSheet) {
    // サマリーに残った古い月行を削除（F列・6行目は除外）
    for (let r = summaryDataEndRow + 1; r < listStartRow; r++) {
      if (r === G6_ROW) continue;
      const val = targetSheet.getRange(r, 1).getValue();
      if (val && /^\d+月$/.test(String(val))) {
        clearRangeExcludingColF_(targetSheet, r, 1, 1, 18);
      }
    }
    // 名簿が短くなった場合の余分行を削除（F列は除外）
    if (oldLastRow > listEndRow) {
      const listLastCol = LIST_MONTH_START_COL + targetMonths.length - 1;
      clearRangeExcludingColF_(targetSheet, listEndRow + 1, oldLastRow - listEndRow, 1, Math.max(manualSnapshot.lastCol, listLastCol, 18));
    }
  }

  const unique352Count = sortedMemberIds.filter(function (id) { return is352SeriesMember_(id); }).length;
  restoreManualCells_(targetSheet, manualSnapshot, unique352Count, targetMonths.length);
  writeRow6Monthly352Counts_(targetSheet, targetMonths, monthlyUniqueUsers);

  let alertMsg = "すべての集計が完了しました。\n\n【352番台】\n";
  alertMsg += "G6（全月で1回でも参加）: " + unique352Count + "人\n";
  for (let mi = 0; mi < targetMonths.length; mi++) {
    const month = targetMonths[mi];
    let month352 = 0;
    monthlyUniqueUsers[month].forEach(function (id) {
      if (is352SeriesMember_(id)) month352++;
    });
    const colLetter = String.fromCharCode(72 + mi);
    alertMsg += month + "（" + colLetter + "6）: " + month352 + "人\n";
  }
  const e6val = manualSnapshot.row6.length >= 5 ? manualSnapshot.row6[4] : "";
  if (e6val !== "" && !isNaN(Number(e6val))) {
    alertMsg += "\nE6（在籍数・手入力）: " + e6val + "人";
    alertMsg += "\n※月別参加 > E6 のときは、出席表の番号表記ゆれで同一人物が別番号になっていないか確認してください。";
  }
  alertMsg += "\n\n同一人物が月内に複数レッスン参加しても、月別は1人1カウントです。";
  SpreadsheetApp.getUi().alert(alertMsg);
}
