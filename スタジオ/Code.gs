/**
 * ホットスタジオの複数月のシートを読み込み、
 * サマリー（全体の参加率・年齢 ＋ 時間帯別の参加率・年齢）と、
 * 会員ごとの参加回数・氏名・年齢を出力するスクリプトです。
 */
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
  
  // --- ユーティリティ：省略された会員番号を10桁に復元する関数 ---
  function normalizeId(id) {
    id = String(id).trim();
    if (id.length >= 10) return id; 
    
    if (id.includes("-")) {
      const parts = id.split("-");
      const prefix = parts[0];
      const suffix = parts[1];
      return prefix + suffix.padStart(10 - prefix.length, "0");
    } else {
      return "1304" + id.padStart(6, "0");
    }
  }

  // --- ユーティリティ：年齢・年代の集計を計算する関数 ---
  function calcStats(agesArray, uniqueCount, totalMembers) {
    let rate = "-";
    if (typeof totalMembers === "number" && totalMembers > 0) {
      rate = (uniqueCount / totalMembers * 100).toFixed(1) + "%";
    }
    
    if (agesArray.length === 0) return { rate: rate, avg: "-", pct: "-" };
    
    const sum = agesArray.reduce((a, b) => a + b, 0);
    const avg = (sum / agesArray.length).toFixed(1) + "歳";
    
    const gens = { "20代以下":0, "30代":0, "40代":0, "50代":0, "60代":0, "70代以上":0 };
    agesArray.forEach(age => {
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
        genStr.push(`${g}:${pct}%`);
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
      const mId = normalizeId(rawMId); 
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
    SpreadsheetApp.getUi().alert(`「${memberDataSheetName}」シートが見つかりません。シート名を変更するか、シートを作成してください。`);
    return;
  }
  
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(targetSheetName);
  } else {
    targetSheet.clear();
  }
  
  const allSheets = ss.getSheets();
  const attendanceCount = {}; 
  
  const monthlyUniqueUsers = {}; 
  const monthlyAges = {}; 
  const monthlyUniqueUsersByPeriod = {};
  const monthlyAgesByPeriod = {};
  
  const targetMonths = []; 
  
  // --- データ抽出ロジック（複数シート対応） ---
  for (let s = 0; s < allSheets.length; s++) {
    const sheet = allSheets[s];
    const sheetName = sheet.getName();
    
    if (/\d+月/.test(sheetName) && !sheetName.includes(targetSheetName) && !sheetName.includes("原本") && !sheetName.includes("休講") && sheetName !== memberDataSheetName) {
      
      const monthMatch = sheetName.match(/(\d+月)/);
      if (!monthMatch) continue;
      const monthKey = monthMatch[1];
      
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
        
        // ★行番号による時間帯の絶対指定
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
                  const memberId = normalizeId(rawMemberId);
                  
                  if (!monthlyUniqueUsers[monthKey].has(memberId)) {
                    monthlyUniqueUsers[monthKey].add(memberId);
                    if (memberInfo[memberId] && memberInfo[memberId].age !== null) {
                      monthlyAges[monthKey].push(memberInfo[memberId].age);
                    }
                  }
                  
                  if (currentPeriod !== "不明") {
                    if (!monthlyUniqueUsersByPeriod[monthKey][currentPeriod].has(memberId)) {
                      monthlyUniqueUsersByPeriod[monthKey][currentPeriod].add(memberId);
                      if (memberInfo[memberId] && memberInfo[memberId].age !== null) {
                        monthlyAgesByPeriod[monthKey][currentPeriod].push(memberInfo[memberId].age);
                      }
                    }
                  }
                  
                  if (!attendanceCount[memberId]) {
                    attendanceCount[memberId] = { total: 0 };
                  }
                  if (!attendanceCount[memberId][monthKey]) {
                    attendanceCount[memberId][monthKey] = 0;
                  }
                  attendanceCount[memberId][monthKey]++;
                  attendanceCount[memberId].total++;
                }
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
  
  targetMonths.sort((a, b) => parseInt(a) - parseInt(b));
  
  let currentRow = 1;
  
  const summaryHeader = [[
    "対象月", "月初会員数", "利用者数(全体)", "利用率(全体)", "全体平均年齢", "全体年代比率",
    "【朝】利用者数", "【朝】参加率", "【朝】平均年齢", "【朝】年代比率",
    "【昼】利用者数", "【昼】参加率", "【昼】平均年齢", "【昼】年代比率",
    "【夜】利用者数", "【夜】参加率", "【夜】平均年齢", "【夜】年代比率"
  ]];
  
  targetSheet.getRange(currentRow, 1, 1, summaryHeader[0].length).setValues(summaryHeader)
    .setBackground("#f3f3f3").setFontWeight("bold");
  currentRow++;
  
  for (let month of targetMonths) {
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
    
    targetSheet.getRange(currentRow, 1, 1, rowData.length).setValues([rowData]);
    currentRow++;
  }
  
  currentRow += 2; 
  
  // ★ 名簿の形式に合わせて、会員番号・氏名・年齢を左側に固定
  const listHeader = ["会員番号", "会員氏名", "年齢", "合計参加回数"];
  for (let month of targetMonths) {
    listHeader.push(month + " 参加回数");
  }
  
  const outputData = [listHeader];
  const sortedMemberIds = Object.keys(attendanceCount).sort((a, b) => attendanceCount[b].total - attendanceCount[a].total);
  
  for (let memberId of sortedMemberIds) {
    const name = memberInfo[memberId] ? memberInfo[memberId].name : "名簿に未登録";
    const age = (memberInfo[memberId] && memberInfo[memberId].age !== null) ? memberInfo[memberId].age : "不明";
    
    const rowData = [memberId, name, age, attendanceCount[memberId].total];
    for (let month of targetMonths) {
      rowData.push(attendanceCount[memberId][month] || 0); 
    }
    
    outputData.push(rowData);
  }
  
  targetSheet.getRange(currentRow, 1, outputData.length, outputData[0].length).setValues(outputData);
  targetSheet.getRange(currentRow, 1, 1, outputData[0].length)
    .setBackground("#e2efda").setFontWeight("bold");
  
  SpreadsheetApp.getUi().alert("すべての集計が完了しました！時間帯別のサマリー（朝・昼・夜）を確認してください。");
}
