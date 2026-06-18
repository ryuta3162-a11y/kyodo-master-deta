/**
 * 経堂利用人数集計（作業用コピー）
 *
 * 正本: 元データ/Code.gs
 * ブック ID: 1qWOf5apYTULT3YWw8oF2ZQ8p73jBM4-7Y7NY8adcsXg
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 ジム施設集計システム')
    .addItem('全システム構築 (初回1回のみ実行)', 'generateAllSheets')
    .addToUi();
}

/**
 * 1. 4つのシート（FW/TMの入力シートと分析ダッシュボード）と関数をすべて自動構築するメイン処理
 */
function generateAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ==========================================
  // シートの準備（存在すればクリア、なければ新規作成）
  // ==========================================
  let fwDataSheet = ss.getSheetByName('FW_データ入力');
  if (!fwDataSheet) fwDataSheet = ss.insertSheet('FW_データ入力');
  else fwDataSheet.clear();

  let fwDashSheet = ss.getSheetByName('FW_分析ダッシュボード');
  if (!fwDashSheet) fwDashSheet = ss.insertSheet('FW_分析ダッシュボード');
  else fwDashSheet.clear();

  let tmDataSheet = ss.getSheetByName('TM_データ入力');
  if (!tmDataSheet) tmDataSheet = ss.insertSheet('TM_データ入力');
  else tmDataSheet.clear();

  let tmDashSheet = ss.getSheetByName('TM_分析ダッシュボード');
  if (!tmDashSheet) tmDashSheet = ss.insertSheet('TM_分析ダッシュボード');
  else tmDashSheet.clear();

  const timeSlots = ["0〜4時", "4〜7時", "7〜9時", "9〜11時", "11〜13時", "13〜15時", "15〜17時", "17〜19時", "19〜21時", "21〜0時"];
  
  // 会員数マスタ (2025年4月〜2026年5月)
  const memberData = {
    "2025-4": 1276, "2025-5": 1276, "2025-6": 1282, "2025-7": 1307, "2025-8": 1336, "2025-9": 1335,
    "2025-10": 1327, "2025-11": 1307, "2025-12": 1321, "2026-1": 1409, "2026-2": 1307, "2026-3": 1312,
    "2026-4": 1393, "2026-5": 1416
  };
  const baseMembers = 1416;

  // 祝日リスト
  const holidays = [
    "2025-4-29", "2025-5-3", "2025-5-4", "2025-5-5", "2025-5-6", "2025-7-21", "2025-8-11", "2025-9-15", "2025-9-23", "2025-10-13", "2025-11-3", "2025-11-23", "2025-11-24", "2025-12-30", "2025-12-31",
    "2026-1-1", "2026-1-2", "2026-1-3", "2026-1-12", "2026-2-11", "2026-2-23", "2026-3-20", "2026-4-29", "2026-5-3", "2026-5-4", "2026-5-5", "2026-5-6"
  ];
  
  const dayOfWeekStr = ['日', '月', '火', '水', '木', '金', '土'];

  // ====================================================================
  // 🔨 【FW】フリーウェイト集計（最大収容人数：18名）の構築
  // ====================================================================
  buildDataSheet(fwDataSheet, 'FW', 18, memberData, baseMembers, holidays, timeSlots, dayOfWeekStr);
  buildDashSheet(fwDashSheet, 'FW', 18, memberData, baseMembers, holidays, timeSlots);

  // ====================================================================
  // 🔨 【TM】トレッドミル集計（最大収容人数：12名）の構築
  // ====================================================================
  buildDataSheet(tmDataSheet, 'TM', 12, memberData, baseMembers, holidays, timeSlots, dayOfWeekStr);
  buildDashSheet(tmDashSheet, 'TM', 12, memberData, baseMembers, holidays, timeSlots);

  // 完了アラート
  SpreadsheetApp.getUi().alert(
    '🎉 システム全構築が完了しました！\n\n' +
    '以下の4つのシートが完全に連携した状態で作成されました。\n' +
    '・FW_データ入力 ／ FW_分析ダッシュボード (収容18名基準)\n' +
    '・TM_データ入力 ／ TM_分析ダッシュボード (収容12名基準)\n\n' +
    '【今回の高度なゆらぎ変更】\n' +
    '・同じ曜日でも、天候や突発イベント（±25%）のランダムノイズを日別に付与\n' +
    '・トレッドミル(TM)には「雨の日は外を走れない人が流入して混雑する」独自の逆転仕様を実装\n' +
    '・4月・5月の夜間（19時〜24時）における急増トレンド（1.28倍ブースト）を適合'
  );
}

/**
 * データ入力シートを自動生成・データ入力する関数
 * @param {Sheet} sheet スプレッドシートのシートオブジェクト
 * @param {string} type 'FW' or 'TM'
 * @param {number} maxCapacity 最大収容人数
 */
function buildDataSheet(sheet, type, maxCapacity, memberData, baseMembers, holidays, timeSlots, dayOfWeekStr) {
  sheet.getRange('C1:L1').merge().setValue(`MAX人数入力 (${type}・シミュレーション予測値セット済み)`)
    .setHorizontalAlignment('center').setBackground('#e2e3e5').setFontWeight('bold');
  
  const headers = ['年月', '日付', ...timeSlots, '日平均'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers])
    .setHorizontalAlignment('center').setBackground('#f3f3f3').setFontWeight('bold');

  sheet.setColumnWidth(1, 100); // 年月
  sheet.setColumnWidth(2, 120); // 日付
  for(let i=3; i<=13; i++) sheet.setColumnWidth(i, 80);
  
  let currentRow = 3;
  
  // 2025年4月〜2026年5月までループ処理
  for (let y = 2025; y <= 2026; y++) {
    let startM = (y === 2025) ? 4 : 1;
    let endM = (y === 2026) ? 5 : 12;
    for (let m = startM; m <= endM; m++) {
      
      const currentMembers = memberData[`${y}-${m}`] || baseMembers; 
      const ratio = currentMembers / baseMembers; 
      const daysInMonth = new Date(y, m, 0).getDate();
      
      let startMonthRow = currentRow;
      let monthData = [];
      let monthFormulas = [];
      
      // 1ヶ月分の日付と、ノイズを考慮した複雑な数値を生成
      for (let d = 1; d <= daysInMonth; d++) {
        const dateObj = new Date(y, m - 1, d);
        const dayOfWeek = dateObj.getDay();
        const dow = dayOfWeekStr[dayOfWeek];
        const isHoliday = holidays.includes(`${y}-${m}-${d}`);
        
        let basePattern;
        
        if (type === 'FW') {
          // ==========================================
          // フリーウェイト（FW）専用基本パターン
          // ==========================================
          if (isHoliday) {
            basePattern = [1.8, 2.5, 4.5, 3.0, 5.8, 5.2, 7.5, 8.5, 8.0, 7.5]; // 祝日
          } else {
            switch (dayOfWeek) {
              case 0: // 日曜日
                basePattern = [1.8, 2.5, 4.5, 3.0, 5.8, 5.2, 7.5, 8.5, 8.0, 7.5]; break;
              case 1: // 月曜日
                basePattern = [1.8, 2.5, 3.5, 3.5, 4.2, 4.5, 5.0, 8.5, 9.0, 8.0]; break;
              case 2: // 火曜日
                basePattern = [2.0, 3.0, 4.0, 4.0, 4.8, 5.0, 5.5, 9.5, 9.5, 9.5]; break;
              case 3: // 水曜日
                basePattern = [2.2, 3.2, 4.2, 3.8, 5.0, 5.5, 5.8, 10.5, 11.0, 10.5]; break;
              case 4: // 木曜日
                basePattern = [2.1, 2.9, 3.9, 4.1, 4.7, 5.2, 5.6, 9.8, 10.0, 9.8]; break;
              case 5: // 金曜日
                basePattern = [2.5, 3.0, 4.0, 4.0, 5.0, 5.5, 6.0, 11.0, 12.0, 11.5]; break;
              case 6: // 土曜日
                basePattern = [2.0, 3.0, 5.0, 3.5, 6.5, 6.0, 8.5, 9.5, 9.5, 10.0]; break;
            }
          }
        } else {
          // ==========================================
          // トレッドミル（TM）専用基本パターン (有酸素に特化)
          // ==========================================
          if (isHoliday) {
            basePattern = [1.0, 1.8, 3.0, 4.5, 5.5, 5.0, 6.5, 7.0, 6.0, 3.0]; // 祝日（お昼が中心）
          } else {
            switch (dayOfWeek) {
              case 0: // 日曜日（昼間だらだら混む）
                basePattern = [1.0, 1.8, 3.0, 4.5, 5.5, 5.0, 6.5, 7.0, 6.0, 3.0]; break;
              case 1: // 月曜日（週初めの夜はそこそこ走る）
                basePattern = [1.0, 1.5, 3.5, 2.5, 3.0, 3.5, 4.5, 7.5, 8.5, 5.0]; break;
              case 2: // 火曜日
                basePattern = [1.2, 1.8, 4.0, 3.0, 3.5, 4.0, 5.0, 8.0, 9.0, 5.5]; break;
              case 3: // 水曜日（週の中休み・夜間の有酸素増）
                basePattern = [1.5, 2.0, 4.5, 3.5, 3.8, 4.2, 5.5, 8.5, 9.5, 6.0]; break;
              case 4: // 木曜日
                basePattern = [1.1, 1.7, 3.8, 2.8, 3.2, 3.8, 4.8, 7.8, 8.8, 5.2]; break;
              case 5: // 金曜日（明日は休みなので夜は少し多め）
                basePattern = [1.3, 1.9, 4.0, 3.0, 3.5, 4.0, 5.0, 9.0, 10.0, 6.5]; break;
              case 6: // 土曜日（夕方まで活況）
                basePattern = [1.2, 2.2, 4.5, 4.5, 6.0, 5.5, 7.0, 8.0, 7.5, 5.0]; break;
            }
          }
        }
        
        // 2. 日別の天候・突発ノイズ (FWとTMで個別に乱数を回して連動しないようにする)
        let dayNoise = 1.0;
        const randDay = Math.random();
        
        if (type === 'FW') {
          // FW：通常の雨などの減退
          if (randDay < 0.08) dayNoise = 0.75;      // 大雨
          else if (randDay < 0.16) dayNoise = 1.20; // 混雑イベント
          else if (randDay < 0.28) dayNoise = 0.90; // 少し暇
          else if (randDay < 0.40) dayNoise = 1.10; // 少し混雑
        } else {
          // TM：雨の日はロードランナーが押し寄せるため、逆に「有酸素は上がる」という複雑設定！
          if (randDay < 0.08) dayNoise = 1.25;      // 外が雨なので室内ラン急増！
          else if (randDay < 0.16) dayNoise = 0.75; // 晴れて外が気持ち良いのでみんな外を走り、室内は減少
          else if (randDay < 0.28) dayNoise = 0.90; 
          else if (randDay < 0.40) dayNoise = 1.10; 
        }
        
        // 3. 各セルの予測値に、さらに「微細ゆらぎ」を乗せる
        const rowValues = basePattern.map((val, idx) => {
          let boost = 1.0;
          
          if (m === 4 || m === 5) {
            // 4月、5月の特異ブースト（特に19時〜24時は劇的に伸びる）
            if (idx === 8 || idx === 9) {
              boost = (type === 'FW') ? 1.28 : 1.32; // 夜間に大幅ブースト
            } else if (idx === 7) {
              boost = 1.15; // 夕方
            } else {
              boost = 1.08; // 全体も微増
            }
          }
          
          let calculated = val * ratio * boost * dayNoise;
          
          // 個別の微細ゆらぎ (±15% のブレ + 振れ幅)
          const cellFluctuation = (Math.random() * 0.3 - 0.15) * calculated + (Math.random() * 1.6 - 0.8);
          calculated += cellFluctuation;
          
          // 収容制限（0〜最大定員まで）
          calculated = Math.max(0, Math.min(maxCapacity, calculated));
          
          // 生っぽさを表現するために「整数」に丸める
          return Math.round(calculated);
        });
        
        const yearMonthStr = (d === 1) ? `${y}年${m}月` : ""; 
        monthData.push([yearMonthStr, `${m}月${d}日 (${dow})`, ...rowValues]);
        
        monthFormulas.push([`=IFERROR(ROUND(AVERAGE(C${currentRow}:L${currentRow}), 2), "")`]);
        currentRow++;
      }
      
      // 1ヶ月分のデータをスプレッドシートに書き込み
      sheet.getRange(startMonthRow, 1, daysInMonth, 12).setValues(monthData)
        .setHorizontalAlignment('right');
      sheet.getRange(startMonthRow, 13, daysInMonth, 1).setFormulas(monthFormulas)
        .setBackground('#e8f4fd').setFontColor('#0d47a1').setFontWeight('bold');
      
      // 月ごとの平均行の作成
      let avgRow = currentRow;
      sheet.getRange(avgRow, 1, 1, 2).merge().setValue(`${m}月 平均`)
        .setHorizontalAlignment('center').setBackground('#fff3cd').setFontWeight('bold').setFontColor('#856404');
      
      let colAvgFormulas = [];
      for(let c = 3; c <= 12; c++) {
        const colLetter = String.fromCharCode(64 + c); 
        colAvgFormulas.push(`=IFERROR(ROUND(AVERAGE(${colLetter}${startMonthRow}:${colLetter}${avgRow-1}), 2), "")`);
      }
      sheet.getRange(avgRow, 3, 1, 10).setFormulas([colAvgFormulas])
        .setBackground('#fff3cd').setFontWeight('bold').setFontColor('#856404');
      
      sheet.getRange(avgRow, 13).setFormula(`=IFERROR(ROUND(AVERAGE(M${startMonthRow}:M${avgRow-1}), 2), "")`)
        .setBackground('#ffeeba').setFontWeight('bold').setFontColor('#d32f2f'); 
      
      sheet.getRange(startMonthRow, 1, daysInMonth + 1, 13).setBorder(true, true, true, true, true, true);
      
      currentRow += 2; // 月と月の間をあける
    }
  }

  // 🎨 データ入力シートの 混雑度シグナル（条件付き書式）をセット
  applyColorConditionRules(sheet, currentRow, maxCapacity);
}

/**
 * 分析ダッシュボードを作成する関数
 * @param {Sheet} dashSheet スプレッドシートのダッシュボードオブジェクト
 * @param {string} type 'FW' or 'TM'
 * @param {number} maxCapacity 最大収容人数
 */
function buildDashSheet(dashSheet, type, maxCapacity, memberData, baseMembers, holidays, timeSlots) {
  dashSheet.getRange('A1:M1').merge().setValue(`📊 ${type} 年間分析ダッシュボード (収容定員: ${maxCapacity}名)`)
    .setFontWeight('bold').setFontSize(14).setBackground('#e8f0fe').setFontColor('#1a73e8').setHorizontalAlignment('center');
    
  dashSheet.getRange('A2:M2').merge().setValue(`※ 黄色：7〜8名（混雑始め） ｜ 緑色：9名（適正ピーク） ｜ 赤色：10名以上（満員寸前・制限注意）`)
    .setFontColor('#555').setVerticalAlignment('middle');

  const dashHeaders = ['対象月 (会員数)', ...timeSlots, '総合平均', `平均利用率 (収容${maxCapacity}名)`];
  dashSheet.getRange(3, 1, 1, dashHeaders.length).setValues([dashHeaders])
    .setBackground('#34a853').setFontColor('white').setFontWeight('bold').setHorizontalAlignment('center');
  
  // 各月の「平均」が配置されている行を計算してダッシュボードと連動
  let dashRow = 4;
  let currentRowTracker = 3;
  
  for (let y = 2025; y <= 2026; y++) {
    let startM = (y === 2025) ? 4 : 1;
    let endM = (y === 2026) ? 5 : 12;
    for (let m = startM; m <= endM; m++) {
      const currentMembers = memberData[`${y}-${m}`] || baseMembers; 
      const daysInMonth = new Date(y, m, 0).getDate();
      
      // この月の「平均」が何行目に位置するか
      const targetAvgRow = currentRowTracker + daysInMonth;
      
      dashSheet.getRange(dashRow, 1).setValue(`${y}年${m}月 (${currentMembers}名)`)
        .setBackground('#f8f9fa').setFontWeight('bold');
      
      // データ入力シートの平均セル（C列〜M列）を参照する数式
      let formulas = [];
      for(let c = 3; c <= 13; c++) {
        const colLetter = String.fromCharCode(64 + c); 
        formulas.push(`='${type}_データ入力'!${colLetter}${targetAvgRow}`);
      }
      dashSheet.getRange(dashRow, 2, 1, 11).setFormulas([formulas]).setBackground('#ffffff');
      
      // 利用率 = 総合平均 / maxCapacity
      dashSheet.getRange(dashRow, 13).setFormula(`=L${dashRow}/${maxCapacity}`)
        .setBackground('#fff8e1').setFontColor('#b06000').setFontWeight('bold').setHorizontalAlignment('center')
        .setNumberFormat('0.0%');
      
      dashRow++;
      currentRowTracker = targetAvgRow + 2; // 次の月に移行
    }
  }
  
  const totalMonths = dashRow - 4;
  dashSheet.getRange(3, 1, totalMonths + 1, 13).setBorder(true, true, true, true, true, true);
  dashSheet.setColumnWidth(1, 140);
  for(let i=2; i<=12; i++) dashSheet.setColumnWidth(i, 80);
  dashSheet.setColumnWidth(13, 180);

  // 🎨 ダッシュボードの混雑度シグナル（条件付き書式）をセット
  applyColorConditionRulesForDash(dashSheet, totalMonths);
}

/**
 * データ入力シートの条件付き書式（7〜8名：黄、9名：緑、10名以上：赤）を設定する補助関数
 */
function applyColorConditionRules(sheet, maxRow, maxCapacity) {
  const rules = [];
  const range = sheet.getRange(3, 3, maxRow - 3, 10); // C列〜L列
  
  // 1. 10名以上：赤色（混雑レベル大）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(10)
    .setFontColor('#d32f2f') 
    .setBackground('#fce8e6') 
    .setBold(true)
    .setRanges([range])
    .build());

  // 2. 9名：緑色（適正混雑）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(9)
    .setFontColor('#137333') 
    .setBackground('#e6f4ea') 
    .setBold(true)
    .setRanges([range])
    .build());

  // 3. 7〜8名：黄色（徐々に混雑）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(7)
    .setFontColor('#b06000') 
    .setBackground('#fef7e0') 
    .setBold(true)
    .setRanges([range])
    .build());

  sheet.setConditionalFormatRules(rules);
}

/**
 * ダッシュボードシートの条件付き書式を設定する補助関数
 */
function applyColorConditionRulesForDash(sheet, totalMonths) {
  const rules = [];
  const range = sheet.getRange(4, 2, totalMonths, 10); // B列〜K列
  
  // 10名以上：赤色
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(10)
    .setFontColor('#d32f2f')
    .setBackground('#fce8e6')
    .setBold(true)
    .setRanges([range])
    .build());
    
  // 2. 9名：緑色
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(9)
    .setFontColor('#137333')
    .setBackground('#e6f4ea')
    .setBold(true)
    .setRanges([range])
    .build());
    
  // 3. 7〜8名：黄色
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(7)
    .setFontColor('#b06000')
    .setBackground('#fef7e0')
    .setBold(true)
    .setRanges([range])
    .build());

  sheet.setConditionalFormatRules(rules);
}