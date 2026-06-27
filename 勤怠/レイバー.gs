/**
 * レイバースケジュール作成スクリプト（ひばりが丘・経堂 統合版）
 * ★完全安定版：スタッフごとのメモ行位置を固定指定（中田さんは1行下/経堂は2行下）＋ 日付行を3行目に完全適応
 * ★エラー完全排除：古いシートを完全に削除して新品を作り直すことで結合エラーを100%防止
 * ★表示最適化：レッスン表の名前翻訳反映 ＆ タスク(メモ)の重複スタッフ自動集約機能を搭載
 * ★競合排除フィルター：時間が被っている予定の結合エラーを自動で防ぐロジックを追加
 * ★カットオフ処理：データ出力後、下部および右側の不要な空白行列を自動削除
 * ★店舗自動振り分け：勤務地（経堂・ひばり）の記載がある場合、対象店舗のレイバーにのみシフト・メモを反映
 * ★枠線デザイン統一：勤務時間の外枠が消える問題を解消し、統一された実線で綺麗に描画されるよう最適化
 * ★対応不可時間の黒塗り：スタジオレッスンを担当している時間は、スタッフ行のシフトバーを真っ黒にして視認性を向上
 * ★メモ欄＆バーテキスト改善：メモを横並び＆小文字でスッキリ収め、抽出したタスクのバーテキストを「レッスン」に統一
 * ★最終レイアウト最適化：時間ヘッダーを中央揃えに戻し、A4印刷にフィットするよう全列の幅をバランス良く拡張
 * ★テキストサイズ極め調整：レッスン行(★/溶岩浴)を「10の太字」、バー内テキストを「8」、時間列幅をさらに1cm分拡張して完璧にフィット
 * ★機能追加(1)：前月レイバーシートが存在する場合、書き込まれたTo Doリストを曜日で紐付けて自動引き継ぎ
 * ★機能追加(2)：手書き用や急遽の追加用に、各店舗の最後に「ヘルプ」行（白の空欄）を追加
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // 1. カレンダー＆勤怠用のメニュー
  ui.createMenu('📅 勤怠カレンダー')
    .addItem('1. シフトデータ作成 (バイバイ用)', 'showDialog')
    .addItem('2. カレンダーに予定を同期', 'showCalendarSyncDialog')
    .addToUi();

  // 2. レイバー作成用のメニュー
  ui.createMenu('⚙️ レイバー作成')
    .addItem('ひばりが丘分を作成', 'showLaborDialog')
    .addItem('経堂分を作成', 'showLaborDialogKyodo')
    .addSeparator()
    .addItem('経堂：齋木だけ追記（既存維持）', 'showAppendSaikiLaborDialog')
    .addToUi();
}

/** 経堂：齋木 豪（シフト表B20行目・社員304944） */
const SAIKI_KYODO_STAFF = {
  name: "齋木",
  shiftSheetRow: 20,
  memoOffset: 1,
  color: "#cfe2f3"
};

/** 齋木さんの表記ゆれ（シフト表は「齋木」、手入力は「斎木」等） */
function isSaikiStaffName_(name) {
  const n = String(name).trim();
  return n === "齋木" || n === "斎木" || n === "齊木" ||
    n.indexOf("齋") !== -1 || n.indexOf("斎") !== -1 || n.indexOf("齊") !== -1;
}

/** 経堂レイバーの列数・時間設定 */
function getKyodoLaborLayout_() {
  const START_HOUR = 8;
  const END_HOUR = 22;
  const HOURS_COUNT = END_HOUR - START_HOUR + 1;
  const TOTAL_TIME_COLS = HOURS_COUNT * 4;
  return {
    START_HOUR: START_HOUR,
    END_HOUR: END_HOUR,
    HOURS_COUNT: HOURS_COUNT,
    TOTAL_TIME_COLS: TOTAL_TIME_COLS,
    TOTAL_COLS: 2 + TOTAL_TIME_COLS
  };
}

/** 範囲に触れる結合セルをすべて完全解除（部分解除エラー防止） */
function safeBreakApartMergedRanges_(range) {
  try {
    const merged = range.getMergedRanges();
    if (merged.length > 0) {
      const seen = {};
      merged.forEach(function (m) {
        const key = m.getA1Notation();
        if (seen[key]) return;
        seen[key] = true;
        try { m.breakApart(); } catch (e) { /* ignore */ }
      });
    } else {
      try { range.breakApart(); } catch (e) { /* ignore */ }
    }
  } catch (e) { /* ignore */ }
}

/** 結合セルを安全に解除してから再結合 */
function safeMergeRange_(range) {
  try {
    safeBreakApartMergedRanges_(range);
    range.merge();
  } catch (e) { /* ignore */ }
}

/**
 * シート選択ダイアログを表示 (ひばりが丘用)
 */
function showLaborDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let sheetNames = [];
  
  sheets.forEach(sheet => {
    if (sheet.getName().includes("シフト")) {
      sheetNames.push(sheet.getName());
    }
  });

  if (sheetNames.length === 0) {
    SpreadsheetApp.getUi().alert("「シフト」と名のつくシートが見つかりませんでした。");
    return;
  }

  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: sans-serif; padding: 10px;">
      <p style="font-weight:bold; color:#2e7d32;">【ひばりが丘】対象のシフト月を選択:</p>
      <form id="sheetForm">
        <select id="sheetName" style="width: 100%; padding: 10px; margin-bottom: 15px; border-radius:4px; border:1px solid #ccc; font-size:14px;">
          ${sheetNames.map(name => `<option value="${name}">${name}</option>`).join('')}
        </select>
        <br>
        <button id="submitBtn" type="button" onclick="submitForm()" style="width:100%; background-color: #2e7d32; color: white; border: none; padding: 12px 20px; border-radius: 4px; cursor: pointer; font-weight:bold; font-size:14px;">ひばりが丘レイバー作成</button>
      </form>
      <div id="statusMessage" style="margin-top:15px; font-size:12px; color:#666; display:none; background:#e8f5e9; padding:10px; border-radius:4px;">
        <strong style="color:#2e7d32;">● 処理を実行中...</strong><br>
        1. レッスンスケジュール同期（結合セル・重複排除）<br>
        2. スタッフシフト反映<br>
        3. To Doリスト配置<br>
        完了するまで閉じないでください。
      </div>
      
      <script>
        function submitForm() {
          const sheetName = document.getElementById('sheetName').value;
          const btn = document.getElementById('submitBtn');
          const msg = document.getElementById('statusMessage');

          btn.disabled = true;
          btn.innerText = "作成中...";
          btn.style.backgroundColor = "#cccccc";
          msg.style.display = "block";

          google.script.run
            .withSuccessHandler(function() {
              google.script.host.close();
            })
            .withFailureHandler(function(e){
                alert("エラーが発生しました: " + e.message);
                btn.disabled = false;
                btn.innerText = "ひばりが丘レイバー作成";
                btn.style.backgroundColor = "#2e7d32";
            })
            .generateLaborSchedule(sheetName);
        }
      </script>
    </div>
  `)
  .setWidth(350)
  .setHeight(280);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'レイバースケジュール作成');
}

/**
 * メイン処理 (ひばりが丘用)
 */
function generateLaborSchedule(sourceSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName(sourceSheetName);
  const todoSheet = ss.getSheetByName("To Do List");
  const lessonSheet = ss.getSheetByName("レッスンスケジュール");
  
  if (!todoSheet) throw new Error("「To Do List」シートが見つかりません。");
  if (!lessonSheet) throw new Error("「レッスンスケジュール」シートが見つかりません。");
  
  let targetSheetName = "レイバースケジュール(ひばりが丘)";
  const monthMatch = sourceSheetName.match(/(\d{1,2})月/);
  
  // ★機能追加：処理する月を判定し、前月のレイバースケジュールを特定するための準備
  let currentMonth = null;
  if (monthMatch) {
    currentMonth = parseInt(monthMatch[1], 10);
    targetSheetName = `${currentMonth}月 レイバースケジュール(ひばりが丘)`;
  }

  // ★機能追加：前月のレイバースケジュールから To Do リストを引き継ぐ（抽出する）
  let prevTodoMap = {};
  if (currentMonth !== null) {
      let prevMonth = currentMonth === 1 ? 12 : currentMonth - 1;
      let prevSheetName = `${prevMonth}月 レイバースケジュール(ひばりが丘)`;
      let prevSheet = ss.getSheetByName(prevSheetName);
      
      if (prevSheet) {
          const prevData = prevSheet.getDataRange().getValues();
          let currentDow = "";
          for (let r = 0; r < prevData.length; r++) {
              let valA = String(prevData[r][0] || "").trim();
              if (["月", "火", "水", "木", "金", "土", "日"].includes(valA)) {
                  currentDow = valA;
              }
              // "To Do リスト" 行を見つけたら、その右側のデータを曜日ごとに保持
              if (valA === "To Do\nリスト" && currentDow) {
                  if (!prevTodoMap[currentDow]) {
                      prevTodoMap[currentDow] = prevData[r].slice(1); // B列以降のすべての配列
                  }
              }
          }
      }
  }

  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    const sheetIndex = targetSheet.getIndex();
    ss.deleteSheet(targetSheet);
    SpreadsheetApp.flush(); 
    targetSheet = ss.insertSheet(targetSheetName, sheetIndex - 1);
  } else {
    targetSheet = ss.insertSheet(targetSheetName);
  }
  
  targetSheet.setTabColor("#b6d7a8");
  targetSheet.setHiddenGridlines(true);

  ss.toast("データを読み込み中...", "ひばりが丘処理");

  // --- 設定 ---
  const START_HOUR = 8; 
  const END_HOUR = 22;  
  const HOURS_COUNT = END_HOUR - START_HOUR + 1; 
  const TOTAL_TIME_COLS = HOURS_COUNT * 4; 
  const TOTAL_COLS = 4 + TOTAL_TIME_COLS; 

  const HEIGHT_HEADER = 45; 
  const HEIGHT_LESSON = 40; 
  const HEIGHT_ROW = 38;    
  const HEIGHT_TODO_FIXED = 350; 
  const HEIGHT_MEMO = 100;   
  const HEIGHT_SPACER = 50; 

  const COLOR = {
    HEADER: "#f3f3f3",     
    LESSON_ROW_DEFAULT: "#d9ead3", 
    SHIFT_BAR: "#cccccc",  
    TODO_LABEL: "#b6d7a8", 
    TODO_CONTENT_BG: "#efefef", 
    MEMO_LABEL: "#d9ead3",
    DATE_BG: "#d9ead3",   
    TEXT_SAT: "#1155cc",       
    TEXT_SUN: "#ff0000",       
    WHITE: "#ffffff",
    BLACK: "#000000", 
    STRIPE_GRAY: "#f7f7f7",
    STAFF_COLORS: {
      "sayumi": "#fce5cd", "midori": "#d9ead3", "hana": "#ead1dc",
      "kaori": "#cfe2f3", "masayo": "#cfe2f3", "mito": "#d9d2e9"
    }
  };

  const targetStaffs = [
    { name: "sayumi", color: COLOR.STAFF_COLORS["sayumi"], memoOffset: 1 },
    { name: "midori", color: COLOR.STAFF_COLORS["midori"], memoOffset: 1 },
    { name: "hana",   color: COLOR.STAFF_COLORS["hana"], memoOffset: 1 }, 
    { name: "kaori",  color: COLOR.STAFF_COLORS["kaori"], memoOffset: 1 },
    { name: "masayo", color: COLOR.STAFF_COLORS["masayo"], memoOffset: 1 }, 
    { name: "mito",   color: COLOR.STAFF_COLORS["mito"], memoOffset: 1 },
    { name: "ヘルプ", color: COLOR.WHITE, memoOffset: null } 
  ];

  let staffList = [];
  const nameColData = srcSheet.getRange("B1:B100").getValues(); 
  let nameToRowIndex = {};

  for(let r=0; r<nameColData.length; r++){
    let val = String(nameColData[r][0]).trim();
    if(val) {
        nameToRowIndex[val] = r + 1;
    }
  }
  
  targetStaffs.forEach(staff => {
    staffList.push({ 
      name: staff.name, 
      color: staff.color, 
      rowIndex: nameToRowIndex[staff.name] || null,
      memoOffset: staff.memoOffset
    });
  });

  const lastColSrc = srcSheet.getLastColumn();
  
  const DATE_ROW = 3; 
  const DOW_ROW = 4;  

  const daysCount = lastColSrc - 2;
  if (daysCount < 1) throw new Error("日付データが見つかりません。シフト表の形式を確認してください。");

  const dates = srcSheet.getRange(DATE_ROW, 3, 1, daysCount).getValues()[0];
  const dows = srcSheet.getRange(DOW_ROW, 3, 1, daysCount).getValues()[0]; 
  const shiftDataRange = srcSheet.getRange(1, 3, 100 + staffList.length * 3, daysCount).getValues(); 
  
  const todoDataRaw = todoSheet.getRange("A1:C100").getValues(); 
  let todoMap = {};
  todoDataRaw.forEach(row => {
    let contentRaw = String(row[0] || "");
    let dayKey = String(row[1] || "").trim().charAt(0);
    let pos = String(row[2] || "").trim();
    if (!["1", "2", "3"].includes(pos)) pos = "auto";
    if (dayKey && contentRaw) {
      if(!todoMap[dayKey]) todoMap[dayKey] = { "1": [], "2": [], "3": [], "auto": [] };
      contentRaw.split(/\n/).forEach(item => {
        let clean = item.replace(/[□▢■]/g, "").trim();
        if (clean) todoMap[dayKey][pos].push(clean);
      });
    }
  });

  const lessonSourceRange = lessonSheet.getRange(2, 1, 7, 1 + TOTAL_TIME_COLS);
  const lessonSourceValues = lessonSourceRange.getValues();
  const lessonSourceBackgrounds = lessonSourceRange.getBackgrounds();
  const lessonMergedRanges = lessonSourceRange.getMergedRanges();

  let lessonMap = {};
  const dowList = ["月", "火", "水", "木", "金", "土", "日"];

  for (let r = 0; r < 7; r++) {
    let dowLabel = String(lessonSourceValues[r][0]).charAt(0);
    if (dowList.includes(dowLabel)) {
      let rowMerges = [];
      lessonMergedRanges.forEach(range => {
        if (range.getRow() - 2 === r) {
          rowMerges.push({
            startColIndex: range.getColumn() - 2,
            numCols: range.getNumColumns()
          });
        }
      });

      lessonMap[dowLabel] = {
        vals: lessonSourceValues[r].slice(1), 
        bgs: lessonSourceBackgrounds[r].slice(1),
        merges: rowMerges
      };
    }
  }

  let values = [], backgrounds = [], fontColors = [], fontSizes = [], fontWeights = [], hAligns = [], vAligns = [], wraps = [];
  let mergeRanges = [], solidAllRanges = [], timeGridRanges = [], hourMarkerRanges = [], importantBorderRanges = [], summaryBorderRanges = [], shiftBorderRanges = [];
  let rowHeightRequests = []; 

  function createNewRow() {
    return {
      val: new Array(TOTAL_COLS).fill(""), bg: new Array(TOTAL_COLS).fill(COLOR.WHITE), fc: new Array(TOTAL_COLS).fill("black"), 
      fs: new Array(TOTAL_COLS).fill(10), fw: new Array(TOTAL_COLS).fill("normal"), ha: new Array(TOTAL_COLS).fill("center"),
      va: new Array(TOTAL_COLS).fill("middle"), wp: new Array(TOTAL_COLS).fill(false)
    };
  }

  let currentRowIdx = 0; 

  for (let i = 0; i < dates.length; i++) {
    let date = dates[i];
    if (date === "" || date === null || date === undefined) continue; 
    
    let displayDow = String(dows[i] || "").trim().charAt(0);
    
    let dateLabel = "";
    let dowFull = displayDow;

    if (date instanceof Date) {
      dateLabel = Utilities.formatDate(date, "JST", "d") + "日";
      dowFull = ["日","月","火","水","木","金","土"][date.getDay()];
    } else {
      dateLabel = String(date).trim();
      if (/^[0-9０-９]+$/.test(dateLabel)) {
        dateLabel += "日";
      }
    }

    let startBlockRowIdx = currentRowIdx; 

    let rowH = createNewRow();
    rowH.val[0] = dowFull; rowH.fs[0] = 16; rowH.fw[0] = "bold"; 
    if (dowFull === "土") rowH.fc[0] = COLOR.TEXT_SAT;
    else if (dowFull === "日") rowH.fc[0] = COLOR.TEXT_SUN;
    rowH.val[2] = "開始"; rowH.val[3] = "終了";
    for (let h=0; h<HOURS_COUNT; h++) {
      let cIdx = 4 + (h*4);
      rowH.val[cIdx] = (START_HOUR+h)+":00"; 
      rowH.fs[cIdx] = 10; 
      rowH.fw[cIdx] = "bold";
      rowH.ha[cIdx] = "center"; 
      for(let q=0; q<4; q++) rowH.bg[cIdx+q] = COLOR.HEADER;
      mergeRanges.push({ row: currentRowIdx+1, col: cIdx+1, numRow: 1, numCol: 4 });
    }
    values.push(rowH.val); backgrounds.push(rowH.bg); fontColors.push(rowH.fc); fontSizes.push(rowH.fs); fontWeights.push(rowH.fw); hAligns.push(rowH.ha); vAligns.push(rowH.va); wraps.push(rowH.wp);
    rowHeightRequests.push({ row: currentRowIdx+1, height: HEIGHT_HEADER });
    currentRowIdx++;

    let rowL = createNewRow();
    rowL.val[1] = "レッスン"; rowL.fw[1] = "bold"; rowL.fs[1] = 10;
    rowL.val[0] = dateLabel; rowL.fw[0] = "bold";
    rowL.bg[1] = COLOR.WHITE; backgrounds[currentRowIdx-1][0] = COLOR.DATE_BG;

    if (lessonMap[dowFull]) {
      let lData = lessonMap[dowFull];
      for (let c = 0; c < TOTAL_TIME_COLS; c++) {
        rowL.val[4 + c] = String(lData.vals[c] || "");
        rowL.bg[4 + c] = (lData.bgs[c] !== "#ffffff" && lData.bgs[c] !== "white") ? lData.bgs[c] : COLOR.LESSON_ROW_DEFAULT;
        rowL.fs[4 + c] = 10;
        rowL.fw[4 + c] = "bold";
      }
      
      lData.merges.forEach(m => {
        if (m.startColIndex >= 0 && m.startColIndex < TOTAL_TIME_COLS) {
          for (let k = 1; k < m.numCols; k++) {
            let targetIdx = 4 + m.startColIndex + k;
            if (targetIdx < TOTAL_COLS) rowL.val[targetIdx] = "";
          }
          mergeRanges.push({ 
            row: currentRowIdx + 1, 
            col: 5 + m.startColIndex, 
            numRow: 1, 
            numCol: m.numCols 
          });
        }
      });
    } else {
      for(let t=0; t<TOTAL_TIME_COLS; t++) rowL.bg[4+t] = COLOR.LESSON_ROW_DEFAULT;
    }
    
    values.push(rowL.val); backgrounds.push(rowL.bg); fontColors.push(rowL.fc); fontSizes.push(rowL.fs); fontWeights.push(rowL.fw); hAligns.push(rowL.ha); vAligns.push(rowL.va); wraps.push(rowL.wp);
    importantBorderRanges.push({ row: currentRowIdx + 1, col: 5, numRow: 1, numCol: TOTAL_TIME_COLS });
    rowHeightRequests.push({ row: currentRowIdx+1, height: HEIGHT_LESSON });
    currentRowIdx++;

    let dailyNotesMap = {}; 

    staffList.forEach(staff => {
      let rowS = createNewRow();
      
      let myColor = staff.color || COLOR.WHITE;

      let shiftText = "";
      let noteText = "";

      if (staff.rowIndex) {
        shiftText = String(shiftDataRange[staff.rowIndex-1][i] || "").trim();
        if (staff.memoOffset) {
           noteText = String(shiftDataRange[staff.rowIndex - 1 + staff.memoOffset][i] || "").trim();
        }
      }

      // メモ行やシフト行に「経堂」等の記載があれば他店舗扱い
      let isOtherShop = false;
      let combinedText = shiftText + " " + noteText;
      if (combinedText.includes("経堂")) {
          isOtherShop = true;
      } else if (combinedText.includes("ひばり")) {
          myColor = "#d9ead3"; 
      }

      rowS.val[1] = staff.name; 
      rowS.bg[1] = isOtherShop ? COLOR.STRIPE_GRAY : myColor; 
      rowS.fw[1] = "bold";
      rowS.fs[1] = 10;
      if (isOtherShop) rowS.fc[1] = "#999999"; 

      for (let t = 4; t < TOTAL_COLS; t++) {
        rowS.fs[t] = 10;
        rowS.fw[t] = "bold"; 
      }

      if (staff.rowIndex && !isOtherShop) {
        let shiftStartMin = -1;
        let shiftEndMin = -1;

        if (shiftText && /\d/.test(shiftText)) {
          let clean = shiftText.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0)-0xFEE0)).replace(/\s/g, "");
          let m = clean.match(/^(\d{1,2})[:：]?(\d{2})?[-~～](\d{1,2})[:：]?(\d{2})?/);
          if (m) {
            let sH = parseInt(m[1]), sM = parseInt(m[2] || "00"), eH = parseInt(m[3]), eM = parseInt(m[4] || "00");
            rowS.val[2] = sH + ":" + (m[2] || "00"); rowS.val[3] = eH + ":" + (m[4] || "00");
            
            rowS.fs[2] = 10; 
            rowS.fs[3] = 10; 

            shiftStartMin = sH * 60 + sM;
            shiftEndMin = eH * 60 + eM;

            let sB = (sH-START_HOUR)*4 + Math.floor(sM/15), eB = (eH-START_HOUR)*4 + Math.floor(eM/15);
            for(let k=sB; k<eB && (4+k)<TOTAL_COLS; k++) rowS.bg[4+k] = COLOR.SHIFT_BAR;
          }
        }

        if (noteText) {
          if (!dailyNotesMap[noteText]) {
             dailyNotesMap[noteText] = [];
          }
          dailyNotesMap[noteText].push(staff.name);

          let timeMatch = noteText.match(/(\d{1,2})[:：]?(\d{2})?[-~～](\d{1,2})[:：]?(\d{2})?/);
          if (timeMatch) {
            let nStartH = parseInt(timeMatch[1]);
            let nStartM = parseInt(timeMatch[2] || "00");
            let nEndH = parseInt(timeMatch[3]);
            let nEndM = parseInt(timeMatch[4] || "00");
            
            let nStartMin = nStartH * 60 + nStartM;
            let nEndMin = nEndH * 60 + nEndM;

            let durationMin = nEndMin - nStartMin;
            let isSameAsShift = (nStartMin === shiftStartMin && nEndMin === shiftEndMin);

            // シフトと同じ時間、または4時間(240分)以上の長大なタスクはバーを描画しない
            if (!isSameAsShift && durationMin > 0 && durationMin < 240) {
                let nStartLimit = START_HOUR * 60;
                if (nEndMin > nStartLimit && nStartMin < (END_HOUR + 1) * 60) {
                   let sB_N = Math.floor((Math.max(nStartMin, nStartLimit) - nStartLimit) / 15);
                   let eB_N = Math.ceil((Math.min(nEndMin, (END_HOUR+1)*60) - nStartLimit) / 15);
                   
                   let startColIdxN = 4 + sB_N; 
                   let durationN = eB_N - sB_N;
                   
                   if (durationN > 0) {
                       // ★修正：タスクの時間は「黒塗り」のみにする。文字は入れない。
                       for(let k=0; k<durationN && (startColIdxN+k)<TOTAL_COLS; k++) {
                           rowS.bg[startColIdxN+k] = COLOR.BLACK;
                       }
                       
                       shiftBorderRanges.push({ 
                           row: currentRowIdx + 1, 
                           col: startColIdxN + 1, 
                           numRow: 1, 
                           numCol: durationN 
                       });
                   }
                }
            }
          }
        }

        if (shiftStartMin >= START_HOUR * 60 && shiftEndMin > shiftStartMin) {
          let sB = Math.floor((shiftStartMin - START_HOUR * 60) / 15);
          let eB = Math.floor((shiftEndMin - START_HOUR * 60) / 15);
          
          let startColIdx = 4 + sB; 
          let duration = eB - sB;
          
          if (duration > 0) {
              shiftBorderRanges.push({ 
                  row: currentRowIdx + 1, 
                  col: startColIdx + 1, 
                  numRow: 1, 
                  numCol: duration 
              });
          }
        }

        if (lessonMap[dowFull]) {
          let lData = lessonMap[dowFull];
          let themeColor = myColor.toLowerCase(); 
          
          if (themeColor && themeColor !== "#ffffff") {
            for (let c = 0; c < TOTAL_TIME_COLS; c++) {
              let lessonBg = (lData.bgs[c] || "").toLowerCase();
              if (lessonBg === themeColor && rowS.bg[4 + c] === COLOR.SHIFT_BAR) {
                rowS.bg[4 + c] = COLOR.BLACK; 
                rowS.val[4 + c] = "";
              }
            }
            
            lData.merges.forEach(m => {
              if (m.startColIndex >= 0 && m.startColIndex < TOTAL_TIME_COLS) {
                let targetBg = (lData.bgs[m.startColIndex] || "").toLowerCase();
                if (targetBg === themeColor) {
                  if (rowS.bg[4 + m.startColIndex] === COLOR.BLACK) {
                    rowS.val[4 + m.startColIndex] = ""; 
                    mergeRanges.push({ 
                      row: currentRowIdx + 1, 
                      col: 5 + m.startColIndex, 
                      numRow: 1, 
                      numCol: m.numCols 
                    });
                  }
                }
              }
            });
          }
        }
      }
      values.push(rowS.val); backgrounds.push(rowS.bg); fontColors.push(rowS.fc); fontSizes.push(rowS.fs); fontWeights.push(rowS.fw); hAligns.push(rowS.ha); vAligns.push(rowS.va); wraps.push(rowS.wp);
      rowHeightRequests.push({ row: currentRowIdx+1, height: HEIGHT_ROW });
      currentRowIdx++;
    });

    mergeRanges.push({ row: startBlockRowIdx+2, col: 1, numRow: 1 + staffList.length, numCol: 1 });

    let dayTodo = todoMap[dowFull] || { "1":[], "2":[], "3":[], "auto":[] };
    let c1 = [...dayTodo["1"]], c2 = [...dayTodo["2"]], c3 = [...dayTodo["3"]];
    dayTodo["auto"].forEach(item => {
      if (c1.length <= c2.length && c1.length <= c3.length) c1.push(item);
      else if (c2.length <= c1.length && c2.length <= c3.length) c2.push(item);
      else c3.push(item);
    });
    let rowT = createNewRow();
    rowT.val[0] = "To Do\nリスト"; rowT.bg[0] = COLOR.TODO_LABEL; rowT.fw[0] = "bold";
    for(let k=1; k<TOTAL_COLS; k++) rowT.bg[k] = COLOR.TODO_CONTENT_BG; 
    let w = Math.floor((TOTAL_COLS-1)/3);
    
    // ★機能追加：前月のレイバースケジュールにその曜日の To Do があれば、それを優先して引き継ぐ
    if (prevTodoMap[dowFull]) {
        let prevCols = prevTodoMap[dowFull];
        rowT.val[1] = String(prevCols[0] || "");
        rowT.val[1+w] = String(prevCols[w] || "");
        rowT.val[1+w*2] = String(prevCols[w*2] || "");
    } else {
        // もし前月のシートがない、または書き込みがなければ今まで通り「To Do List」シートから取得
        rowT.val[1] = c1.map(it => "□ " + it).join("\n");
        rowT.val[1+w] = c2.map(it => "□ " + it).join("\n");
        rowT.val[1+w*2] = c3.map(it => "□ " + it).join("\n");
    }

    [1, 1+w, 1+w*2].forEach(idx => { rowT.ha[idx] = "left"; rowT.va[idx] = "top"; rowT.wp[idx] = true; rowT.fs[idx] = 10; });
    values.push(rowT.val); backgrounds.push(rowT.bg); fontColors.push(rowT.fc); fontSizes.push(rowT.fs); fontWeights.push(rowT.fw); hAligns.push(rowT.ha); vAligns.push(rowT.va); wraps.push(rowT.wp);
    mergeRanges.push({ row: currentRowIdx+1, col: 2, numRow: 1, numCol: w });
    mergeRanges.push({ row: currentRowIdx+1, col: 2+w, numRow: 1, numCol: w });
    mergeRanges.push({ row: currentRowIdx+1, col: 2+w*2, numRow: 1, numCol: TOTAL_COLS-(1+w*2) });
    summaryBorderRanges.push({ row: currentRowIdx+1, col: 1, numRow: 1, numCol: TOTAL_COLS });
    rowHeightRequests.push({ row: currentRowIdx+1, height: HEIGHT_TODO_FIXED });
    currentRowIdx++;

    let dailyNotesStrList = [];
    for (let note in dailyNotesMap) {
       let staffsStr = dailyNotesMap[note].join("・");
       dailyNotesStrList.push(`●${staffsStr}：${note}`);
    }
    let memoContent = dailyNotesStrList.join("　　"); 

    let rowM = createNewRow();
    rowM.val[0] = "メモ"; rowM.bg[0] = COLOR.MEMO_LABEL; rowM.fw[0] = "bold";
    rowM.val[1] = memoContent; 
    rowM.fs[1] = 9; 
    rowM.ha[1] = "left"; rowM.va[1] = "top"; rowM.wp[1] = true;
    values.push(rowM.val); backgrounds.push(rowM.bg); fontColors.push(rowM.fc); fontSizes.push(rowM.fs); fontWeights.push(rowM.fw); hAligns.push(rowM.ha); vAligns.push(rowM.va); wraps.push(rowM.wp);
    mergeRanges.push({ row: currentRowIdx+1, col: 2, numRow: 1, numCol: TOTAL_COLS-1 });
    summaryBorderRanges.push({ row: currentRowIdx+1, col: 1, numRow: 1, numCol: TOTAL_COLS });
    rowHeightRequests.push({ row: currentRowIdx+1, height: HEIGHT_MEMO });
    currentRowIdx++;

    solidAllRanges.push({ row: startBlockRowIdx+1, col: 1, numRow: currentRowIdx-startBlockRowIdx, numCol: TOTAL_COLS });
    timeGridRanges.push({ row: startBlockRowIdx+1, col: 5, numRow: staffList.length+2, numCol: TOTAL_TIME_COLS });
    for (let h=1; h<HOURS_COUNT; h++) hourMarkerRanges.push({ row: startBlockRowIdx+1, col: 5+(h*4), numRow: staffList.length+2, numCol: 1 });

    values.push(new Array(TOTAL_COLS).fill("")); backgrounds.push(new Array(TOTAL_COLS).fill(COLOR.WHITE));
    fontColors.push(new Array(TOTAL_COLS).fill("black")); fontSizes.push(new Array(TOTAL_COLS).fill(10)); fontWeights.push(new Array(TOTAL_COLS).fill("normal"));
    hAligns.push(new Array(TOTAL_COLS).fill("center")); vAligns.push(new Array(TOTAL_COLS).fill("middle")); wraps.push(new Array(TOTAL_COLS).fill(false));
    rowHeightRequests.push({ row: currentRowIdx+1, height: HEIGHT_SPACER });
    currentRowIdx++;
  }

  ss.toast("シートを整形中...", "最終段階", 10);
  const totalRows = values.length;
  
  if (totalRows === 0) throw new Error("シフト表から日付データを読み取れませんでした。3行目に日付があるか確認してください。");

  function removeOverlaps(ranges) {
    let result = [];
    for (let i = ranges.length - 1; i >= 0; i--) {
      let current = ranges[i];
      let isOverlap = result.some(r => {
        let rowOverlap = (current.row < r.row + r.numRow) && (current.row + current.numRow > r.row);
        let colOverlap = (current.col < r.col + r.numCol) && (current.col + current.numCol > r.col);
        return rowOverlap && colOverlap;
      });
      if (!isOverlap) result.push(current);
    }
    return result;
  }

  mergeRanges = removeOverlaps(mergeRanges);

  const range = targetSheet.getRange(1, 1, totalRows, TOTAL_COLS);
  range.setValues(values).setBackgrounds(backgrounds).setFontColors(fontColors).setFontSizes(fontSizes).setFontWeights(fontWeights)
       .setHorizontalAlignments(hAligns).setVerticalAlignments(vAligns).setWraps(wraps).setFontFamily("Meiryo");

  targetSheet.setColumnWidth(1, 50); 
  targetSheet.setColumnWidth(2, 75); 
  targetSheet.setColumnWidth(3, 42); 
  targetSheet.setColumnWidth(4, 42); 
  for (let c=5; c<=TOTAL_COLS; c++) targetSheet.setColumnWidth(c, 17); 

  rowHeightRequests.forEach(req => { if(req.row <= totalRows) targetSheet.setRowHeight(req.row, req.height); });

  solidAllRanges.forEach(r => {
      try { targetSheet.getRange(r.row, r.col, r.numRow, r.numCol).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID); } catch(e){}
  });
  
  timeGridRanges.forEach(r => {
      try { targetSheet.getRange(r.row, r.col, r.numRow, r.numCol).setBorder(null, null, null, null, true, null, "black", SpreadsheetApp.BorderStyle.DOTTED); } catch(e){}
  });
  
  hourMarkerRanges.forEach(r => {
      try { targetSheet.getRange(r.row, r.col, r.numRow, r.numCol).setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID); } catch(e){}
  });

  shiftBorderRanges.forEach(r => {
      if (r.numCol > 0) {
          try {
              targetSheet.getRange(r.row, r.col, r.numRow, r.numCol)
                  .setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
          } catch(e){}
      }
  });

  summaryBorderRanges.forEach(r => {
      try { targetSheet.getRange(r.row, r.col, r.numRow, r.numCol).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM); } catch(e){}
  });

  mergeRanges.forEach(m => {
    try {
      if (m.numRow === 1 && m.numCol === 1) return;
      targetSheet.getRange(m.row, m.col, m.numRow, m.numCol).merge();
    } catch(e) {}
  });

  const maxRows = targetSheet.getMaxRows();
  if (maxRows > totalRows) {
    targetSheet.deleteRows(totalRows + 1, maxRows - totalRows);
  }
  const maxCols = targetSheet.getMaxColumns();
  if (maxCols > TOTAL_COLS) {
    targetSheet.deleteColumns(TOTAL_COLS + 1, maxCols - TOTAL_COLS);
  }

  ss.toast("ひばりが丘レイバー作成完了！", "成功", 3);
  
  const ui = SpreadsheetApp.getUi();
  ui.alert("作成完了！\nレッスン枠の結合とテキストの重複排除が完了しました。");
}

/**
 * シート選択ダイアログを表示 (経堂用)
 */
function showLaborDialogKyodo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let sheetNames = [];
  
  sheets.forEach(sheet => {
    if (sheet.getName().includes("シフト")) {
      sheetNames.push(sheet.getName());
    }
  });

  if (sheetNames.length === 0) {
    SpreadsheetApp.getUi().alert("「シフト」と名のつくシートが見つかりませんでした。");
    return;
  }

  const htmlOutput = HtmlService.createHtmlOutput(`
    <div style="font-family: sans-serif; padding: 10px;">
      <p style="font-weight:bold; color:#b71c1c;">【経堂】対象のシフト月を選択:</p>
      <form id="sheetForm">
        <select id="sheetName" style="width: 100%; padding: 10px; margin-bottom: 15px; border-radius:4px; border:1px solid #ccc; font-size:14px;">
          ${sheetNames.map(name => `<option value="${name}">${name}</option>`).join('')}
        </select>
        <br>
        <button id="submitBtn" type="button" onclick="submitForm()" style="width:100%; background-color: #b71c1c; color: white; border: none; padding: 12px 20px; border-radius: 4px; cursor: pointer; font-weight:bold; font-size:14px;">経堂レイバー作成</button>
      </form>
      <div id="statusMessage" style="margin-top:15px; font-size:12px; color:#666; display:none; background:#fbe9e7; padding:10px; border-radius:4px;">
        <strong style="color:#b71c1c;">● 処理を実行中...</strong><br>
        メモ欄を統合して作成中。<br>
        完了するまで閉じないでください。
      </div>
      
      <script>
        function submitForm() {
          const sheetName = document.getElementById('sheetName').value;
          const btn = document.getElementById('submitBtn');
          const msg = document.getElementById('statusMessage');

          btn.disabled = true;
          btn.innerText = "作成中...";
          btn.style.backgroundColor = "#cccccc";
          msg.style.display = "block";

          google.script.run
            .withSuccessHandler(function() {
              google.script.host.close();
            })
            .withFailureHandler(function(e){
                alert("エラーが発生しました: " + e.message);
                btn.disabled = false;
                btn.innerText = "経堂レイバー作成";
                btn.style.backgroundColor = "#b71c1c";
            })
            .generateLaborScheduleKyodo(sheetName);
        }
      </script>
    </div>
  `)
  .setWidth(350)
  .setHeight(260);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'レイバースケジュール作成(経堂)');
}

/**
 * メイン処理：レイバースケジュール生成（経堂用）
 */
function generateLaborScheduleKyodo(sourceSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName(sourceSheetName);
  
  const pgSheet = ss.getSheetByName("レッスン表経堂");
  if (!pgSheet) throw new Error("「レッスン表経堂」シートが見つかりません。");
  
  let targetSheetName = "レイバースケジュール(経堂)"; 
  const monthMatch = sourceSheetName.match(/(\d{1,2})月/);
  if (monthMatch) {
    targetSheetName = `${monthMatch[1]}月 レイバースケジュール(経堂)`;
  }
  
  let targetSheet = ss.getSheetByName(targetSheetName);
  if (targetSheet) {
    const sheetIndex = targetSheet.getIndex();
    ss.deleteSheet(targetSheet);
    SpreadsheetApp.flush(); 
    targetSheet = ss.insertSheet(targetSheetName, sheetIndex - 1);
  } else {
    targetSheet = ss.insertSheet(targetSheetName);
  }
  targetSheet.setTabColor("#ff0000"); 

  ss.toast("データを読み込み中...", "経堂用処理", 3);

  const START_HOUR = 8; 
  const END_HOUR = 22;  
  const HOURS_COUNT = END_HOUR - START_HOUR + 1; 
  const TOTAL_TIME_COLS = HOURS_COUNT * 4; 
  const TOTAL_COLS = 2 + TOTAL_TIME_COLS; 

  const HEIGHT_HEADER = 28; 
  const HEIGHT_ROW = 24;    
  const HEIGHT_BLOCK_SEP = 8; 
  const HEIGHT_SUMMARY_HEADER = 22; 
  const HEIGHT_SUMMARY_FIXED = 110; 
  const HEIGHT_SPACER = 15; 

  const COLOR = {
    HEADER: "#f3f3f3",     
    SHIFT_BAR: "#cccccc",  
    SUMMARY_HEADER: "#eeeeee",
    WHITE: "#ffffff",
    BLACK: "#000000",
    STRIPE_GRAY: "#f7f7f7",
    STAFF_ORANGE: "#fce5cd", 
    STAFF_BLUE: "#cfe2f3",   
    STAFF_PINK: "#ead1dc",   
    SUMMARY_MEMO_TITLE_BG: "#cc0000", 
    SUMMARY_MEMO_TITLE_TXT: "#ffffff",
    LESSON_HOT_BG: "#cc0000", 
    LESSON_HOT_TXT: "#ffffff", 
    LESSON_REFORMER_BG: "#d9ead3", 
    LESSON_REFORMER_TXT: "#000000", 
    ROW_STUDIO_NAME_BG: "#fce5cd" 
  };

  const targetStaffs = [
    { name: "蜂谷", color: COLOR.STAFF_ORANGE, memoOffset: 1 },
    { name: "日下", color: COLOR.STAFF_ORANGE, memoOffset: 1 },
    { name: "中田", color: COLOR.STAFF_BLUE, memoOffset: 1 },
    { name: "美絵", color: COLOR.STAFF_BLUE, memoOffset: 1 },
    { name: "由岐恵", color: COLOR.STAFF_BLUE, memoOffset: 1 },
    { name: "澤野", color: COLOR.STAFF_BLUE, memoOffset: 1 },
    { name: "みと", color: COLOR.STAFF_PINK, memoOffset: 1 },
    { name: "齋木", color: COLOR.STAFF_BLUE, memoOffset: 1 } // 20行目・社員304944
  ];

  const instructorMap = {
    "YUKI": "由岐恵",
    "Yuki": "由岐恵",
    "MIE": "美絵",
    "Mie": "美絵",
    "HACHI": "蜂谷",
    "Hachi": "蜂谷",
    "Hana": "中田",
    "HANA": "中田",
    "みと": "みと",
    "由岐恵": "由岐恵",
    "美絵": "美絵",
    "蜂谷": "蜂谷",
    "中田": "中田",
    "日下": "日下",
    "澤野": "澤野"
  };

  let staffList = [];
  const nameColData = srcSheet.getRange("B1:B100").getValues(); 
  let nameToRowIndex = {};

  for(let r=0; r<nameColData.length; r++){
    let val = String(nameColData[r][0]).trim();
    if(val){ 
        nameToRowIndex[val] = r + 1; 
    }
  }
  
  targetStaffs.forEach(staff => {
    staffList.push({ 
      name: staff.name, 
      color: staff.color,
      rowIndex: nameToRowIndex[staff.name] || null,
      memoOffset: staff.memoOffset
    });
  });

  const lastColSrc = srcSheet.getLastColumn();
  
  const DATE_ROW = 3; 
  const DOW_ROW = 4;  

  const daysCount = lastColSrc - 2; 
  if (daysCount < 1) throw new Error("日付データが見つかりません。シフト表の形式を確認してください。");

  const dates = srcSheet.getRange(DATE_ROW, 3, 1, daysCount).getValues()[0];
  const dows = srcSheet.getRange(DOW_ROW, 3, 1, daysCount).getValues()[0]; 
  const shiftDataRange = srcSheet.getRange(1, 3, 100 + staffList.length * 3, daysCount).getValues(); 

  const pgDataRaw = pgSheet.getRange("A2:H100").getValues();
  let pgAllLessons = {};
  let pgLessonMap = {}; 
  
  const timeToMinutes = (t) => {
    if (t instanceof Date) {
      return t.getHours() * 60 + t.getMinutes();
    }
    if (typeof t === 'string' && t.includes(':')) {
      const p = t.split(':');
      return parseInt(p[0]) * 60 + parseInt(p[1]);
    }
    return null;
  };

  pgDataRaw.forEach(row => {
    let dow = row[0];
    if (!dow) return;
    dow = String(dow).trim().charAt(0);

    let area = String(row[1] || "").trim(); 
    let startMin = timeToMinutes(row[2]);
    let endMin = timeToMinutes(row[3]);
    let instructorRaw = String(row[6]).trim();

    if (startMin === null || endMin === null) return;

    let mappedName = instructorMap[instructorRaw];
    if (!mappedName) {
        if (targetStaffs.some(s => s.name === instructorRaw)) {
            mappedName = instructorRaw;
        } else {
            mappedName = instructorRaw; 
        }
    }

    let lessonObj = {
        start: startMin,
        end: endMin,
        instructor: instructorRaw,
        displayName: mappedName, 
        area: area
    };

    if (!pgAllLessons[dow]) pgAllLessons[dow] = [];
    pgAllLessons[dow].push(lessonObj);

    if (mappedName) {
        if (!pgLessonMap[dow]) pgLessonMap[dow] = {};
        if (!pgLessonMap[dow][mappedName]) pgLessonMap[dow][mappedName] = [];
        
        pgLessonMap[dow][mappedName].push(lessonObj);
    }
  });

  let values = [], backgrounds = [], fontColors = [], fontSizes = [], fontWeights = [], hAligns = [], vAligns = [], wraps = [];
  let mergeRanges = [], solidAllRanges = [], timeGridRanges = [], hourMarkerRanges = [], shiftBorderRanges = [], summaryBorderRanges = [];
  let rowHeightRequests = []; 

  function createNewRow() {
    let bgRow = new Array(TOTAL_COLS).fill(COLOR.WHITE);
    let fcRow = new Array(TOTAL_COLS).fill("black"); 
    for (let t=0; t<TOTAL_TIME_COLS; t++) { if (Math.floor(t/4)%2 !== 0) bgRow[2+t] = COLOR.STRIPE_GRAY; }
    return {
      val: new Array(TOTAL_COLS).fill(""), bg: bgRow, fc: fcRow, fs: new Array(TOTAL_COLS).fill(10), 
      fw: new Array(TOTAL_COLS).fill("normal"), ha: new Array(TOTAL_COLS).fill("center"),
      va: new Array(TOTAL_COLS).fill("middle"), wp: new Array(TOTAL_COLS).fill(false)
    };
  }

  let currentRowIdx = 0; 

  for (let i = 0; i < dates.length; i++) {
    let date = dates[i];
    let displayDow = String(dows[i] || "").trim().charAt(0);
    if (!date) continue;

    let dateLabel = (date instanceof Date) ? Utilities.formatDate(date, "JST", "d") + "日" : date;
    let dowFull = (date instanceof Date) ? ["日","月","火","水","木","金","土"][date.getDay()] : displayDow;

    let startBlockRowIdx = currentRowIdx; 

    let row1 = createNewRow();
    row1.val[0] = dateLabel; 
    row1.val[1] = dowFull;   
    row1.fs[0] = 13; row1.fw[0] = "bold"; 
    row1.fs[1] = 14; row1.fw[1] = "bold";
    
    for (let h=0; h<HOURS_COUNT; h++) {
      let cIdx = 2 + (h*4);
      row1.val[cIdx] = (START_HOUR+h)+":00"; 
      row1.fs[cIdx] = 10; 
      row1.fw[cIdx] = "bold";
      row1.ha[cIdx] = "center";
      mergeRanges.push({ row: currentRowIdx+1, col: cIdx+1, numRow: 1, numCol: 4 });
    }
    values.push(row1.val); backgrounds.push(row1.bg); fontColors.push(row1.fc); fontSizes.push(row1.fs); fontWeights.push(row1.fw); hAligns.push(row1.ha); vAligns.push(row1.va); wraps.push(row1.wp);
    
    rowHeightRequests.push({ row: currentRowIdx+1, height: HEIGHT_HEADER }); 
    currentRowIdx++;

    const studioRows = [
        { label: "スタジオ", keyword: "ホット" }, 
        { label: "ピラティス", keyword: "ピラティス" } 
    ];

    studioRows.forEach(studio => {
        let rowSt = createNewRow();
        rowSt.val[1] = studio.label;
        
        if (studio.label === "スタジオ") {
             rowSt.bg[1] = COLOR.LESSON_HOT_BG; 
             rowSt.fc[1] = COLOR.LESSON_HOT_TXT; 
        } else {
             rowSt.bg[1] = COLOR.LESSON_REFORMER_BG; 
             rowSt.fc[1] = COLOR.LESSON_REFORMER_TXT; 
        }

        rowSt.fw[1] = "bold";
        rowSt.fs[1] = 10; 

        if (pgAllLessons[dowFull]) {
            pgAllLessons[dowFull].forEach(lesson => {
                if (lesson.area.includes(studio.keyword) || (studio.keyword === "ピラティス" && lesson.area.includes("リフォーマー"))) {
                    let startLimit = START_HOUR * 60;
                    if (lesson.end > startLimit && lesson.start < (END_HOUR + 1) * 60) {
                        let sB = Math.floor((Math.max(lesson.start, startLimit) - startLimit) / 15);
                        let eB = Math.ceil((Math.min(lesson.end, (END_HOUR+1)*60) - startLimit) / 15);
                        
                        let startColIdx = 2 + sB;
                        let duration = eB - sB;
                        
                        if (duration > 0) {
                            let bgColor = COLOR.LESSON_HOT_BG;
                            let txtColor = COLOR.LESSON_HOT_TXT;
                            
                            if (lesson.area.includes("ピラティス") || lesson.area.includes("リフォーマー")) {
                                bgColor = COLOR.LESSON_REFORMER_BG;
                                txtColor = COLOR.LESSON_REFORMER_TXT;
                            }

                            for(let k=0; k<duration && (startColIdx+k)<TOTAL_COLS; k++) {
                                rowSt.bg[startColIdx+k] = bgColor;
                            }
                            
                            rowSt.val[startColIdx] = lesson.displayName;
                            rowSt.fs[startColIdx] = 10; 
                            rowSt.fw[startColIdx] = "bold";
                            rowSt.fc[startColIdx] = txtColor; 

                            mergeRanges.push({ 
                                row: currentRowIdx + 1, 
                                col: startColIdx + 1, 
                                numRow: 1, 
                                numCol: duration 
                            });
                            
                            shiftBorderRanges.push({ 
                                row: currentRowIdx + 1, 
                                col: startColIdx + 1, 
                                numRow: 1, 
                                numCol: duration 
                            });
                        }
                    }
                }
            });
        }
        
        values.push(rowSt.val); backgrounds.push(rowSt.bg); fontColors.push(rowSt.fc); fontSizes.push(rowSt.fs); fontWeights.push(rowSt.fw); hAligns.push(rowSt.ha); vAligns.push(rowSt.va); wraps.push(rowSt.wp);
        rowHeightRequests.push({ row: currentRowIdx+1, height: HEIGHT_ROW });
        currentRowIdx++;
    });

    let dailyNotesMap = {}; 

    staffList.forEach(staff => {
      let rowS = createNewRow();

      let myColor = staff.color || COLOR.WHITE;
      let shiftText = "";
      let noteText = "";

      if (staff.rowIndex) {
        shiftText = String(shiftDataRange[staff.rowIndex-1][i] || "").trim();
        if (staff.memoOffset) {
           noteText = String(shiftDataRange[staff.rowIndex - 1 + staff.memoOffset][i] || "").trim();
        }
      }

      // メモ行やシフト行に「ひばり」等の記載があれば他店舗扱い
      let isOtherShop = false;
      let combinedText = shiftText + " " + noteText;
      if (combinedText.includes("ひばり")) {
          isOtherShop = true;
      } else if (combinedText.includes("経堂")) {
          myColor = "#f4cccc"; 
      }

      rowS.val[1] = staff.name; 
      rowS.bg[1] = isOtherShop ? COLOR.STRIPE_GRAY : myColor; 
      rowS.fw[1] = "bold";
      rowS.fs[1] = 10; 
      if (isOtherShop) rowS.fc[1] = "#999999";

      for (let t = 2; t < TOTAL_COLS; t++) {
        rowS.fs[t] = 10;
        rowS.fw[t] = "bold"; 
      }

      if (staff.rowIndex && !isOtherShop) {
        let shiftStartMin = -1;
        let shiftEndMin = -1;

        if (shiftText && /\d/.test(shiftText)) {
          let clean = shiftText.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0)-0xFEE0)).replace(/\s/g, "");
          let m = clean.match(/^(\d{1,2})[:：]?(\d{2})?[-~～](\d{1,2})[:：]?(\d{2})?/);
          if (m) {
            let sH = parseInt(m[1]), sM = parseInt(m[2] || "00"), eH = parseInt(m[3]), eM = parseInt(m[4] || "00");
            
            shiftStartMin = sH * 60 + sM;
            shiftEndMin = eH * 60 + eM;

            let sB = (sH-START_HOUR)*4 + Math.floor(sM/15), eB = (eH-START_HOUR)*4 + Math.floor(eM/15);
            for(let k=sB; k<eB && (2+k)<TOTAL_COLS; k++) rowS.bg[2+k] = COLOR.SHIFT_BAR;
          }
        }

        if (noteText) {
          if (!dailyNotesMap[noteText]) {
             dailyNotesMap[noteText] = [];
          }
          dailyNotesMap[noteText].push(staff.name);

          let timeMatch = noteText.match(/(\d{1,2})[:：]?(\d{2})?[-~～](\d{1,2})[:：]?(\d{2})?/);
          if (timeMatch) {
            let nStartH = parseInt(timeMatch[1]);
            let nStartM = parseInt(timeMatch[2] || "00");
            let nEndH = parseInt(timeMatch[3]);
            let nEndM = parseInt(timeMatch[4] || "00");
            
            let nStartMin = nStartH * 60 + nStartM;
            let nEndMin = nEndH * 60 + nEndM;

            let durationMin = nEndMin - nStartMin;
            let isSameAsShift = (nStartMin === shiftStartMin && nEndMin === shiftEndMin);

            // シフトと同じ時間、または4時間(240分)以上の長大なタスクはバーを描画しない
            if (!isSameAsShift && durationMin > 0 && durationMin < 240) {
               let nStartLimit = START_HOUR * 60;
               if (nEndMin > nStartLimit && nStartMin < (END_HOUR + 1) * 60) {
                   let sB_N = Math.floor((Math.max(nStartMin, nStartLimit) - nStartLimit) / 15);
                   let eB_N = Math.ceil((Math.min(nEndMin, (END_HOUR+1)*60) - nStartLimit) / 15);
                   
                   let startColIdxN = 2 + sB_N; 
                   let durationN = eB_N - sB_N;
                   
                   if (durationN > 0) {
                       // ★修正：タスクの時間は「黒塗り」のみにする。文字は入れない。
                       for(let k=0; k<durationN && (startColIdxN+k)<TOTAL_COLS; k++) {
                           rowS.bg[startColIdxN+k] = COLOR.BLACK;
                       }

                       shiftBorderRanges.push({ 
                           row: currentRowIdx + 1, 
                           col: startColIdxN + 1, 
                           numRow: 1, 
                           numCol: durationN 
                       });
                   }
               }
            }
          }
        }

        if (shiftStartMin >= START_HOUR * 60 && shiftEndMin > shiftStartMin) {
          let sB = Math.floor((shiftStartMin - START_HOUR * 60) / 15);
          let eB = Math.floor((shiftEndMin - START_HOUR * 60) / 15);
          
          let startColIdx = 2 + sB; 
          let duration = eB - sB;
          
          if (duration > 0) {
              shiftBorderRanges.push({ 
                  row: currentRowIdx + 1, 
                  col: startColIdx + 1, 
                  numRow: 1, 
                  numCol: duration 
              });
          }
        }

        if (pgLessonMap[dowFull] && pgLessonMap[dowFull][staff.name]) {
            pgLessonMap[dowFull][staff.name].forEach(lesson => {
                let startLimit = START_HOUR * 60;
                
                if (lesson.end > startLimit && lesson.start < (END_HOUR + 1) * 60) {
                    let sB_L = Math.floor((Math.max(lesson.start, startLimit) - startLimit) / 15);
                    let eB_L = Math.ceil((Math.min(lesson.end, (END_HOUR+1)*60) - startLimit) / 15);
                    
                    let startColIdxL = 2 + sB_L;
                    let durationL = eB_L - sB_L;
                    
                    if (durationL > 0) {
                        for(let k=0; k<durationL && (startColIdxL+k)<TOTAL_COLS; k++) {
                            rowS.bg[startColIdxL+k] = COLOR.BLACK; 
                        }

                        shiftBorderRanges.push({ 
                            row: currentRowIdx + 1, 
                            col: startColIdxL + 1, 
                            numRow: 1, 
                            numCol: durationL 
                        });
                    }
                }
            });
        }
      }
      values.push(rowS.val); backgrounds.push(rowS.bg); fontColors.push(rowS.fc); fontSizes.push(rowS.fs); fontWeights.push(rowS.fw); hAligns.push(rowS.ha); vAligns.push(rowS.va); wraps.push(rowS.wp);
      
      rowHeightRequests.push({ row: currentRowIdx+1, height: HEIGHT_ROW }); 
      currentRowIdx++;
    });

    let blockHeight = 1 + 2 + staffList.length; 
    mergeRanges.push({ row: startBlockRowIdx+1, col: 1, numRow: blockHeight, numCol: 1 });
    backgrounds[startBlockRowIdx][0] = COLOR.WHITE; 
    vAligns[startBlockRowIdx][0] = "middle";

    solidAllRanges.push({ row: startBlockRowIdx+1, col: 1, numRow: blockHeight, numCol: TOTAL_COLS });
    timeGridRanges.push({ row: startBlockRowIdx+1, col: 3, numRow: blockHeight, numCol: TOTAL_TIME_COLS });
    for (let h=1; h<HOURS_COUNT; h++) hourMarkerRanges.push({ row: startBlockRowIdx+1, col: 3+(h*4), numRow: blockHeight, numCol: 1 });

    let rowSep = createNewRow();
    values.push(rowSep.val); backgrounds.push(rowSep.bg); fontColors.push(rowSep.fc); fontSizes.push(rowSep.fs); fontWeights.push(rowSep.fw); hAligns.push(rowSep.ha); vAligns.push(rowSep.va); wraps.push(rowSep.wp);
    rowHeightRequests.push({ row: currentRowIdx+1, height: HEIGHT_BLOCK_SEP });
    currentRowIdx++;

    let dailyNotesStrList = [];
    let memoLines = 0;
    for (let note in dailyNotesMap) {
       let staffsStr = dailyNotesMap[note].join("・");
       dailyNotesStrList.push(`●${staffsStr}：${note}`);
       memoLines++;
    }
    let memoContent = dailyNotesStrList.join("\n");
    let calculatedHeight = HEIGHT_SUMMARY_FIXED;

    let rowSumH = createNewRow();
    for(let c=0; c<TOTAL_COLS; c++) rowSumH.bg[c] = COLOR.WHITE; 
    
    let idx1 = 0; 
    
    rowSumH.val[idx1] = "メモ"; 
    
    rowSumH.bg[idx1] = COLOR.SUMMARY_MEMO_TITLE_BG; 
    rowSumH.fc[idx1] = COLOR.SUMMARY_MEMO_TITLE_TXT; 
    rowSumH.fw[idx1] = "bold";
    
    values.push(rowSumH.val); backgrounds.push(rowSumH.bg); fontColors.push(rowSumH.fc); fontSizes.push(rowSumH.fs); fontWeights.push(rowSumH.fw); hAligns.push(rowSumH.ha); vAligns.push(rowSumH.va); wraps.push(rowSumH.wp);
    
    mergeRanges.push({ row: currentRowIdx+1, col: 1, numRow: 1, numCol: TOTAL_COLS });
    
    rowHeightRequests.push({ row: currentRowIdx+1, height: HEIGHT_SUMMARY_HEADER }); 
    currentRowIdx++;

    let rowSumB = createNewRow();
    for(let c=0; c<TOTAL_COLS; c++) rowSumB.bg[c] = COLOR.WHITE; 
    
    rowSumB.val[idx1] = memoContent; 

    let fsMemo = 10;
    if (memoLines > 15) fsMemo = 6;
    else if (memoLines > 10) fsMemo = 8;

    rowSumB.fs[idx1] = fsMemo;
    rowSumB.ha[idx1] = "left"; rowSumB.va[idx1] = "top"; rowSumB.wp[idx1] = true;

    values.push(rowSumB.val); backgrounds.push(rowSumB.bg); fontColors.push(rowSumB.fc); fontSizes.push(rowSumB.fs); fontWeights.push(rowSumB.fw); hAligns.push(rowSumB.ha); vAligns.push(rowSumB.va); wraps.push(rowSumB.wp);
    
    mergeRanges.push({ row: currentRowIdx+1, col: 1, numRow: 1, numCol: TOTAL_COLS });
    
    rowHeightRequests.push({ row: currentRowIdx+1, height: calculatedHeight }); 
    currentRowIdx++;

    summaryBorderRanges.push({ row: currentRowIdx-1, col: 1, numRow: 2, numCol: TOTAL_COLS });

    if (i !== dates.length - 1) {
      values.push(new Array(TOTAL_COLS).fill("")); 
      backgrounds.push(new Array(TOTAL_COLS).fill(COLOR.WHITE));
      fontColors.push(new Array(TOTAL_COLS).fill("black"));
      fontSizes.push(new Array(TOTAL_COLS).fill(10)); 
      fontWeights.push(new Array(TOTAL_COLS).fill("normal"));
      hAligns.push(new Array(TOTAL_COLS).fill("center")); 
      vAligns.push(new Array(TOTAL_COLS).fill("middle")); 
      wraps.push(new Array(TOTAL_COLS).fill(false));
      
      rowHeightRequests.push({ row: currentRowIdx+1, height: HEIGHT_SPACER }); 
      currentRowIdx++;
    }
  }

  ss.toast("シートを整形中...", "最終段階", 10);
  const totalRows = values.length;
  
  if (totalRows === 0) throw new Error("シフト表から日付データを読み取れませんでした。3行目に日付があるか確認してください。");

  function removeOverlaps(ranges) {
    let result = [];
    for (let i = ranges.length - 1; i >= 0; i--) {
      let current = ranges[i];
      let isOverlap = result.some(r => {
        let rowOverlap = (current.row < r.row + r.numRow) && (current.row + current.numRow > r.row);
        let colOverlap = (current.col < r.col + r.numCol) && (current.col + current.numCol > r.col);
        return rowOverlap && colOverlap;
      });
      if (!isOverlap) result.push(current);
    }
    return result;
  }

  mergeRanges = removeOverlaps(mergeRanges);

  const range = targetSheet.getRange(1, 1, totalRows, TOTAL_COLS);
  range.setValues(values).setBackgrounds(backgrounds).setFontColors(fontColors).setFontSizes(fontSizes).setFontWeights(fontWeights)
       .setHorizontalAlignments(hAligns).setVerticalAlignments(vAligns).setWraps(wraps).setFontFamily("Meiryo");

  targetSheet.setColumnWidth(1, 50); 
  targetSheet.setColumnWidth(2, 75); 
  for (let c=3; c<=TOTAL_COLS; c++) targetSheet.setColumnWidth(c, 17); 

  rowHeightRequests.forEach(req => { if(req.row <= totalRows) targetSheet.setRowHeight(req.row, req.height); });

  solidAllRanges.forEach(r => {
      try { targetSheet.getRange(r.row, r.col, r.numRow, r.numCol).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID); } catch(e){}
  });
  
  timeGridRanges.forEach(r => {
      try { targetSheet.getRange(r.row, r.col, r.numRow, r.numCol).setBorder(null, null, null, null, true, null, "black", SpreadsheetApp.BorderStyle.DOTTED); } catch(e){}
  });
  
  hourMarkerRanges.forEach(r => {
      try { targetSheet.getRange(r.row, r.col, r.numRow, r.numCol).setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID); } catch(e){}
  });

  shiftBorderRanges.forEach(r => {
      if (r.numCol > 0) {
          try {
              targetSheet.getRange(r.row, r.col, r.numRow, r.numCol)
                  .setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
          } catch(e){}
      }
  });

  summaryBorderRanges.forEach(r => {
      try { targetSheet.getRange(r.row, r.col, r.numRow, r.numCol).setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM); } catch(e){}
  });

  mergeRanges.forEach(m => {
    try {
      if (m.numRow === 1 && m.numCol === 1) return;
      targetSheet.getRange(m.row, m.col, m.numRow, m.numCol).merge();
    } catch(e) {}
  });

  const maxRowsKyodo = targetSheet.getMaxRows();
  if (maxRowsKyodo > totalRows) {
    targetSheet.deleteRows(totalRows + 1, maxRowsKyodo - totalRows);
  }
  const maxColsKyodo = targetSheet.getMaxColumns();
  if (maxColsKyodo > TOTAL_COLS) {
    targetSheet.deleteColumns(TOTAL_COLS + 1, maxColsKyodo - TOTAL_COLS);
  }

  ss.toast("経堂レイバー作成完了！", "成功", 3);
  
  const ui = SpreadsheetApp.getUi();
  ui.alert("作成完了！\n印刷設定で「A4縦」「幅に合わせる」を選択してください。\n下部メモ欄を統合し、全幅で表示するようにしました。");
}

/* ========================================================
 * 経堂レイバー：既存シートを消さず1名だけ追記（手入力維持用）
 * ======================================================== */

function showAppendSaikiLaborDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = [];
  ss.getSheets().forEach(function (sheet) {
    if (sheet.getName().indexOf("シフト") !== -1) sheetNames.push(sheet.getName());
  });
  if (sheetNames.length === 0) {
    SpreadsheetApp.getUi().alert("「シフト」と名のつくシートが見つかりませんでした。");
    return;
  }

  const htmlOutput = HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;padding:10px;">' +
    '<p style="font-weight:bold;color:#b71c1c;">【経堂】齋木さんの行だけ追記</p>' +
    '<p style="font-size:12px;color:#666;">既存のレイバーシートは削除しません。手入力した内容はそのまま残ります。<br>既に追記済みで位置がずれている場合も、実行すると1行上へ戻して枠線を整えます。</p>' +
    '<select id="sheetName" style="width:100%;padding:10px;margin-bottom:15px;">' +
    sheetNames.map(function (n) { return '<option value="' + n + '">' + n + '</option>'; }).join("") +
    '</select>' +
    '<button type="button" id="submitBtn" onclick="run()" style="width:100%;background:#b71c1c;color:#fff;border:none;padding:12px;border-radius:4px;font-weight:bold;cursor:pointer;">齋木さんを追記</button>' +
    '<script>' +
    'function run(){' +
    'var b=document.getElementById("submitBtn");b.disabled=true;b.innerText="処理中...";' +
    'google.script.run.withSuccessHandler(function(m){alert(m);google.script.host.close();})' +
    '.withFailureHandler(function(e){alert("エラー: "+e.message);b.disabled=false;b.innerText="齋木さんを追記";})' +
    '.appendKyodoStaffToExistingLabor(document.getElementById("sheetName").value);' +
    '}' +
    '</script></div>'
  ).setWidth(380).setHeight(260);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "齋木さんをレイバー追記");
}

/**
 * 既存の経堂レイバーシートに、指定スタッフの行を日別ブロックへ追記
 * @return {string} 完了メッセージ
 */
function appendKyodoStaffToExistingLabor(sourceSheetName) {
  return appendKyodoStaffToExistingLabor_(sourceSheetName, SAIKI_KYODO_STAFF);
}

function appendKyodoStaffToExistingLabor_(sourceSheetName, staffSpec) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName(sourceSheetName);
  if (!srcSheet) throw new Error("シフトシートが見つかりません: " + sourceSheetName);

  const monthMatch = sourceSheetName.match(/(\d{1,2})月/);
  const targetSheetName = monthMatch
    ? monthMatch[1] + "月 レイバースケジュール(経堂)"
    : "レイバースケジュール(経堂)";
  const targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    throw new Error("「" + targetSheetName + "」がありません。シート名を確認してください。");
  }

  const layout = getKyodoLaborLayout_();
  const { START_HOUR, END_HOUR, TOTAL_COLS } = layout;

  const daysCount = srcSheet.getLastColumn() - 2;
  if (daysCount < 1) throw new Error("シフト表の日付列が読み取れません。");

  const shiftRow = staffSpec.shiftSheetRow;
  const shiftData = srcSheet.getRange(shiftRow, 3, 1, daysCount).getValues()[0];
  const memoData = staffSpec.memoOffset
    ? srcSheet.getRange(shiftRow + staffSpec.memoOffset, 3, 1, daysCount).getValues()[0]
    : [];

  // 既に1行下にずれている齋木行を先に修復
  let relocated = fixAllMisplacedSaikiRows_(targetSheet);

  const data = targetSheet.getDataRange().getValues();
  const blockStarts = [];
  for (let r = 0; r < data.length; r++) {
    if (/^\d+日$/.test(String(data[r][0]).trim())) blockStarts.push(r);
  }
  if (blockStarts.length === 0) {
    throw new Error("レイバーシートに日付ブロック（○日）が見つかりません。");
  }

  let inserted = 0;
  let skipped = 0;
  let repainted = 0;

  for (let b = blockStarts.length - 1; b >= 0; b--) {
    const dayIdx = b;
    if (dayIdx >= daysCount) continue;

    const start = blockStarts[b];
    const end = (b + 1 < blockStarts.length) ? blockStarts[b + 1] : data.length;
    let memoR = -1;
    let saikiR = -1;
    for (let r = start; r < end; r++) {
      const aVal = String(data[r][0]).trim();
      const bVal = String(data[r][1]).trim();
      if (aVal === "メモ" && memoR < 0) memoR = r;
      if (isSaikiStaffName_(bVal)) saikiR = r;
    }
    if (memoR < 0) continue;

    const shiftText = String(shiftData[dayIdx] || "").trim();
    const noteText = memoData.length ? String(memoData[dayIdx] || "").trim() : "";

    if (saikiR >= 0) {
      skipped++;
      continue;
    }

    // メモ直前＝区切り行。区切り行の上にスタッフ行を挿入
    const insertAt = memoR;
    targetSheet.insertRowBefore(insertAt);

    paintKyodoStaffRowOnSheet_(
      targetSheet, insertAt, staffSpec.name, staffSpec.color,
      shiftText, noteText, TOTAL_COLS, START_HOUR, END_HOUR
    );
    inserted++;
  }

  SpreadsheetApp.flush();
  const ui = SpreadsheetApp.getActiveSpreadsheet();
  let repainted = 0;

  try {
    if (inserted === 0 && relocated === 0) {
      // 追記済み：齋木行の書式だけ直す（全ブロック修復はスキップ＝高速）
      ui.toast("齋木行の書式を修正中...", "処理中", 60);
      repainted = refreshAllSaikiStaffRowsFast_(targetSheet, data, blockStarts, shiftData, staffSpec, TOTAL_COLS, START_HOUR, END_HOUR);
    } else {
      ui.toast("レイアウトを整備中...", "処理中", 120);
      repairKyodoLaborSheetLayout_(targetSheet);
      repainted = refreshAllSaikiStaffRowsFast_(targetSheet, data, blockStarts, shiftData, staffSpec, TOTAL_COLS, START_HOUR, END_HOUR);
    }
    ui.toast("", "", 1);
  } catch (layoutErr) {
    return "齋木さん（" + staffSpec.name + "）の追記は完了しましたが、レイアウト整備でエラーがありました。\n" +
      "追記: " + inserted + "日分 / スキップ: " + skipped + "日分\n\n" + layoutErr.message +
      "\n\nもう一度実行すると整備だけ再試行されます。";
  }

  let msg = "齋木さん（" + staffSpec.name + "）の行を追記しました。\n" +
    "追記: " + inserted + "日分 / 既存: " + skipped + "日分（書式再適用: " + repainted + "日分）";
  if (relocated > 0) {
    msg += "\n位置修復: " + relocated + "日分（1行上に戻しました）";
  }
  msg += "\n\n結合セル・枠線・文字色（黒）を整えました。";
  return msg;
}

/** メモ直前に誤配置された齋木行を区切り行の上へ移動 */
function fixAllMisplacedSaikiRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  let fixed = 0;
  for (let r = lastRow - 1; r >= 1; r--) {
    if (String(colA[r][0]).trim() !== "メモ") continue;
    if (fixMisplacedSaikiBeforeMemo_(sheet, r + 1)) fixed++;
  }
  return fixed;
}

function fixMisplacedSaikiBeforeMemo_(sheet, memoRow) {
  try {
    const saikiRow = memoRow - 1;
    if (saikiRow < 2) return false;

    const saikiName = String(sheet.getRange(saikiRow, 2).getValue()).trim();
    if (!isSaikiStaffName_(saikiName)) return false;

    const sepRow = memoRow - 2;
    const sepName = String(sheet.getRange(sepRow, 2).getValue()).trim();
    const sepDate = String(sheet.getRange(sepRow, 1).getValue()).trim();
    if (sepName !== "" || sepDate !== "" || /^\d+日$/.test(sepDate)) return false;

    const layout = getKyodoLaborLayout_();
    const totalCols = Math.max(sheet.getLastColumn(), layout.TOTAL_COLS);

    safeBreakApartMergedRanges_(sheet.getRange(sepRow, 1, memoRow - sepRow + 1, totalCols));

    sheet.deleteRow(saikiRow);
    sheet.insertRowBefore(sepRow);

    paintKyodoStaffRowOnSheet_(
      sheet, sepRow, SAIKI_KYODO_STAFF.name, SAIKI_KYODO_STAFF.color,
      "", "", totalCols, layout.START_HOUR, layout.END_HOUR
    );

    resetKyodoSeparatorRow_(sheet, sepRow + 1, totalCols);
    return true;
  } catch (e) {
    return false;
  }
}

/** 日別ブロックの区切り行を空行に戻す */
function resetKyodoSeparatorRow_(sheet, row, totalCols) {
  const rng = sheet.getRange(row, 1, 1, totalCols);
  safeBreakApartMergedRanges_(rng);
  rng.clearContent();
  rng.setBackground("#ffffff");
  sheet.setRowHeight(row, 8);
}

/** 全日期ブロックのA列結合・枠線・メモ結合・日付ヘッダーを再整備 */
function repairKyodoLaborSheetLayout_(sheet) {
  const layout = getKyodoLaborLayout_();
  const { HOURS_COUNT, TOTAL_TIME_COLS, TOTAL_COLS } = layout;

  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;

  const colA = sheet.getRange(1, 1, lastRow, 1).getValues();
  const dateRows = [];
  for (let r = 0; r < colA.length; r++) {
    if (/^\d+日$/.test(String(colA[r][0]).trim())) dateRows.push(r + 1);
  }

  for (let i = 0; i < dateRows.length; i++) {
    try {
      const dateRow = dateRows[i];
      const nextBound = (i + 1 < dateRows.length) ? dateRows[i + 1] : lastRow + 1;

      let memoRow = -1;
      for (let r = dateRow; r < nextBound; r++) {
        if (String(colA[r - 1][0]).trim() === "メモ") {
          memoRow = r;
          break;
        }
      }
      if (memoRow < 0) continue;

      const studioRow = dateRow + 1;
      const blockEndRow = memoRow - 1;
      const blockHeight = blockEndRow - studioRow + 1;
      if (blockHeight < 1) continue;

      // 日付行は触らず、スタジオ〜区切り行だけ結合解除
      safeBreakApartMergedRanges_(sheet.getRange(studioRow, 1, blockHeight, TOTAL_COLS));

      // A列：日付セルは単独、本文は縦結合
      safeBreakApartMergedRanges_(sheet.getRange(dateRow, 1));
      const aBlock = sheet.getRange(studioRow, 1, blockHeight, 1);
      safeMergeRange_(aBlock);
      aBlock.setVerticalAlignment("middle");
      sheet.getRange(dateRow, 1).setVerticalAlignment("middle");

      // 日付ヘッダー行（時刻ラベル・曜日）を復元
      repairKyodoDateHeaderRow_(sheet, dateRow, layout);

      // 区切り行
      resetKyodoSeparatorRow_(sheet, blockEndRow, TOTAL_COLS);

      // 外枠・時間グリッド
      sheet.getRange(studioRow, 1, blockHeight, TOTAL_COLS)
        .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

      sheet.getRange(studioRow, 3, blockHeight, TOTAL_TIME_COLS)
        .setBorder(null, null, null, null, true, null, "black", SpreadsheetApp.BorderStyle.DOTTED);

      for (let h = 1; h < HOURS_COUNT; h++) {
        sheet.getRange(studioRow, 3 + h * 4, blockHeight, 1)
          .setBorder(null, true, null, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
      }

      // メモ欄
      repairKyodoMemoRows_(sheet, memoRow, nextBound, TOTAL_COLS);
    } catch (e) { /* 1日分失敗しても他は続行 */ }
  }
}

/** 日付行の時刻ヘッダー結合・書式を復元 */
function repairKyodoDateHeaderRow_(sheet, dateRow, layout) {
  const { HOURS_COUNT, START_HOUR, TOTAL_COLS } = layout;

  safeBreakApartMergedRanges_(sheet.getRange(dateRow, 3, 1, TOTAL_COLS - 2));

  for (let h = 0; h < HOURS_COUNT; h++) {
    const col = 3 + h * 4;
    const hourLabel = (START_HOUR + h) + ":00";
    safeMergeRange_(sheet.getRange(dateRow, col, 1, 4));
    sheet.getRange(dateRow, col)
      .setValue(hourLabel)
      .setFontWeight("bold")
      .setFontSize(10)
      .setFontColor("#000000")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setFontFamily("Meiryo");
  }

  sheet.getRange(dateRow, 1)
    .setFontWeight("bold")
    .setFontSize(13)
    .setFontColor("#000000")
    .setVerticalAlignment("middle")
    .setFontFamily("Meiryo");

  sheet.getRange(dateRow, 2)
    .setFontWeight("bold")
    .setFontSize(14)
    .setFontColor("#000000")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontFamily("Meiryo");

  sheet.setRowHeight(dateRow, 28);
}

/** メモ見出し・本文の結合と書式を復元 */
function repairKyodoMemoRows_(sheet, memoRow, nextBound, totalCols) {
  safeBreakApartMergedRanges_(sheet.getRange(memoRow, 1, 1, totalCols));
  safeMergeRange_(sheet.getRange(memoRow, 1, 1, totalCols));
  sheet.getRange(memoRow, 1)
    .setValue("メモ")
    .setBackground("#cc0000")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontFamily("Meiryo");
  sheet.setRowHeight(memoRow, 22);

  if (memoRow + 1 < nextBound) {
    const nextA = String(sheet.getRange(memoRow + 1, 1).getValue()).trim();
    if (nextA !== "メモ" && !/^\d+日$/.test(nextA)) {
      safeBreakApartMergedRanges_(sheet.getRange(memoRow + 1, 1, 1, totalCols));
      safeMergeRange_(sheet.getRange(memoRow + 1, 1, 1, totalCols));
      sheet.getRange(memoRow + 1, 1)
        .setFontColor("#000000")
        .setHorizontalAlignment("left")
        .setVerticalAlignment("top")
        .setWrap(true)
        .setFontFamily("Meiryo");
    }
  }
}

/** シート上の全齋木行に書式を再適用（一括読み込みで高速化） */
function refreshAllSaikiStaffRowsFast_(sheet, data, blockStarts, shiftData, staffSpec, totalCols, startHour, endHour) {
  let count = 0;
  for (let r = 0; r < data.length; r++) {
    const nm = String(data[r][1] || "").trim();
    if (!isSaikiStaffName_(nm)) continue;

    let dayIdx = 0;
    for (let i = blockStarts.length - 1; i >= 0; i--) {
      if (blockStarts[i] <= r) {
        dayIdx = i;
        break;
      }
    }

    const shiftText = (shiftData && dayIdx < shiftData.length)
      ? String(shiftData[dayIdx] || "").trim()
      : "";

    paintKyodoStaffRowOnSheet_(
      sheet, r + 1, staffSpec.name, staffSpec.color,
      shiftText, "", totalCols, startHour, endHour
    );
    count++;
  }
  return count;
}

/** @deprecated refreshAllSaikiStaffRowsFast_ を使用 */
function refreshAllSaikiStaffRows_(sheet, shiftData, staffSpec, totalCols, startHour, endHour) {
  const data = sheet.getDataRange().getValues();
  const blockStarts = [];
  for (let r = 0; r < data.length; r++) {
    if (/^\d+日$/.test(String(data[r][0]).trim())) blockStarts.push(r);
  }
  return refreshAllSaikiStaffRowsFast_(sheet, data, blockStarts, shiftData, staffSpec, totalCols, startHour, endHour);
}

/** レイバー1行分のシフトバー描画（追記用・通常作成と同じ書式） */
function paintKyodoStaffRowOnSheet_(sheet, row, staffName, staffColor, shiftText, noteText, TOTAL_COLS, START_HOUR, END_HOUR) {
  const SHIFT_BAR = "#cccccc";
  const WHITE = "#ffffff";
  const STRIPE = "#f7f7f7";
  const BLACK = "#000000";
  const timeColCount = TOTAL_COLS - 1;

  const bg = [];
  const fc = [];
  const fw = [];
  const fs = [];
  const ha = [];
  const va = [];
  for (let t = 0; t < timeColCount; t++) {
    bg.push(t >= 1 && Math.floor(t / 4) % 2 !== 0 ? STRIPE : WHITE);
    fc.push(BLACK);
    fw.push(t === 0 ? "bold" : "bold");
    fs.push(10);
    ha.push("center");
    va.push("middle");
  }
  bg[0] = staffColor;
  fc[0] = BLACK;

  const rowRange = sheet.getRange(row, 2, 1, timeColCount);
  rowRange
    .setBackgrounds([bg])
    .setFontColors([fc])
    .setFontWeights([fw])
    .setFontSizes([fs])
    .setHorizontalAlignments([ha])
    .setVerticalAlignments([va])
    .setFontFamily("Meiryo");

  rowRange.getCell(1, 1)
    .setValue(staffName)
    .setFontWeight("bold")
    .setFontSize(10)
    .setFontColor(BLACK)
    .setBackground(staffColor)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  if (shiftText && /\d/.test(shiftText)) {
    const clean = shiftText.replace(/[０-９]/g, function (s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    }).replace(/\s/g, "");
    const m = clean.match(/^(\d{1,2})[:：]?(\d{2})?[-~～](\d{1,2})[:：]?(\d{2})?/);
    if (m) {
      const sH = parseInt(m[1], 10);
      const sM = parseInt(m[2] || "0", 10);
      const eH = parseInt(m[3], 10);
      const eM = parseInt(m[4] || "0", 10);
      const sB = (sH - START_HOUR) * 4 + Math.floor(sM / 15);
      const eB = (eH - START_HOUR) * 4 + Math.floor(eM / 15);
      for (let k = sB; k < eB && (2 + k) < TOTAL_COLS; k++) {
        sheet.getRange(row, 3 + k).setBackground(SHIFT_BAR);
      }
      if (eB > sB) {
        try {
          sheet.getRange(row, 3 + sB, 1, eB - sB)
            .setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID);
        } catch (e) { /* ignore */ }
      }
    }
  }

  // シフトバー設定後も名前セルは必ず黒（白文字化を防止）
  sheet.getRange(row, 2)
    .setFontColor(BLACK)
    .setFontFamily("Meiryo");

  sheet.setRowHeight(row, 24);
}