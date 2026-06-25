/**
 * 勤怠システム貼り付け ＆ カレンダー自動登録 統合スクリプト (第7エリア用)
 *
 * 【カレンダー同期の見た目】
 * ・勤務 … 勤務場所アイコン「勤務場所: 経堂」（棒線）。太い「勤務」ブロックは作らない
 * ・公休・有休・勤務外 … 表示名「不在」（勤務外は outOfOffice・斜線）
 * ・会議 … 通常予定（不透明・予約不可）
 *
 * 【必須】GAS「サービス」→ Google Calendar API（ID: Calendar）v3
 * 【反映】スプレッドシートで クリア → 登録（古い勤務ブロックを消すため）
 */

/* ========================================================
 * 【設定1】カレンダー同期用の設定
 * ======================================================== */
const CAL_CONFIG = {
  // ★変更箇所1：同期対象スタッフの設定
  // memoOffset: 名前の行から何行下にメモがあるか（指定がない場合は1行下になります）
  STAFF_LIST: [
    { name: "日下 竜汰", email: "r-kusaka@okamoto-group.co.jp", memoOffset: 1, defaultStore: "経堂" },
    { name: "蜂谷 有加", email: "yuka-hachiya@okamoto-group.co.jp", memoOffset: 1 },
    { name: "中田 花子", email: "h-nakata@okamoto-group.co.jp", memoOffset: 2 }, // +1行目勤務地、+2行目メモ
    { name: "石田 美絵", email: "mie-ishida@okamoto-group.co.jp", memoOffset: 1, defaultStore: "経堂" }
  ],

  /**
   * イレギュラー勤務（カレンダー同期のみ）
   * ・指定曜日以外 → シートの記載に関わらず「不在」(公休扱い)
   * ・指定曜日 → 固定時間（シートが休み/有休ならシート優先）
   * キーは STAFF_LIST の name と一致させる
   */
  STAFF_WEEKLY_OVERRIDE: {
    "石田 美絵": {
      weekdayShifts: {
        2: "11:00-17:00", // 火曜
        3: "13:00-17:00"  // 水曜
      },
      offLabel: "公休",
      respectSheetOff: true
    }
  },

  // 配色テーマ（DISPLAY.COLOR_THEME で切り替え）
  // 使える色: GRAY, GRAPHITE, CHARCOAL, PALE_BLUE, PALE_GREEN, PALE_RED,
  //   LAVENDER, ORANGE, BLUE, GREEN, RED など CalendarApp.EventColor.*
  COLOR_THEMES: {
    classic: {
      WORK: CalendarApp.EventColor.PALE_RED,
      WORK_HIBARI: CalendarApp.EventColor.PALE_GREEN,
      OFF: CalendarApp.EventColor.PALE_BLUE,
      PTO: CalendarApp.EventColor.LAVENDER,
      UNDEF: CalendarApp.EventColor.ORANGE,
      BLOCK: CalendarApp.EventColor.GRAY,
      TASK: CalendarApp.EventColor.PALE_GREEN
    },
    monochrome: {
      WORK: CalendarApp.EventColor.GRAPHITE,
      WORK_HIBARI: CalendarApp.EventColor.GRAY,
      OFF: CalendarApp.EventColor.GRAY,
      PTO: CalendarApp.EventColor.CHARCOAL,
      UNDEF: CalendarApp.EventColor.GRAPHITE,
      BLOCK: CalendarApp.EventColor.GRAY,
      TASK: CalendarApp.EventColor.CHARCOAL
    }
  },

  DISPLAY: {
    // 'monochrome' | 'classic' — 変更後は クリア → 登録 で反映
    COLOR_THEME: 'monochrome',
    // 勤務外を予約不可にする（予約ページ対策）
    USE_OFF_BLOCKS: true,
    // outOfOffice=斜線表示＋予約ブロック、legacy=従来の×
    BLOCK_STYLE: 'outOfOffice',
    // 公休・有休・勤務外など「休み」系の統一表示名
    STATUS_TITLE: '不在',
    // 笠原さん型: 勤務場所イベント（経堂・ひばりが丘）
    USE_WORKING_LOCATION: true,
    DEFAULT_STORE: '経堂',
    // 勤務=勤務場所アイコンのみ（太い「勤務」ブロックは作らない）
    SHIFT_STYLE: 'location_bar',
    // 勤務場所の表示名（Google の勤務場所アイコン用）
    LOCATION_TITLE_PREFIX: '勤務場所: '
  },

  RE: {
    OFF:    /(休|公休|法定|法定外|欠勤|欠|×|不可|off)/i,
    PTO:    /(有休|有給|特休|特別休暇)/i, 
    TASK:   /(タスク)/i
  },

  THROTTLE_MS: 250, MAX_RETRY: 5, BACKOFF_BASE_MS: 500,
  PAUSE_EVERY_N_EVENTS: 20, PAUSE_MS: 2000, MIN_BLOCK_MINUTES: 10
};


/* ========================================================
 * 【設定2】勤怠システム（バイバイ）用の設定
 * ======================================================== */
const CONFIG = {
  DEST_SHEET_NAME: "バイバイ貼り付け",
  
  SHOPS: {
    "経堂": {
      "石田　美絵": 301532, "石田 美絵": 301532, "石田美絵": 301532, "石田": 301532,
      "行徳　由岐恵": 302703, "行徳 由岐恵": 302703, "行徳由岐恵": 302703, "行徳": 302703,
      "日下　竜汰": 303879, "日下 竜汰": 303879, "日下竜汰": 303879, "日下": 303879,
      "澤野　郁哉": 304642, "澤野 郁哉": 304642, "澤野郁哉": 304642, "澤野": 304642,
      "中田　花子": 30468, "中田 花子": 30468, "中田花子": 30468, "中田": 30468,
      "蜂谷　有加": 80779, "蜂谷 有加": 80779, "蜂谷有加": 80779, "蜂谷": 80779
    },
    "ひばりが丘": {
      "津田　加奈": "030396", "津田 加奈": "030396", "津田加奈": "030396", "津田": "030396",
      "吉田　薫理": "030400", "吉田 薫理": "030400", "吉田薫理": "030400", "吉田": "030400",
      "黒川　沙由美": 30204, "黒川 沙由美": 30204, "黒川沙由美": 30204, "黒川": 30204,
      "徳重　翠": 30331, "徳重 翠": 30331, "徳重翠": 30331, "徳重": 30331,
      "大野　雅代": 303523, "大野 雅代": 303523, "大野雅代": 303523, "大野": 303523,
      "手塚　柚衣": 30469, "手塚 柚衣": 30469, "手塚柚衣": 30469, "手塚": 30469
    }
  },
  
  PART_TIME_PTO_STAFF: {
    "黒川 沙由美": "10:00-16:00R0:00",
    "黒川": "10:00-16:00R0:00"
  },

  ROW_TYPES: [
    { code: 215001, name: "休日・休暇" },
    { code: 215201, name: "シフト" },
    { code: 215013, name: "休日・休暇(自動展開)" },
    { code: 215231, name: "シフト(自動展開)" }
  ],
  CODE_MAP: { "有給": 61, "有休": 61, "欠勤": 80, "欠": 80, "特別休暇": 62, "特休": 62, "半休前": 63, "半休後": 64, "半休": 63, "育休": 77, "法定": 10, "法定外": 20 },
  BREAK_THRESHOLD_MINUTES: 360
};

/*
 * ※ メニューの追加（onOpen）は「レイバー作成」のコードで一括して行っているため、
 * ここでは不要です（もし単独で動かす場合は追加してください）。
 */


/* ========================================================
 * 【機能1】勤怠システム（バイバイ）用の処理
 * ======================================================== */
function showDialog() {
  const htmlStr = `
    <!DOCTYPE html>
    <html>
      <head><base target="_top"><style>
        body { font-family: 'Meiryo', sans-serif; padding: 20px; color: #333; }
        h3 { margin-top: 0; color: #b71c1c; border-bottom: 2px solid #b71c1c; padding-bottom: 5px; }
        .form-group { margin-bottom: 15px; }
        label { display: block; margin-bottom: 5px; font-weight: bold; font-size: 14px; color: #555; }
        select { padding: 10px; font-size: 14px; width: 100%; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; }
        button { padding: 12px 20px; font-size: 16px; background-color: #0277bd; color: white; border: none; border-radius: 4px; cursor: pointer; width: 100%; font-weight: bold; margin-top: 10px; }
        button:disabled { background-color: #9e9e9e; cursor: not-allowed; }
        #message { margin-top: 15px; font-weight: bold; text-align: center; height: 20px; }
      </style></head>
      <body>
        <h3>勤怠データ作成ツール</h3>
        <div class="form-group"><label>店舗を選択</label>
          <select id="shopSelect">
            <option value="経堂">経堂</option>
            <option value="ひばりが丘">ひばりが丘</option>
          </select>
        </div>
        <div class="form-group"><label>対象シートを選択</label><select id="sheetSelect"><option value="">読込中...</option></select></div>
        <button id="runBtn" onclick="runScript()">作成する</button>
        <div id="message"></div>
        <script>
          google.script.run.withSuccessHandler(function(sheets) {
            var select = document.getElementById('sheetSelect'); select.innerHTML = '';
            if (sheets.length === 0) { select.innerHTML = '<option value="">シフトシートが見つかりません</option>'; document.getElementById('runBtn').disabled = true; return; }
            sheets.forEach(function(s) { var o = document.createElement('option'); o.value = s; o.text = s; select.appendChild(o); });
          }).getShiftSheetNames();
          function runScript() {
            var btn = document.getElementById('runBtn'), msg = document.getElementById('message'), sheet = document.getElementById('sheetSelect').value, shop = document.getElementById('shopSelect').value;
            if (!sheet || !shop) return;
            btn.disabled = true; btn.innerText = '処理中...'; msg.innerText = '';
            google.script.run.withSuccessHandler(function(r) { btn.innerText = '完了！'; msg.innerText = shop + 'のデータを作成しました'; msg.style.color = '#2e7d32'; setTimeout(function(){google.script.host.close();}, 2500); }).withFailureHandler(function(e) { btn.disabled = false; btn.innerText = '作成する'; msg.innerText = 'エラー: ' + e.message; msg.style.color = '#c62828'; }).processShiftData(sheet, shop);
          }
        </script>
      </body>
    </html>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(htmlStr).setWidth(350).setHeight(360), 'シフトデータ作成');
}

function getShiftSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()).filter(n => n.includes('シフト') && n !== CONFIG.DEST_SHEET_NAME);
}

function processShiftData(sourceSheetName, shopName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName(sourceSheetName);
  const destSheet = ss.getSheetByName(CONFIG.DEST_SHEET_NAME);
  if (!srcSheet || !destSheet) throw new Error("シートが見つかりません。");
  
  const staffCodeMap = {};
  const codeToFullNameMap = {}; 
  
  const targetShops = [CONFIG.SHOPS[shopName]];
  if (targetShops.length === 0 || !targetShops[0]) throw new Error("店舗の設定が見つかりません。");

  targetShops.forEach(shopMap => {
    Object.assign(staffCodeMap, shopMap);
    for (const key in shopMap) {
      const code = shopMap[key];
      if (!codeToFullNameMap[code] || key.includes('　')) {
        codeToFullNameMap[code] = key.replace(/[ 　]+/g, "　"); 
      }
    }
  });

  let year = new Date().getFullYear(), month = new Date().getMonth() + 1;
  const match = sourceSheetName.match(/(\d{4})年(\d{1,2})月/);
  if (match) { year = parseInt(match[1], 10); month = parseInt(match[2], 10); }
  const daysInMonth = new Date(year, month, 0).getDate();

  const dateHeaders = [], dayOfWeekStr = ["日", "月", "火", "水", "木", "金", "土"];
  for (let d = 1; d <= daysInMonth; d++) {
    const t = new Date(year, month - 1, d);
    dateHeaders.push(`${t.getMonth() + 1}/${t.getDate()}(${dayOfWeekStr[t.getDay()]})`);
  }
  const headerRow1 = [`${year}/${month}/1`, "～", `${year}/${month}/${daysInMonth}`, ""].concat(dateHeaders);

  destSheet.clear();
  if (destSheet.getMaxColumns() < headerRow1.length) {
    destSheet.insertColumnsAfter(destSheet.getMaxColumns(), headerRow1.length - destSheet.getMaxColumns());
  }
  
  const headerRange = destSheet.getRange(1, 1, 1, headerRow1.length);
  headerRange.setValues([headerRow1]);
  headerRange.setBackground('#404040'); 
  headerRange.setFontColor('#ffffff');  
  headerRange.setFontWeight('bold');    

  const lastRowSrc = Math.max(5, srcSheet.getLastRow()), lastColSrc = srcSheet.getLastColumn();
  const dateStartIdx = 2;

  const srcData = srcSheet.getRange(5, 1, lastRowSrc - 4, lastColSrc).getValues();
  const processedStaffCodes = new Set(); // ★1回処理したスタッフを記録して重複を防ぐ
  const outputData = [];

  srcData.forEach(srcRow => {
    const srcName = srcRow[0]; if (!srcName) return;
    const cleanName = String(srcName).trim();
    const normalizedName = cleanName.replace(/　/g, " ");
    const nameParts = normalizedName.split(" ");
    const lastName = nameParts.length > 0 ? nameParts[0] : cleanName;

    const staffCode = staffCodeMap[cleanName] || staffCodeMap[normalizedName] || staffCodeMap[lastName];
    if (!staffCode) return; 

    // ★自動合算をやめ、最初に名前が見つかった行（基本シフト行）のみを確実に処理する
    if (processedStaffCodes.has(staffCode)) return; 
    processedStaffCodes.add(staffCode);

    const outputName = codeToFullNameMap[staffCode] || cleanName.replace(/[ 　]+/g, "　");
    const noSpaceName = cleanName.replace(/[ 　]/g, "");
    
    const rowsForStaff = [];
    for (let r = 0; r < 4; r++) {
      const row = [staffCode, outputName, CONFIG.ROW_TYPES[r].code, CONFIG.ROW_TYPES[r].name];
      for (let d = 0; d < daysInMonth; d++) row.push("");
      rowsForStaff.push(row);
    }

    let weekHolidays = [];
    let hasExplicitHoutei = false;

    for (let d = 0; d < daysInMonth; d++) {
      const cellValue = String(srcRow[dateStartIdx + d] || "").trim();
      const currentDate = new Date(year, month - 1, d + 1);
      
      if (currentDate.getDay() === 0 && d !== 0) {
        if (weekHolidays.length > 0) {
          if (hasExplicitHoutei) {
            weekHolidays.forEach(idx => rowsForStaff[0][4 + idx] = 20);
          } else {
            rowsForStaff[0][4 + weekHolidays.pop()] = 10;
            weekHolidays.forEach(idx => rowsForStaff[0][4 + idx] = 20);
          }
          weekHolidays = [];
        }
        hasExplicitHoutei = false;
      }

      let valForShift = "";
      let isHoliday = false;

      if (cellValue === "") {
        isHoliday = true;
      } else {
        const timeRangeStr = convertToTimeFormat(cellValue);

        if (cellValue.includes("法定外")) {
          rowsForStaff[0][4 + d] = 20;
        } else if (cellValue.includes("法定")) {
          rowsForStaff[0][4 + d] = 10;
          hasExplicitHoutei = true;
        } else if (cellValue.includes("欠勤") || cellValue === "欠") {
          rowsForStaff[0][4 + d] = 80;
        } else if (cellValue.includes("半休前")) {
          rowsForStaff[0][4 + d] = 63;
          valForShift = "8:30-17:30R1:00"; 
        } else if (cellValue.includes("半休後")) {
          rowsForStaff[0][4 + d] = 64; 
          valForShift = "8:30-17:30R1:00"; 
        } else if (cellValue.includes("有給") || cellValue.includes("有休") || cellValue.includes("特別休暇") || cellValue.includes("特休")) {
          let code = 61;
          if (cellValue.includes("特別休暇") || cellValue.includes("特休")) code = 62;
          rowsForStaff[0][4 + d] = code;

          let customPtoShift = "";
          for (const pName in CONFIG.PART_TIME_PTO_STAFF) {
            const tName = pName.replace(/[ 　]/g, "");
            if (noSpaceName === tName || tName.startsWith(noSpaceName) || noSpaceName.startsWith(tName) || tName.includes(noSpaceName) || noSpaceName.includes(tName)) {
              customPtoShift = CONFIG.PART_TIME_PTO_STAFF[pName];
              break;
            }
          }
          
          if (customPtoShift) {
            valForShift = customPtoShift;
          } else {
            valForShift = "8:30-17:30R1:00";
          }
        } else if (cellValue === "育休") {
          rowsForStaff[0][4 + d] = 77;
          valForShift = "09:00-18:00R01:00";
        } else if (CONFIG.CODE_MAP[cellValue]) {
          rowsForStaff[0][4 + d] = CONFIG.CODE_MAP[cellValue];
        } else if (timeRangeStr) {
          valForShift = timeRangeStr;
        } else {
          isHoliday = true;
        }
      }

      if (isHoliday) weekHolidays.push(d);
      if (valForShift !== "") rowsForStaff[1][4 + d] = valForShift;

      if (d === daysInMonth - 1) {
        if (weekHolidays.length > 0) {
          if (hasExplicitHoutei) {
            weekHolidays.forEach(idx => rowsForStaff[0][4 + idx] = 20);
          } else {
            rowsForStaff[0][4 + weekHolidays.pop()] = 10;
            weekHolidays.forEach(idx => rowsForStaff[0][4 + idx] = 20);
          }
        }
      }
    }
    rowsForStaff.forEach(r => outputData.push(r));
  });

  if (outputData.length > 0) {
    if (destSheet.getMaxRows() < outputData.length + 1) {
      destSheet.insertRowsAfter(destSheet.getMaxRows(), (outputData.length + 1) - destSheet.getMaxRows());
    }
    destSheet.getRange(2, 1, outputData.length, 1).setNumberFormat('@');
    destSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
    
    const color1 = '#eaf5f9'; 
    const color2 = '#fff6e5'; 
    const numStaff = outputData.length / 4;
    
    for (let i = 0; i < numStaff; i++) {
      const bgColor = (i % 2 === 0) ? color1 : color2;
      destSheet.getRange(2 + i * 4, 1, 4, outputData[0].length).setBackground(bgColor);
    }
    const maxCols = destSheet.getMaxColumns();
    const usedCols = headerRow1.length;
    if (maxCols > usedCols) {
      destSheet.deleteColumns(usedCols + 1, maxCols - usedCols);
    }
    const maxRows = destSheet.getMaxRows();
    const usedRows = outputData.length + 1;
    if (maxRows > usedRows) {
      destSheet.deleteRows(usedRows + 1, maxRows - usedRows);
    }
  } else {
    SpreadsheetApp.getUi().alert(shopName + ' のデータが見つかりませんでした。');
  }
  return true;
}

function convertToTimeFormat(inputStr) {
  // ★修正：エラーの原因だった引数ミスを修正
  const src = String(inputStr||'').replace(/\r?\n/g, ' ').replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/[：:]/g, ':').replace(/[－ー〜−‐—]/g, '-').replace(/\s+/g, ' ').trim();
  let m = src.match(/(\d{1,2})(?::?(\d{2}))?\s*-\s*(\d{1,2})(?::?(\d{2}))?/);
  if (!m) return "";
  const sh = parseInt(m[1],10), sm = parseInt(m[2]||0,10);
  const eh = parseInt(m[3],10), em = parseInt(m[4]||0,10);
  const start = `${('0'+sh).slice(-2)}:${('0'+sm).slice(-2)}`;
  const end = `${('0'+eh).slice(-2)}:${('0'+em).slice(-2)}`;
  
  let duration = (eh * 60 + em) - (sh * 60 + sm);
  if (duration < 0) duration += 1440;
  let breakStr = (duration > CONFIG.BREAK_THRESHOLD_MINUTES) ? "R1:00" : "R0:00";
  return `${start}-${end}${breakStr}`;
}

/* ========================================================
 * 【機能2】カレンダー自動登録用の処理（★固定オフセット安定版）
 * ======================================================== */
function showCalendarSyncDialog() {
  const htmlStr = `
    <!DOCTYPE html>
    <html>
      <head><base target="_top"><style>
        * { box-sizing: border-box; }
        body { font-family: 'Segoe UI', 'Meiryo', sans-serif; margin: 0; padding: 24px; color: #1a1a1a; background: #fafafa; }
        h3 { margin: 0 0 20px; font-size: 15px; font-weight: 600; letter-spacing: 0.02em; }
        label { display: block; margin-bottom: 6px; font-size: 11px; color: #666; text-transform: uppercase; letter-spacing: 0.06em; }
        select { width: 100%; padding: 10px 12px; font-size: 14px; border: 1px solid #e0e0e0; border-radius: 8px; background: #fff; margin-bottom: 16px; }
        .btn-container { display: flex; gap: 8px; margin-top: 8px; }
        button { flex: 1; padding: 11px; font-size: 13px; font-weight: 600; border: none; border-radius: 8px; cursor: pointer; color: #fff; }
        #syncBtn { background: #1a1a1a; } #syncBtn:disabled { background: #bdbdbd; }
        #clearBtn { background: #fff; color: #1a1a1a; border: 1px solid #e0e0e0; } #clearBtn:disabled { color: #bdbdbd; }
        #message { margin-top: 16px; font-size: 12px; text-align: center; min-height: 20px; color: #666; white-space: pre-wrap; }
      </style></head>
      <body>
        <h3>カレンダー同期</h3>
        <label>シート</label><select id="sheetSelect"><option value="">…</option></select>
        <label>スタッフ</label><select id="staffSelect"></select>
        <div class="btn-container"><button id="syncBtn" onclick="runAction('sync')">登録</button><button id="clearBtn" onclick="runAction('clear')">クリア</button></div>
        <div id="message"></div>
        <script>
          google.script.run.withSuccessHandler(function(sheets) {
            var s = document.getElementById('sheetSelect'); s.innerHTML = '';
            sheets.forEach(function(sh) { var o = document.createElement('option'); o.value = sh; o.text = sh; s.appendChild(o); });
          }).getShiftSheetNamesForCalendar();
          google.script.run.withSuccessHandler(function(list) {
            var s = document.getElementById('staffSelect');
            list.forEach(function(st) { var o = document.createElement('option'); o.value = st.email; o.text = st.name; o.dataset.name = st.name; s.appendChild(o); });
          }).getStaffList();
          function runAction(type) {
            var sheet = document.getElementById('sheetSelect').value, staffSelect = document.getElementById('staffSelect'), email = staffSelect.value;
            var name = staffSelect.options[staffSelect.selectedIndex].dataset.name;
            if(!sheet || !email) return;
            document.getElementById('syncBtn').disabled = true; document.getElementById('clearBtn').disabled = true;
            document.getElementById('message').innerText = type==='sync' ? '処理中...しばらくお待ち下さい' : '削除中...';
            document.getElementById('message').style.color='#333';
            google.script.run.withSuccessHandler(function(r){ document.getElementById('syncBtn').disabled = false; document.getElementById('clearBtn').disabled = false; document.getElementById('message').innerText = r.message; document.getElementById('message').style.color='#2e7d32';}).withFailureHandler(function(e){document.getElementById('syncBtn').disabled = false; document.getElementById('clearBtn').disabled = false; document.getElementById('message').innerText = 'エラー: '+e.message; document.getElementById('message').style.color='#c62828';})[type==='sync'?'serverSyncCalendar':'serverClearCalendar'](sheet, name, email);
          }
        </script>
      </body>
    </html>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(htmlStr).setWidth(320).setHeight(300), 'カレンダー同期');
}

function getShiftSheetNamesForCalendar() { return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName()).filter(n => n.includes('シフト') && !n.includes('バイバイ')); }
function getStaffList() { return CAL_CONFIG.STAFF_LIST; }

function serverClearCalendar(sheetName, staffName, staffEmail) {
  const cal = CalendarApp.getCalendarById(staffEmail);
  if (!cal) throw new Error("アクセス権限がありません。");
  const { year, month } = getYearMonthFromSheetName(sheetName);
  const { sod, eod } = monthBoundsCal(year, month);
  const calendarId = resolveCalendarIdCal(cal, staffEmail);
  let deleted = 0, ops = 0;
  const evs = cal.getEvents(sod, eod);
  for (const ev of evs) {
    const title = String(ev.getTitle() || '');
    if (title === '予定あり' || title === '×' || isManagedCalendarEventCal(ev, year, month)) {
      deleteEventSafeCal(ev);
      deleted++;
      ops = afterOpPauseCal(ops);
    }
  }
  deleted += deleteApiManagedEventsCal(calendarId, sod, eod, year, month);
  return { success: true, message: `${staffName}さんの${month}月の同期予定を\n${deleted}件 クリアしました\n※再度「登録」で反映` };
}

function parseMemoForTasksCal(memoStr) {
  if (!memoStr) return { timeTasks: [], textMemos: [] };
  
  const timeTasks = [];
  const textMemos = [];
  
  const parts = memoStr.split(/[\/\n]/);
  const normalize = (s) => String(s).replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/[：:]/g, ':').replace(/[－ー〜−‐—]/g, '-');
  const timeRegex = /(\d{1,2})(?::?(\d{2}))?\s*-\s*(\d{1,2})(?::?(\d{2}))?/;

  for (let p of parts) {
    p = p.trim();
    if (!p) continue;
    
    const normP = normalize(p);
    const match = normP.match(timeRegex);
    
    if (match) {
      const startH = parseInt(match[1], 10), startM = parseInt(match[2] || 0, 10);
      const endH = parseInt(match[3], 10), endM = parseInt(match[4] || 0, 10);
      
      let title = normP.replace(match[0], '').trim();
      title = title.replace(/\s+/g, ' '); 
      if (!title) title = "タスク";
      
      timeTasks.push({ title, startH, startM, endH, endM, original: p });
    } else {
      textMemos.push(p);
    }
  }
  return { timeTasks, textMemos };
}

/** スタッフ名の正規化（照合用） */
function normalizeStaffKeyCal(name) {
  return String(name || "").replace(/[\s　]/g, "");
}

/** STAFF_WEEKLY_OVERRIDE のルール取得 */
function findStaffWeeklyOverrideCal(staffName) {
  const overrides = CAL_CONFIG.STAFF_WEEKLY_OVERRIDE;
  if (!overrides) return null;
  const tName = normalizeStaffKeyCal(staffName);
  for (const key in overrides) {
    const n = normalizeStaffKeyCal(key);
    if (n && (n === tName || tName.includes(n) || n.includes(tName))) {
      return overrides[key];
    }
  }
  return null;
}

/** シフト表のセルが「休み」扱いか（火水の固定勤務日でシート優先するとき用） */
function isCalendarSheetOffDayCal(rawShift) {
  const s = String(rawShift || "").trim();
  if (!s) return false;
  if (CAL_CONFIG.RE.PTO.test(s)) return true;
  if (CAL_CONFIG.RE.OFF.test(s)) return true;
  if (s.includes("半休")) return true;
  return false;
}

/**
 * 曜日ベースのイレギュラー勤務を適用（カレンダー同期のみ）
 * 石田さん例: 火11-17 / 水13-17、それ以外は不在、火水はシートが休みなら休み
 */
function applyStaffWeeklyOverrideCal(staffName, year, month, day, rawShift, rawMemo) {
  const rule = findStaffWeeklyOverrideCal(staffName);
  if (!rule) return { rawShift: rawShift, rawMemo: rawMemo };

  const dow = new Date(year, month - 1, day).getDay();
  const sheetShift = String(rawShift || "").trim();
  const weekdayShifts = rule.weekdayShifts || {};

  if (weekdayShifts[dow] !== undefined) {
    if (rule.respectSheetOff !== false && isCalendarSheetOffDayCal(sheetShift)) {
      return { rawShift: sheetShift, rawMemo: rawMemo };
    }
    return { rawShift: weekdayShifts[dow], rawMemo: rawMemo };
  }

  return { rawShift: rule.offLabel || "公休", rawMemo: rawMemo };
}

function serverSyncCalendar(sheetName, staffName, staffEmail) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  const cal = CalendarApp.getCalendarById(staffEmail);
  if (!sh || !cal) throw new Error("シートが見つからないか、カレンダー権限がありません。");
  const calendarId = resolveCalendarIdCal(cal, staffEmail);

  const { year, month } = getYearMonthFromSheetName(sheetName);
  const daysInMonth = new Date(year, month, 0).getDate();
  const lastCol = sh.getLastColumn();
  const data = sh.getRange(1, 1, Math.max(5, sh.getLastRow()), lastCol).getValues();
  
  const dateStartIdx = 2; 

  const mergedStaffRowData = new Array(daysInMonth).fill("");
  const mergedMemoRowData = new Array(daysInMonth).fill("");
  let found = false;

  // ★自動合算を廃止：設定からスタッフごとのメモ行のオフセット（相対位置）を取得
  let memoOffset = 1; // デフォルトは1行下
  const staffConfig = CAL_CONFIG.STAFF_LIST.find(s => s.name === staffName || staffName.includes(s.name));
  if (staffConfig && staffConfig.memoOffset) {
    memoOffset = staffConfig.memoOffset;
  }

  for (let i = 2; i < data.length; i++) {
    const sName = String(data[i][0]).replace(/[\s　]/g, ""); 
    const tName = staffName.replace(/[\s　]/g, "");          
    if (sName && (sName === tName || tName.includes(sName) || sName.includes(tName))) {
      found = true;
      for (let d = 0; d < daysInMonth; d++) {
        // 1. 基本シフトの取得
        const cellValue = String(data[i][dateStartIdx + d] || "").trim();
        mergedStaffRowData[d] = cellValue;
        
        // 2. メモ・タスクの取得（指定されたオフセット行を読む）
        if (i + memoOffset < data.length) {
          const memoValue = String(data[i + memoOffset][dateStartIdx + d] || "").trim();
          mergedMemoRowData[d] = memoValue;
        }
      }
      break; // ★1回見つけたら終了（予期せぬ行の巻き込みを防ぐ）
    }
  }
  
  if (!found) throw new Error(`A列に見つかりません`);

  let created = 0, skipped = 0, ops = 0;
  const apiWarnings = [];

  for (let d = 0; d < daysInMonth; d++) {
    const day = d + 1;
    let rawShift = mergedStaffRowData[d];
    const rawMemo = mergedMemoRowData[d];

    const adjusted = applyStaffWeeklyOverrideCal(staffName, year, month, day, rawShift, rawMemo);
    rawShift = adjusted.rawShift;
    
    if (!rawShift && !rawMemo) continue;

    const {sod, eod} = dayBoundsCal(year, month, day);
    const parsedTasks = parseMemoForTasksCal(rawMemo);
    
    let extraDesc = "";
    if (parsedTasks.textMemos.length > 0) {
      extraDesc = "\n\n【タスク・メモ】\n・" + parsedTasks.textMemos.join("\n・");
    }

    let storeLabel = getStoreLocationLabelCal(rawMemo, staffName);
    const colors = getCalColorsCal();
    let shiftColor = colors.WORK;
    if (storeLabel === 'ひばりが丘') {
      shiftColor = colors.WORK_HIBARI || CalendarApp.EventColor.GRAY;
    }

    // ==========================================
    // 1. 基本シフト行の処理
    // ==========================================
    if (rawShift) {
      if (CAL_CONFIG.RE.PTO.test(rawShift)) {
        const key = makeSyncKeyCal(year, month, day, 'PTO');
        purgeSyncedEventsInRangeCal(cal, calendarId, sod, eod, key);
        if (!createStatusAllDayEventCal(calendarId, sod, key, rawShift + extraDesc, apiWarnings)) {
          const ev = createAllDayEventSafeCal(cal, getStatusTitleCal(), sod, { description: makeDescriptionCal(key, rawShift + extraDesc, getTZCal()) });
          applyEventAppearanceCal(ev, 'PTO');
        }
        created++; ops = afterOpPauseCal(ops);
      }
      else if (rawShift.includes("半休")) {
        const range = parseTimeRangeCal(rawShift);
        if (range) {
          const { start, end } = buildDateRangeCal(year, month, day, range[0], range[1], range[2], range[3]);
          const r = syncShiftToCalendarCal(cal, staffEmail, start, end, year, month, day, rawShift, extraDesc, storeLabel, shiftColor, apiWarnings);
          if (r.created) { created++; ops = afterOpPauseCal(ops); } else skipped++;
        }
      }
      else if (CAL_CONFIG.RE.OFF.test(rawShift)) {
        const key = makeSyncKeyCal(year, month, day, 'OFF');
        purgeSyncedEventsInRangeCal(cal, calendarId, sod, eod, key);
        if (!createStatusAllDayEventCal(calendarId, sod, key, rawShift + extraDesc, apiWarnings)) {
          const ev = createAllDayEventSafeCal(cal, getStatusTitleCal(), sod, { description: makeDescriptionCal(key, rawShift + extraDesc, getTZCal()) });
          applyEventAppearanceCal(ev, 'OFF');
        }
        created++; ops = afterOpPauseCal(ops);
      }
      else if (CAL_CONFIG.RE.TASK.test(rawShift) && !parseTimeRangeCal(rawShift)) {
        const key = makeSyncKeyCal(year, month, day, 'TASK_ALLDAY');
        if (hasEventWithKeyCal(cal, sod, eod, key, calendarId)) { skipped++; }
        else {
          const ev = createAllDayEventSafeCal(cal, 'タスク(時間未定)', sod, { description: makeDescriptionCal(key, rawShift + extraDesc, getTZCal()) });
          applyEventAppearanceCal(ev, 'UNDEF');
          created++; ops = afterOpPauseCal(ops);
        }
      }
      else {
        const range = parseTimeRangeCal(rawShift);
        if (range) {
          const { start, end } = buildDateRangeCal(year, month, day, range[0], range[1], range[2], range[3]);
          const r = syncShiftToCalendarCal(cal, staffEmail, start, end, year, month, day, rawShift, extraDesc, storeLabel, shiftColor, apiWarnings);
          if (r.created) { created++; ops = afterOpPauseCal(ops); } else skipped++;
        }
      }
    } else if (extraDesc) {
      const key = makeSyncKeyCal(year, month, day, 'MEMO_ONLY');
      if (hasEventWithKeyCal(cal, sod, eod, key, calendarId)) { skipped++; }
      else {
         const ev = createAllDayEventSafeCal(cal, parsedTasks.textMemos.join(" / "), sod, { description: makeDescriptionCal(key, rawMemo, getTZCal()) });
         applyEventAppearanceCal(ev, 'TASK');
         created++; ops = afterOpPauseCal(ops);
      }
    }

    // ==========================================
    // 2. メモ行の時間指定タスクの処理（別予定として作成）
    // ==========================================
    let taskIdx = 0;
    for (const t of parsedTasks.timeTasks) {
      const taskKey = makeSyncKeyCal(year, month, day, 'TSK_' + taskIdx);
      taskIdx++;
      const { start, end } = buildDateRangeCal(year, month, day, t.startH, t.startM, t.endH, t.endM);
      const title = t.title; 

      if (!hasEventWithKeyCal(cal, startOfDayCal(start), endOfDayCal(start), taskKey, calendarId) && !existsSameSpanAndTitleCal(cal, start, end, title)) {
        const ev = createEventSafeCal(cal, title, start, end, { description: makeDescriptionCal(taskKey, t.original, getTZCal()) });
        applyEventAppearanceCal(ev, 'TASK');
        created++; ops = afterOpPauseCal(ops);
      } else skipped++;
    }
  }
  let msg = `${staffName}さんの予定を作成しました！\n（新規: ${created}件 / ｽｷｯﾌﾟ: ${skipped}件）`;
  if (apiWarnings.length) msg += '\n\n⚠ ' + apiWarnings.join('\n⚠ ');
  return { success: true, message: msg };
}

/** 現在の配色テーマ（CAL_CONFIG.COLOR_THEMES） */
function getCalColorsCal() {
  const theme = (CAL_CONFIG.DISPLAY && CAL_CONFIG.DISPLAY.COLOR_THEME) || 'classic';
  return CAL_CONFIG.COLOR_THEMES[theme] || CAL_CONFIG.COLOR_THEMES.classic;
}

/** 初谷さん型の見た目を予定に適用（内容・SYNC_KEY は変更しない） */
function applyEventAppearanceCal(ev, kind, shiftColor) {
  const colors = getCalColorsCal();
  safeCal(function () {
    ev.setVisibility(CalendarApp.Visibility.PUBLIC);
    switch (kind) {
      case 'SHIFT':
        // 元仕様: 勤務帯は予約ページで予約可能。勤務外は outOfOffice、会議は TASK で不可
        ev.setTransparency(CalendarApp.EventTransparency.TRANSPARENT);
        if (shiftColor) ev.setColor(shiftColor);
        break;
      case 'OFF':
        if (colors.OFF) ev.setColor(colors.OFF);
        break;
      case 'PTO':
        if (colors.PTO) ev.setColor(colors.PTO);
        else if (colors.UNDEF) ev.setColor(colors.UNDEF);
        break;
      case 'TASK':
        if (colors.TASK) ev.setColor(colors.TASK);
        ev.setTransparency(CalendarApp.EventTransparency.OPAQUE);
        break;
      case 'UNDEF':
        if (colors.UNDEF) ev.setColor(colors.UNDEF);
        break;
      case 'BLOCK':
        if (colors.BLOCK) ev.setColor(colors.BLOCK);
        ev.setTransparency(CalendarApp.EventTransparency.OPAQUE);
        break;
    }
  });
}

function getStoreLocationLabelCal(rawMemo, staffName) {
  const m = String(rawMemo || '');
  if (m.includes('ひばり')) return 'ひばりが丘';
  if (m.includes('経堂')) return '経堂';
  const tName = String(staffName || '').replace(/[\s　]/g, '');
  const staff = CAL_CONFIG.STAFF_LIST.find(function (s) {
    const n = String(s.name || '').replace(/[\s　]/g, '');
    return n && (n === tName || tName.includes(n) || n.includes(tName));
  });
  if (staff && staff.defaultStore) return staff.defaultStore;
  return CAL_CONFIG.DISPLAY.DEFAULT_STORE || '経堂';
}

function resolveCalendarIdCal(cal, staffEmail) {
  try {
    const id = cal.getId();
    if (id) return id;
  } catch (_) {}
  return String(staffEmail || '').trim();
}

/** 同期スクリプトが作った予定か（クリア用） */
function isManagedCalendarEventCal(ev, year, month) {
  const desc = ev.getDescription() || '';
  const stampPrefix = `[SYNC_KEY:${year}-${('0' + month).slice(-2)}-`;
  if (desc.indexOf(stampPrefix) !== -1) return true;
  if (desc.indexOf('[SYNC_KEY:') === -1) return false;
  const title = String(ev.getTitle() || '');
  return /^(勤務 |予定あり|不在|×|勤務場所:|公休|有休|タスク\(時間未定\)|タスク)/.test(title)
    || title === getStatusTitleCal()
    || title === '経堂' || title === 'ひばりが丘';
}

/** Calendar API で作った予定をリストから削除（description 不可の種別は extendedProperties で判定） */
function deleteApiManagedEventsCal(calendarId, sod, eod, year, month) {
  if (typeof Calendar === 'undefined' || !Calendar.Events) return 0;
  const monthPrefix = `${year}-${('0' + month).slice(-2)}-`;
  const stampPrefix = `[SYNC_KEY:${monthPrefix}`;
  let deleted = 0;
  try {
    const list = Calendar.Events.list(calendarId, {
      timeMin: sod.toISOString(),
      timeMax: eod.toISOString(),
      singleEvents: true,
      orderBy: 'startTime',
      maxResults: 2500,
    });
    const items = (list.items || []).slice();
    for (const item of items) {
      if (!isApiManagedEventItemCal(item, monthPrefix, stampPrefix)) continue;
      try {
        Calendar.Events.remove(calendarId, item.id);
        deleted++;
      } catch (_) {}
    }
  } catch (e) {
    Logger.log('deleteApiManagedEventsCal: ' + e);
  }
  return deleted;
}

function isApiManagedEventItemCal(item, monthPrefix, stampPrefix) {
  const desc = item.description || '';
  if (desc.indexOf(stampPrefix) !== -1) return true;
  const priv = (item.extendedProperties && item.extendedProperties.private) || {};
  const k = String(priv.gasSyncKey || '');
  const k2 = String(priv.gasSyncKey2 || '');
  if (k.indexOf(monthPrefix) === 0 || k2.indexOf(monthPrefix) === 0) return true;
  if (priv.gasSyncSource === 'kyodo-gas') return true;
  return false;
}

function toRfc3339Cal(dt) {
  return Utilities.formatDate(dt, getTZCal(), "yyyy-MM-dd'T'HH:mm:ss");
}

/**
 * 勤務シフトをカレンダーへ反映
 * location_bar: 棒線（workingLocation）のみ。会議は別途 TASK として不透明ブロック
 * block: 従来の「勤務」色ブロック＋勤務場所
 */
function syncShiftToCalendarCal(cal, staffEmail, start, end, y, m, d, rawShift, extraDesc, storeLabel, shiftColor, apiWarnings) {
  const shiftKey = makeSyncKeyCal(y, m, d, 'SHIFT');
  const workLocKey = makeSyncKeyCal(y, m, d, 'WORK_LOC');
  const title = `勤務 ${fmtRangeCal(start, end)}`;
  const style = (CAL_CONFIG.DISPLAY && CAL_CONFIG.DISPLAY.SHIFT_STYLE) || 'location_bar';
  const dayStart = startOfDayCal(start);
  const dayEnd = endOfDayCal(start);
  const srcText = String(rawShift || '') + String(extraDesc || '');
  const calendarId = resolveCalendarIdCal(cal, staffEmail);
  const warnings = apiWarnings || [];

  removeLegacyShiftBlockCal(cal, dayStart, dayEnd, shiftKey);

  if (hasEventWithKeyCal(cal, dayStart, dayEnd, workLocKey, calendarId)) {
    // 勤務場所はそのまま、勤務外だけ「予定あり」→「不在」に更新
    afterShiftExtrasCal(cal, calendarId, start, end, storeLabel, y, m, d, shiftKey, srcText, warnings);
    return { created: false, skipped: true };
  }
  if (style === 'block') {
    if (hasEventWithKeyCal(cal, dayStart, dayEnd, shiftKey, calendarId) || existsSameSpanAndTitleCal(cal, start, end, title)) {
      return { created: false, skipped: true };
    }
    const ev = createEventSafeCal(cal, title, start, end, {
      description: makeDescriptionCal(shiftKey, srcText, getTZCal()),
    });
    applyEventAppearanceCal(ev, 'SHIFT', shiftColor);
  }

  afterShiftExtrasCal(cal, calendarId, start, end, storeLabel, y, m, d, shiftKey, srcText, warnings);
  return { created: true, skipped: false };
}

/** location_bar 時に残った旧「勤務」塗りブロックを削除 */
function removeLegacyShiftBlockCal(cal, dayStart, dayEnd, shiftKey) {
  const style = (CAL_CONFIG.DISPLAY && CAL_CONFIG.DISPLAY.SHIFT_STYLE) || 'location_bar';
  if (style !== 'location_bar') return;
  cal.getEvents(dayStart, dayEnd).forEach(function (ev) {
    const t = String(ev.getTitle() || '');
    const desc = ev.getDescription() || '';
    if (t.indexOf('勤務 ') === 0 && desc.indexOf('[SYNC_KEY:' + shiftKey + ']') !== -1) {
      deleteEventSafeCal(ev);
    }
  });
}

function getWorkingLocationSummaryCal(locationLabel) {
  const prefix = (CAL_CONFIG.DISPLAY && CAL_CONFIG.DISPLAY.LOCATION_TITLE_PREFIX) || '勤務場所: ';
  return prefix + locationLabel;
}

function afterShiftExtrasCal(cal, calendarId, shiftStart, shiftEnd, storeLabel, y, m, d, shiftKey, shiftDesc, apiWarnings) {
  ensureOffDutyBlocksCal(cal, calendarId, shiftStart, shiftEnd, apiWarnings);
  ensureWorkingLocationCal(cal, calendarId, shiftStart, shiftEnd, storeLabel, y, m, d, shiftKey, shiftDesc, apiWarnings);
}

/** 勤務外ブロック（初谷さん型: outOfOffice＝斜線＋予約不可） */
function ensureOffDutyBlocksCal(cal, calendarId, shiftStart, shiftEnd, apiWarnings) {
  if (!CAL_CONFIG.DISPLAY.USE_OFF_BLOCKS) return;
  const sod = startOfDayCal(shiftStart);
  const eod = endOfDayCal(shiftEnd);
  const y = shiftStart.getFullYear();
  const m = shiftStart.getMonth() + 1;
  const d = shiftStart.getDate();

  if ((shiftStart - sod) / 60000 >= CAL_CONFIG.MIN_BLOCK_MINUTES) {
    const key = makeSyncKeyCal(y, m, d, 'OFF_BEFORE');
    purgeSyncedEventsInRangeCal(cal, calendarId, sod, shiftStart, key);
    createOffDutyBlockCal(cal, calendarId, sod, shiftStart, key, apiWarnings);
  }
  if ((eod - shiftEnd) / 60000 >= CAL_CONFIG.MIN_BLOCK_MINUTES) {
    const key = makeSyncKeyCal(y, m, d, 'OFF_AFTER');
    purgeSyncedEventsInRangeCal(cal, calendarId, shiftEnd, eod, key);
    createOffDutyBlockCal(cal, calendarId, shiftEnd, eod, key, apiWarnings);
  }
}

/** 公休・有休・勤務外ブロックなどの統一タイトル */
function getStatusTitleCal() {
  return (CAL_CONFIG.DISPLAY && CAL_CONFIG.DISPLAY.STATUS_TITLE) || '不在';
}

/** 終日の公休・有休 → outOfOffice「不在」（勤務外ブロックと同じ見た目） */
function createStatusAllDayEventCal(calendarId, sod, syncKey, note, apiWarnings) {
  assertCalendarApiCal();
  const tz = getTZCal();
  const endDay = new Date(sod.getFullYear(), sod.getMonth(), sod.getDate() + 1);
  try {
    withRetryCal(function () {
      return Calendar.Events.insert({
        summary: getStatusTitleCal(),
        eventType: 'outOfOffice',
        start: { date: Utilities.formatDate(sod, tz, 'yyyy-MM-dd') },
        end: { date: Utilities.formatDate(endDay, tz, 'yyyy-MM-dd') },
        visibility: 'public',
        transparency: 'opaque',
        extendedProperties: { private: gasSyncPrivatePropsCal(syncKey, null, note) }
      }, calendarId);
    }, 'outOfOfficeAllDay');
    return true;
  } catch (e) {
    const msg = '終日不在の作成失敗: ' + e;
    Logger.log(msg);
    if (apiWarnings) apiWarnings.push(msg);
    return false;
  }
}

function createOffDutyBlockCal(cal, calendarId, start, end, syncKey, apiWarnings) {
  const tz = getTZCal();
  assertCalendarApiCal();
  try {
    withRetryCal(function () {
      return Calendar.Events.insert({
        summary: getStatusTitleCal(),
        eventType: 'outOfOffice',
        start: { dateTime: toRfc3339Cal(start), timeZone: tz },
        end: { dateTime: toRfc3339Cal(end), timeZone: tz },
        visibility: 'public',
        transparency: 'opaque',
        extendedProperties: { private: gasSyncPrivatePropsCal(syncKey, null, 'off_block') }
      }, calendarId);
    }, 'outOfOffice');
  } catch (e) {
    const msg = '勤務外（斜線）作成失敗: ' + e;
    Logger.log(msg);
    if (apiWarnings) apiWarnings.push(msg);
  }
}

/** 笠原さん型: 勤務場所アイコン（勤務場所: 経堂 など） */
function ensureWorkingLocationCal(cal, calendarId, shiftStart, shiftEnd, locationLabel, y, m, d, shiftKey, shiftDesc, apiWarnings) {
  if (!CAL_CONFIG.DISPLAY.USE_WORKING_LOCATION) return;
  const workLocKey = makeSyncKeyCal(y, m, d, 'WORK_LOC');
  const dayStart = startOfDayCal(shiftStart);
  const dayEnd = endOfDayCal(shiftEnd);
  if (hasEventWithKeyCal(cal, dayStart, dayEnd, workLocKey, calendarId)) return;

  const tz = getTZCal();
  const summary = getWorkingLocationSummaryCal(locationLabel);
  assertCalendarApiCal();
  try {
    withRetryCal(function () {
      return Calendar.Events.insert({
        summary: summary,
        eventType: 'workingLocation',
        start: { dateTime: toRfc3339Cal(shiftStart), timeZone: tz },
        end: { dateTime: toRfc3339Cal(shiftEnd), timeZone: tz },
        visibility: 'public',
        transparency: 'transparent',
        extendedProperties: { private: gasSyncPrivatePropsCal(workLocKey, shiftKey, locationLabel) },
        workingLocationProperties: {
          type: 'officeLocation',
          officeLocation: { label: locationLabel }
        }
      }, calendarId);
    }, 'workingLocation');
  } catch (e) {
    const msg = '勤務場所（' + locationLabel + '）作成失敗: ' + e;
    Logger.log(msg);
    if (apiWarnings) apiWarnings.push(msg);
  }
}

function assertCalendarApiCal() {
  if (typeof Calendar === 'undefined' || !Calendar.Events) {
    throw new Error('Google Calendar API（拡張サービス ID: Calendar）が未設定です。エディタの「サービス」から追加してください。');
  }
}

function withRetryCal(fn, label) {
  let delay = CAL_CONFIG.BACKOFF_BASE_MS;
  for (let i = 0; i < CAL_CONFIG.MAX_RETRY; i++) {
    try { const res = fn(); if (CAL_CONFIG.THROTTLE_MS) Utilities.sleep(CAL_CONFIG.THROTTLE_MS); return res; }
    catch (e) { if (/creating or deleting too many calendars?|too many .*events|Service invoked too many times|Rate Limit/i.test(String(e))) { Utilities.sleep(delay + Math.floor(Math.random()*200)); delay *= 2; continue; } throw e; }
  }
  throw new Error(`${label} failed due to rate limit`);
}

function createEventSafeCal(cal, t, s, e, o) { return withRetryCal(function () { return cal.createEvent(t, s, e, o); }, 'createEvent'); }
function createAllDayEventSafeCal(cal, t, d, o) { return withRetryCal(function () { return cal.createAllDayEvent(t, d, o); }, 'createAllDayEvent'); }
function deleteEventSafeCal(ev) { return withRetryCal(function () { ev.deleteEvent(); return true; }, 'deleteEvent'); }
function safeCal(fn){ try { fn(); } catch(_) {} }
function afterOpPauseCal(ops) { ops++; if (ops % CAL_CONFIG.PAUSE_EVERY_N_EVENTS === 0) Utilities.sleep(CAL_CONFIG.PAUSE_MS); return ops; }
function getYearMonthFromSheetName(name) { let y = new Date().getFullYear(), m = new Date().getMonth() + 1, mt = name.match(/(\d{4})年(\d{1,2})月/); if (mt) { y = parseInt(mt[1], 10); m = parseInt(mt[2], 10); } return { year: y, month: m }; }
function monthBoundsCal(y, m){ return {sod: new Date(y, m-1, 1, 0,0,0,0), eod: new Date(y, m, 0, 23,59,59,999)}; }
function dayBoundsCal(y, m, d) { return { sod: new Date(y, m - 1, d, 0, 0, 0, 0), eod: new Date(y, m - 1, d, 23, 59, 59, 999) }; }
function startOfDayCal(dt){ return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate(), 0,0,0,0); }
function endOfDayCal(dt){ return new Date(dt.getFullYear(), dt.getMonth(), dt.getDate(), 23,59,59,999); }
function getTZCal(){ return Session.getScriptTimeZone() || 'Asia/Tokyo'; }
function makeSyncKeyCal(y, m, d, type) { return `${y}-${('0'+m).slice(-2)}-${('0'+d).slice(-2)}:${type}`; }
function makeDescriptionCal(key, ex, tz) { return `[SYNC_KEY:${key}]\n原文: ${ex}\n同期: ${Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss')}`; }

/** outOfOffice / workingLocation は description 禁止のため private 拡張プロパティに SYNC_KEY を保存 */
function gasSyncPrivatePropsCal(primaryKey, secondaryKey, note) {
  const p = {
    gasSyncKey: primaryKey,
    gasSyncSource: 'kyodo-gas',
    gasSyncedAt: Utilities.formatDate(new Date(), getTZCal(), 'yyyy/MM/dd HH:mm:ss'),
  };
  if (secondaryKey) p.gasSyncKey2 = secondaryKey;
  if (note) p.gasSyncNote = String(note).slice(0, 200);
  return p;
}

/** 同じ SYNC_KEY の予定と旧タイトル「予定あり」を削除してから作り直す */
function purgeSyncedEventsInRangeCal(cal, calendarId, rangeStart, rangeEnd, syncKey) {
  const legacyBusy = ['予定あり', '×'];
  const statusTitle = getStatusTitleCal();
  cal.getEvents(rangeStart, rangeEnd).forEach(function (ev) {
    const t = String(ev.getTitle() || '');
    const desc = ev.getDescription() || '';
    if (desc.indexOf('[SYNC_KEY:' + syncKey + ']') !== -1) {
      deleteEventSafeCal(ev);
      return;
    }
    if (legacyBusy.indexOf(t) !== -1) {
      deleteEventSafeCal(ev);
      return;
    }
    if (t === statusTitle || t === '不在' || t === '公休' || t === '有休') {
      if (desc.indexOf('[SYNC_KEY:') !== -1) deleteEventSafeCal(ev);
    }
  });
  purgeSyncedEventsInRangeViaApiCal(calendarId, rangeStart, rangeEnd, syncKey);
}

function purgeSyncedEventsInRangeViaApiCal(calendarId, rangeStart, rangeEnd, syncKey) {
  if (!calendarId || typeof Calendar === 'undefined' || !Calendar.Events) return;
  try {
    const list = Calendar.Events.list(calendarId, {
      timeMin: rangeStart.toISOString(),
      timeMax: rangeEnd.toISOString(),
      singleEvents: true,
      maxResults: 100,
    });
    (list.items || []).forEach(function (item) {
      const priv = (item.extendedProperties && item.extendedProperties.private) || {};
      const k = String(priv.gasSyncKey || '');
      const k2 = String(priv.gasSyncKey2 || '');
      const summary = String(item.summary || '');
      if (k === syncKey || k2 === syncKey) {
        try { Calendar.Events.remove(calendarId, item.id); } catch (_) {}
        return;
      }
      if (summary === '予定あり' || summary === '×') {
        if (priv.gasSyncSource === 'kyodo-gas' || item.eventType === 'outOfOffice') {
          try { Calendar.Events.remove(calendarId, item.id); } catch (_) {}
        }
      }
    });
  } catch (e) {
    Logger.log('purgeSyncedEventsInRangeViaApiCal: ' + e);
  }
}

function hasSyncKeyViaApiCal(calendarId, f, t, key) {
  if (!calendarId || typeof Calendar === 'undefined' || !Calendar.Events) return false;
  try {
    const r = Calendar.Events.list(calendarId, {
      timeMin: f.toISOString(),
      timeMax: t.toISOString(),
      singleEvents: true,
      privateExtendedProperty: 'gasSyncKey=' + key,
      maxResults: 5,
    });
    if (r.items && r.items.length) return true;
    const r2 = Calendar.Events.list(calendarId, {
      timeMin: f.toISOString(),
      timeMax: t.toISOString(),
      singleEvents: true,
      privateExtendedProperty: 'gasSyncKey2=' + key,
      maxResults: 5,
    });
    return !!(r2.items && r2.items.length);
  } catch (e) {
    Logger.log('hasSyncKeyViaApiCal: ' + e);
    return false;
  }
}

function hasEventWithKeyCal(cal, f, t, key, calendarId) {
  const cid = calendarId || resolveCalendarIdCal(cal, '');
  if (hasSyncKeyViaApiCal(cid, f, t, key)) return true;
  return cal.getEvents(f, t).some(function (ev) {
    return (ev.getDescription() || '').indexOf('[SYNC_KEY:' + key + ']') !== -1;
  });
}
function existsSameSpanAndTitleCal(cal, s, e, t) { return cal.getEvents(s, e).some(ev => ev.getStartTime().getTime() === s.getTime() && ev.getEndTime().getTime() === e.getTime() && (ev.getTitle() || '') === t); }

// ★修正：エラーの原因だった「inputStr」を正しい引数名「s」に修正しました
function parseTimeRangeCal(s) {
  const src = String(s||'').replace(/\r?\n/g, ' ').replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0)).replace(/[：:]/g, ':').replace(/[－ー〜−‐—]/g, '-').replace(/\s+/g, ' ').trim();
  let m = src.match(/(\d{1,2})(?::?(\d{2}))?\s*-\s*(\d{1,2})(?::?(\d{2}))?/);
  if (m) return [parseInt(m[1],10), parseInt(m[2]||0,10), parseInt(m[3],10), parseInt(m[4]||0,10)];
  return null;
}

function buildDateRangeCal(y, m, d, sh, sm, eh, em) {
  let s = new Date(y, m - 1, d, sh === 24 ? 0 : sh, sm || 0), eD = d, eH = eh;
  if (eh === 24) { eH = 0; eD += 1; }
  let e = new Date(y, m - 1, eD, eH, em || 0);
  if (e <= s) e = new Date(y, m - 1, eD + 1, eH, em || 0);
  return { start: s, end: e };
}
function fmtRangeCal(s, e) { return `${('0'+s.getHours()).slice(-2)}:${('0'+s.getMinutes()).slice(-2)}〜${('0'+e.getHours()).slice(-2)}:${('0'+e.getMinutes()).slice(-2)}`; }