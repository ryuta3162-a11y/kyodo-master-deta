/**
 * JOYFIT24経堂 — オプション契約メール集計
 *
 * 【運用スプレッドシート】
 * https://docs.google.com/spreadsheets/d/14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w/edit?gid=1774229739#gid=1774229739
 * （リンク集: リポジトリ マスタ/各種リンク.md）
 */

/** 運用ブック ID（マスタ/各種リンク.md と同期） */
const OPTION_SPREADSHEET_ID = "14hxiLBzvGTuIpfZcoVjiHpz8b419OzUrtQAr5788h3w";

// 検索するメールの条件
const SEARCH_QUERY = 'subject:("【JOYFIT24経堂】オプションご契約につきまして" OR "【JOYFIT24経堂】ご入会ありがとうございます")';

const SHEET_NAME_LOG = "OPデータ";
const SHEET_NAME_SUMMARY = "集計";
const OP_LOG_HEADERS = [
  "受信日時",
  "氏名",
  "区分",
  "オプション名(メール記載)",
  "オプション名(集計用)",
  "メールID"
];

/**
 * スプレッドシートを開いたときに実行される関数
 * メニューは2つだけ（全部更新 / 日報だけ）
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('JOYFIT')
    .addItem('① 全部更新する', 'smartUpdateData')
    .addItem('② 日報の数値だけ更新する', 'refreshNippoCurrentMonthFromLog')
    .addToUi();
}

/**
 * 時間トリガー用（Gmail取得＋集計＋日報まで一括）
 * 手動で日報だけ直すときは refreshNippoCurrentMonthFromLog
 */
function installOptionDailyTrigger_() {
  const ui = SpreadsheetApp.getUi();
  removeOptionDailyTriggers_();
  ScriptApp.newTrigger("smartUpdateData")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();
  ui.alert(
    "トリガー設定",
    "毎日 7:00〜8:00 に「smartUpdateData」を実行するトリガーを設定しました。\n\n" +
      "内容: Gmail取得 → OPデータ → 集計（B1の月）→ 日報C列（B1と同じ月なら）\n\n" +
      "確認: Apps Script → 左の時計アイコン「トリガー」",
    ui.ButtonSet.OK
  );
}

function removeOptionDailyTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    if (fn === "smartUpdateData" || fn === "refreshNippoCurrentMonthFromLog") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

const OPTION_LIST = [
  "安心サポート", "安心サポートVIP", "水素水", "オンラインレッスン",
  "体組成計", "契約ロッカー1,500", "レンタルマット", "プロテイン12杯",
  "プロテイン無制限", "プロテイン＋水素水", "レンタルタオル", "タンニング",
  "セルフエステ", "ホットスタジオ", "ヨガロッカー", "ピラティスリフォーマー"
];

/**
 * 1. ワンボタンで実行される統合エントリーポイント
 */
function smartUpdateData() {
  // まず必ずシートのフォーマットを最新化・修復する
  setupSpreadsheet(); 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAME_LOG);
  const ui = SpreadsheetApp.getUi();

  // OPデータが空（ヘッダーしかない初回状態）の場合の自動提案
  if (logSheet.getLastRow() <= 1) {
    const response = ui.alert(
      '初回セットアップの確認', 
      '裏側の「OPデータ」が空の状態です。\n最初に過去すべての対象メールを一括取得してデータベースを構築しますか？\n\n※「はい」を強く推奨します（量によって数分かかります）\n※「いいえ」を選ぶと、B1セルで選択されている月のみを取得します。', 
      ui.ButtonSet.YES_NO
    );
    
    if (response == ui.Button.YES) {
      executeFetchAll(ss, logSheet, ui);
      return;
    }
  }

  // 通常運用（B1セルで選ばれている月のみを更新）
  executeFetchMonth(ss, logSheet);
}

/**
 * 2. シートをセットアップし、ダッシュボード型にフォーマットする
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 裏側のデータベース（OPデータ） ---
  let logSheet = ss.getSheetByName(SHEET_NAME_LOG);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_NAME_LOG);
  }
  logSheet.getRange(1, 1, 1, OP_LOG_HEADERS.length).setValues([OP_LOG_HEADERS]);
  logSheet.getRange(1, 1, 1, OP_LOG_HEADERS.length).setFontWeight("bold").setBackground("#f3f3f3");
  logSheet.setFrozenRows(1);

  // --- 表側のダッシュボード（集計） ---
  let summarySheet = ss.getSheetByName(SHEET_NAME_SUMMARY);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SHEET_NAME_SUMMARY);
  }
  
  // 過去のヘッダーの残骸を綺麗に消去する
  summarySheet.getRange("1:2").clear();

  // 1行目：月選択エリア（プルダウンリストを作成）
  summarySheet.getRange("A1").setValue("対象月を選択 ➡").setFontWeight("bold").setHorizontalAlignment("right");
  
  const monthCell = summarySheet.getRange("B1");
  
  // ★追加：過去の「日付形式」などの古い設定を一度完全にリセットする
  monthCell.clearDataValidations();
  monthCell.clearFormat(); 
  monthCell.setNumberFormat("@"); // セルを純粋な「テキスト形式」に指定

  const monthList = [];
  const today = new Date();
  // 過去2年前から1年後までの「〇〇〇〇年〇月」のリストを生成
  for (let y = today.getFullYear() - 2; y <= today.getFullYear() + 1; y++) {
    for (let m = 1; m <= 12; m++) {
      monthList.push(`${y}年${m}月`);
    }
  }
  
  // セルにプルダウン（ドロップダウン）を設定
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(monthList, true).build();
  monthCell.setDataValidation(rule);
  monthCell.setBackground("#fff2cc").setFontWeight("bold").setHorizontalAlignment("center");
  
  // 空欄や、リストにない値（2024/01/01など）が入っている場合は当月で上書きしてエラーを消す
  let currentVal = String(monthCell.getValue());
  if (!monthList.includes(currentVal)) {
    monthCell.setValue(`${today.getFullYear()}年${today.getMonth() + 1}月`);
  }

  // 2行目：ヘッダー（D=利用開始合計, E=停止, F=翌月±）
  summarySheet.getRange(2, 1, 1, 6).setValues([
    ["オプション名", "利用開始(新規入会)", "利用開始(OP追加)", "利用開始合計", "利用停止数", "翌月の±"]
  ]);
  summarySheet.getRange("A2:F2").setFontWeight("bold").setBackground("#e2efda");
  summarySheet.setFrozenRows(2);

  // 3行目以降：B/C/D/E は GAS が書き込み、F は式
  const summaryNames = OPTION_LIST.map((optName) => [optName, "", "", "", "", ""]);
  summarySheet.getRange(3, 1, summaryNames.length, 6).setValues(summaryNames);
  for (let i = 0; i < OPTION_LIST.length; i++) {
    const row = i + 3;
    summarySheet.getRange(row, 6).setFormula("=D" + row + "-E" + row);
  }
}

/**
 * 3. B1セルで選ばれている月のメールだけを再取得して更新する内部処理
 */
function executeFetchMonth(ss, logSheet) {
  const summarySheet = ss.getSheetByName(SHEET_NAME_SUMMARY);
  let targetStr = summarySheet.getRange("B1").getValue();
  
  // プルダウンの文字列（例："2026年4月"）から年と月を抽出
  let match = String(targetStr).match(/^(\d{4})年(\d{1,2})月$/);
  if (!match) {
    const today = new Date();
    targetStr = `${today.getFullYear()}年${today.getMonth() + 1}月`;
    summarySheet.getRange("B1").setValue(targetStr);
    match = [null, today.getFullYear(), today.getMonth() + 1];
  }

  const targetYear = parseInt(match[1]);
  const targetMonth = parseInt(match[2]) - 1; // 0〜11に変換
  const targetMonthStr = String(targetMonth + 1).padStart(2, '0');

  // 対象月以外のデータを保持し、対象月部分だけクリアする（部分洗い替え）
  const logData = logSheet.getDataRange().getValues();
  const header = logData.shift(); 
  
  const retainedData = logData.filter((row) => {
    const d = parseOpLogDate_(row[0]);
    if (!d) return true;
    return !(d.getFullYear() === targetYear && d.getMonth() === targetMonth);
  });

  logSheet.clearContents();
  logSheet.appendRow(header);
  if (retainedData.length > 0) {
    // 修正: 列数は配列の要素数(retainedData[0].length)を指定
    logSheet.getRange(2, 1, retainedData.length, retainedData[0].length).setValues(retainedData);
  }

  // 対象月のメール取得
  const firstDay = `${targetYear}/${targetMonthStr}/01`;
  const nextMonthDate = new Date(targetYear, targetMonth + 1, 1);
  const nextMonthFirstDay = `${nextMonthDate.getFullYear()}/${String(nextMonthDate.getMonth() + 1).padStart(2, '0')}/01`;
  
  const query = `${SEARCH_QUERY} after:${firstDay} before:${nextMonthFirstDay}`;
  const threads = searchGmailAllThreads_(query);
  const newData = extractDataFromThreads(threads, targetYear, targetMonth);

  // 新データを一括書き込み
  if (newData.length > 0) {
    // 修正: 列数は配列の要素数(newData[0].length)を指定
    logSheet.getRange(logSheet.getLastRow() + 1, 1, newData.length, newData[0].length).setValues(newData);
  }

  // 集計 B1 の月＝今回取り込んだ月 → OPデータ から B〜E を書き込み（Fは式）
  syncSummarySheetFromOpData_(ss, targetYear, targetMonth, logSheet);

  const nippoYm = resolveNippoTargetYearMonth_(ss);
  const nippoUpdated =
    nippoYm.year === targetYear && nippoYm.month === targetMonth;
  if (nippoUpdated) {
    updateNippoSheetForMonth(ss, targetYear, targetMonth, logSheet);
  }

  const ui = SpreadsheetApp.getUi();
  let doneMsg =
    (targetMonth + 1) +
    "月分を取り込み、集計シートを更新しました。";
  doneMsg += nippoUpdated
    ? "\n日報（C21:C36＝当月の利用開始合計）も更新しました。"
    : "\n日報は未更新です（日報B1の月が集計B1と違うため）。②で日報だけ更新できます。";
  ui.alert("全部更新", doneMsg, ui.ButtonSet.OK);
}

/**
 * 4. 過去すべてのメールを一括で取得し、OPデータを再構築する内部処理
 */
function executeFetchAll(ss, logSheet, ui) {
  // OPデータシートを完全にクリア（ヘッダーのみ残す）
  const logData = logSheet.getDataRange().getValues();
  const header = logData.length > 0 ? [padLogHeaderRow_(logData[0])] : [OP_LOG_HEADERS];
  logSheet.clearContents();
  logSheet.getRange(1, 1, 1, OP_LOG_HEADERS.length).setValues(header);

  // 日付指定なしで全検索
  const query = SEARCH_QUERY;
  const threads = searchGmailAllThreads_(query);
  
  // 年月指定なし(null)で全抽出
  const newData = extractDataFromThreads(threads, null, null);

  // 新データを一括書き込み（日付の古い順に並び替えてから書き込み）
  if (newData.length > 0) {
    // 修正: 日付データ(配列の0番目)でソートする
    newData.sort((a, b) => a[0].getTime() - b[0].getTime());
    // 修正: 列数は配列の要素数(newData[0].length)を指定
    logSheet.getRange(2, 1, newData.length, newData[0].length).setValues(newData);
  }
  
  ui.alert(`全データの取得完了!\n\n合計 ${newData.length} 件の過去データを構築しました。\nこれでB1セルの月を切り替えるだけで全期間の確認が可能です。`);

  const summaryYm = resolveSummaryTargetYearMonth_(ss);
  syncSummarySheetFromOpData_(ss, summaryYm.year, summaryYm.month, logSheet);

  const nippoYm = resolveNippoTargetYearMonth_(ss);
  updateNippoSheetForMonth(ss, nippoYm.year, nippoYm.month, logSheet);
}

/** Gmail 検索をページ送り（同一件名スレッドの取りこぼし軽減） */
function searchGmailAllThreads_(query) {
  const out = [];
  const pageSize = 100;
  let start = 0;
  while (start < 500) {
    const batch = GmailApp.search(query, start, pageSize);
    if (!batch.length) break;
    for (let i = 0; i < batch.length; i++) out.push(batch[i]);
    start += batch.length;
    if (batch.length < pageSize) break;
  }
  return out;
}

function padLogHeaderRow_(row) {
  const h = row.slice();
  while (h.length < OP_LOG_HEADERS.length) h.push("");
  return h.slice(0, OP_LOG_HEADERS.length);
}

function padLogDataRow_(row, msgId) {
  const r = row.slice();
  while (r.length < 5) r.push("");
  const norm = normalizeOptionName(String(r[4] || r[3] || ""));
  return [r[0], r[1], r[2], r[3], norm, msgId || r[5] || ""];
}

function getMessageBodyText_(message) {
  const plain = message.getPlainBody();
  if (plain && String(plain).trim()) return String(plain);
  const html = message.getBody();
  if (!html) return "";
  return String(html)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .trim();
}

/**
 * 5. スレッド内の「メール1通ずつ」を解析（同じ件名でまとまっても全 message を処理）
 */
function extractDataFromThreads(threads, targetYear, targetMonth) {
  const newData = [];
  const seen = {};

  for (let t = 0; t < threads.length; t++) {
    const messages = threads[t].getMessages();
    for (let m = 0; m < messages.length; m++) {
      const message = messages[m];
      const date = message.getDate();

      if (targetYear !== null && targetMonth !== null) {
        if (date.getFullYear() !== targetYear || date.getMonth() !== targetMonth) continue;
      }

      const subject = String(message.getSubject() || "");
      const body = getMessageBodyText_(message);
      const msgId = message.getId();

      try {
        let parsed = [];
        if (subject.indexOf("オプションご契約につきまして") !== -1) {
          parsed = parseOptionContractEmail(date, body, msgId);
        } else if (subject.indexOf("ご入会ありがとうございます") !== -1) {
          parsed = parseSignupEmail(date, body, msgId);
        }

        for (let i = 0; i < parsed.length; i++) {
          const row = padLogDataRow_(parsed[i], msgId);
          const key = row[5] + "|" + row[2] + "|" + row[4];
          if (seen[key]) continue;
          seen[key] = true;
          newData.push(row);
        }
      } catch (e) {
        Logger.log("メール解析エラー: " + subject + " / " + e.message);
      }
    }
  }
  return newData;
}

function extractSignupName_(body) {
  const m1 = body.match(/お名前\s*[:：]\s*(.+?)\s*様/);
  if (m1) return String(m1[1]).replace(/^[>\s]+/, "").trim();

  const lines = body.split(/\r?\n/);
  for (let i = 0; i < Math.min(lines.length, 12); i++) {
    const line = String(lines[i] || "").trim();
    if (!line || line.indexOf("JOYFIT") !== -1 || line.indexOf("ご利用開始") !== -1) continue;
    const m2 = line.match(/^(.{1,30}?)\s*様\s*$/);
    if (m2) return String(m2[1]).replace(/^[>\s]+/, "").trim();
  }
  return "";
}

function isBreakdownNoiseLine_(line) {
  const s = String(line || "").trim();
  if (!s) return true;
  if (s.indexOf("----") !== -1) return true;
  if (s.indexOf("小計") !== -1 || s.indexOf("合計") !== -1) return true;
  if (s.indexOf("初期費用") !== -1) return true;
  if (s.indexOf("ナショナル会員") !== -1) return true;
  if (s.indexOf("会費") !== -1 && s.indexOf("円") !== -1 && OPTION_LIST.indexOf(normalizeOptionName(s)) < 0) {
    const norm = normalizeOptionName(s);
    if (!OPTION_LIST.includes(norm)) return true;
  }
  return false;
}

/**
 * ご入会メール: 月会費の内訳（フル形式）
 */
function parseSignupBreakdown_(date, body, name, msgId) {
  const results = [];
  const lines = body.split(/\r?\n/);
  let inBreakdown = false;
  const got = {};

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.indexOf("月会費の内訳") !== -1) {
      inBreakdown = true;
      continue;
    }
    if (inBreakdown && (line.indexOf("APP登録方法") !== -1 || line.indexOf("ご契約中のオプション") !== -1)) {
      break;
    }
    if (!inBreakdown) continue;
    if (isBreakdownNoiseLine_(line)) continue;
    if (line.indexOf("円") === -1) continue;

    const norm = normalizeOptionName(line);
    if (!OPTION_LIST.includes(norm) || got[norm]) continue;
    got[norm] = true;
    const raw = line.replace(/\(\d+月分\).*/, "").replace(/（\d+月分）.*/, "").trim();
    results.push([date, name, "利用開始(新規入会)", raw, norm, msgId]);
  }
  return results;
}

/**
 * ご入会メール: 短い通知形式（スレッド内の1通目だけ内訳なし・オプション名(6月分)のみ等）
 */
function parseSignupLooseLines_(date, body, name, msgId, alreadyGot) {
  const results = [];
  const got = alreadyGot || {};
  const lines = body.split(/\r?\n/);

  for (let i = 0; i < lines.length; i++) {
    const line = String(lines[i] || "").trim();
    if (!line || line.indexOf("月分") === -1) continue;
    if (isBreakdownNoiseLine_(line)) continue;

    const norm = normalizeOptionName(line);
    if (!OPTION_LIST.includes(norm) || got[norm]) continue;
    got[norm] = true;
    results.push([date, name, "利用開始(新規入会)", line, norm, msgId]);
  }
  return results;
}

function parseSignupEmail(date, body, msgId) {
  const name = extractSignupName_(body);
  const breakdown = parseSignupBreakdown_(date, body, name, msgId);
  const got = {};
  for (let i = 0; i < breakdown.length; i++) got[breakdown[i][4]] = true;
  const loose = parseSignupLooseLines_(date, body, name, msgId, got);
  return breakdown.concat(loose);
}

/**
 * オプションご契約: 1通に (利用開始) が複数行ある場合もすべて拾う
 */
function parseOptionContractEmail(date, body, msgId) {
  const results = [];
  const nameMatch = body.match(/([^\r\n]{1,40}?)\s*様/);
  const name = nameMatch ? String(nameMatch[1]).replace(/^[>\s]+/, "").trim() : "";

  const re = /[（(](利用開始|利用停止)[）)]\s*([^\r\n]+)/g;
  let m;
  while ((m = re.exec(body)) !== null) {
    const status = m[1] === "利用開始" ? "利用開始(OP追加)" : m[1];
    const optionRaw = String(m[2]).trim();
    if (!optionRaw) continue;
    results.push([date, name, status, optionRaw, normalizeOptionName(optionRaw), msgId]);
  }
  return results;
}

/**
 * 日報の「当月」= 日報シート B1（例: 2606 → 2026年6月）で決める
 */
function resolveNippoTargetYearMonth_(ss) {
  const nippo = ss.getSheetByName("日報");
  if (nippo) {
    const b1 = String(nippo.getRange("B1").getDisplayValue() || nippo.getRange("B1").getValue() || "").trim();
    const yyMM = b1.match(/^(\d{2})(\d{2})$/);
    if (yyMM) {
      const y = 2000 + parseInt(yyMM[1], 10);
      const m = parseInt(yyMM[2], 10) - 1;
      if (m >= 0 && m <= 11) {
        return { year: y, month: m, label: y + "年" + (m + 1) + "月（日報B1:" + b1 + "）" };
      }
    }
    const ym = b1.match(/^(\d{4})年(\d{1,2})月$/);
    if (ym) {
      return {
        year: parseInt(ym[1], 10),
        month: parseInt(ym[2], 10) - 1,
        label: b1
      };
    }
  }

  const summary = ss.getSheetByName(SHEET_NAME_SUMMARY);
  if (summary) {
    const s = String(summary.getRange("B1").getValue() || "").trim();
    const ym = s.match(/^(\d{4})年(\d{1,2})月$/);
    if (ym) {
      return {
        year: parseInt(ym[1], 10),
        month: parseInt(ym[2], 10) - 1,
        label: s + "（集計B1）"
      };
    }
  }

  const now = new Date();
  return {
    year: now.getFullYear(),
    month: now.getMonth(),
    label: now.getFullYear() + "年" + (now.getMonth() + 1) + "月（PCの今日）"
  };
}

/** OPデータ A列の日付（Date／シリアル値／文字列）を Date に */
function parseOpLogDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (typeof v === "number" && !isNaN(v) && v > 30000) {
    const base = new Date(1899, 11, 30);
    const whole = Math.floor(v);
    const frac = v - whole;
    const d = new Date(base.getTime() + whole * 86400000 + Math.round(frac * 86400000));
    return isNaN(d.getTime()) ? null : d;
  }
  const s = String(v || "").trim();
  if (!s) return null;
  const m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
  }
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/** 集計シート B1（2026年6月）から年月を取得 */
function resolveSummaryTargetYearMonth_(ss) {
  const summary = ss.getSheetByName(SHEET_NAME_SUMMARY);
  if (summary) {
    const s = String(summary.getRange("B1").getDisplayValue() || summary.getRange("B1").getValue() || "").trim();
    const ym = s.match(/^(\d{4})年(\d{1,2})月$/);
    if (ym) {
      return {
        year: parseInt(ym[1], 10),
        month: parseInt(ym[2], 10) - 1,
        label: s
      };
    }
  }
  return resolveNippoTargetYearMonth_(ss);
}

/**
 * OPデータ を1通ずつ数える（日報・集計で共通）
 */
function countOpDataForMonth_(logSheet, targetYear, targetMonth) {
  const data = logSheet.getDataRange().getValues();
  const counts = [];
  for (let o = 0; o < OPTION_LIST.length; o++) {
    counts.push({ newSignup: 0, opAdd: 0, stop: 0 });
  }
  const idxByOpt = {};
  for (let o = 0; o < OPTION_LIST.length; o++) idxByOpt[OPTION_LIST[o]] = o;

  let monthRows = 0;
  for (let i = 1; i < data.length; i++) {
    const d = parseOpLogDate_(data[i][0]);
    if (!d || d.getFullYear() !== targetYear || d.getMonth() !== targetMonth) continue;
    monthRows++;
    const cat = String(data[i][2] || "").trim();
    const opt = String(data[i][4] || "").trim();
    const idx = idxByOpt[opt];
    if (idx === undefined) continue;
    if (cat === "利用開始(新規入会)") counts[idx].newSignup++;
    else if (cat === "利用開始(OP追加)") counts[idx].opAdd++;
    else if (cat === "利用停止") counts[idx].stop++;
  }
  return { counts: counts, monthRows: monthRows };
}

/**
 * 集計シート B〜E を OPデータ から書き込み（F列は式 =D−E）
 */
function syncSummarySheetFromOpData_(ss, targetYear, targetMonth, logSheet) {
  const summarySheet = ss.getSheetByName(SHEET_NAME_SUMMARY);
  if (!summarySheet) return;

  SpreadsheetApp.flush();
  const result = countOpDataForMonth_(logSheet, targetYear, targetMonth);
  const bCol = [];
  const cCol = [];
  const dCol = [];
  const eCol = [];
  for (let i = 0; i < OPTION_LIST.length; i++) {
    const start = result.counts[i].newSignup + result.counts[i].opAdd;
    bCol.push([result.counts[i].newSignup]);
    cCol.push([result.counts[i].opAdd]);
    dCol.push([start]);
    eCol.push([result.counts[i].stop]);
  }
  summarySheet.getRange(3, 2, OPTION_LIST.length, 1).setValues(bCol);
  summarySheet.getRange(3, 3, OPTION_LIST.length, 1).setValues(cCol);
  summarySheet.getRange(3, 4, OPTION_LIST.length, 1).setValues(dCol);
  summarySheet.getRange(3, 5, OPTION_LIST.length, 1).setValues(eCol);
  ensureSummaryFormulas_(summarySheet);
}

/** 集計 F列（翌月±）の式を揃える */
function ensureSummaryFormulas_(summarySheet) {
  for (let i = 0; i < OPTION_LIST.length; i++) {
    const row = i + 3;
    summarySheet.getRange(row, 6).setFormula("=D" + row + "-E" + row);
  }
}

function syncSummaryFromOpDataMenu_() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAME_LOG);
  if (!logSheet || logSheet.getLastRow() < 2) {
    ui.alert("集計同期", "OPデータが空です。", ui.ButtonSet.OK);
    return;
  }
  const ym = resolveSummaryTargetYearMonth_(ss);
  syncSummarySheetFromOpData_(ss, ym.year, ym.month, logSheet);
  ui.alert(
    "集計同期",
    "集計シート（B1=" + ym.label + "）を更新しました。\nD=利用開始合計, E=停止, F=翌月±",
    ui.ButtonSet.OK
  );
}

/**
 * 日報 C21:C36 のみ更新（新規入会＋OP追加の合計）
 * ※日報の D/E/F（月初など）は触らない。貼り付け用は集計シートの D/E/F を参照
 */
function updateNippoSheetForMonth(ss, targetYear, targetMonth, logSheet) {
  const nippoSheet = ss.getSheetByName("日報");
  if (!nippoSheet) return { total: 0, monthRows: 0 };

  const result = countOpDataForMonth_(logSheet, targetYear, targetMonth);
  const cCol = result.counts.map((c) => [c.newSignup + c.opAdd]);
  nippoSheet.getRange(21, 3, OPTION_LIST.length, 1).setValues(cCol);

  let total = 0;
  for (let i = 0; i < cCol.length; i++) total += cCol[i][0];
  return { total: total, monthRows: result.monthRows };
}

/** @deprecated 互換用 */
function updateNippoSheetIfCurrentMonth(ss, targetYear, targetMonth, logSheet) {
  return updateNippoSheetForMonth(ss, targetYear, targetMonth, logSheet);
}

/**
 * メニュー／図形ボタン「日報作成」用: OPデータ → 日報 C21:C36（Gmail取得なし）
 */
function refreshNippoCurrentMonthFromLog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(SHEET_NAME_LOG);
  if (!logSheet || logSheet.getLastRow() < 2) {
    ui.alert(
      "日報の数値だけ更新",
      "OPデータが空です。\n先に「① 全部更新する」を実行してください。",
      ui.ButtonSet.OK
    );
    return;
  }

  const ym = resolveNippoTargetYearMonth_(ss);
  const summaryYm = resolveSummaryTargetYearMonth_(ss);
  if (summaryYm.year === ym.year && summaryYm.month === ym.month) {
    syncSummarySheetFromOpData_(ss, summaryYm.year, summaryYm.month, logSheet);
  }
  const result = updateNippoSheetForMonth(ss, ym.year, ym.month, logSheet);

  let msg =
    ym.label +
    " の日報 C21:C36（当月の利用開始合計）を更新しました。\n停止・翌月±は「集計」シートの D/E/F を見てください。";
  if (result.total === 0) {
    msg +=
      "\n\n0件です。①で集計B1の月を取り込んだか、日報B1（例:2606）が見たい月か確認してください。";
  }

  ui.alert("日報の数値だけ更新", msg, ui.ButtonSet.OK);
}

/**
 * 7. オプション名の表記揺れを吸収
 */
function normalizeOptionName(rawName) {
  const s = String(rawName || "");
  if (s.includes("水素水") && s.includes("プロテイン")) return "プロテイン＋水素水";
  if (s.includes("VIP") && (s.includes("あんしん") || s.includes("安心"))) return "安心サポートVIP";
  if (s.includes("安心サポート") || s.includes("あんしんサポート")) return "安心サポート";
  if (s.includes("ボディプランナー") || s.includes("ボディープランナー")) return "体組成計";
  if (s.includes("マットレンタル") || s.includes("レンタルマット")) return "レンタルマット";
  if (s.includes("ピラティス")) return "ピラティスリフォーマー";

  for (let i = 0; i < OPTION_LIST.length; i++) {
    if (s.includes(OPTION_LIST[i])) return OPTION_LIST[i];
  }
  return s;
}