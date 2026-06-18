/**
 * JOYFIT24経堂 — ご入会確認メール → スプレッドシート取込
 *
 * 【運用スプレッドシート】
 * https://docs.google.com/spreadsheets/d/1S1Noom6Y2LYovkBwUO_SuzSFvC4W0VB3zM0Hjd5kkK0/edit?gid=1946882884#gid=1946882884
 * （リンク集: リポジトリ マスタ/各種リンク.md）
 *
 * 【貼り付け】上記ブック → 拡張機能 → Apps Script → 本ファイル全文
 *
 * 【期間】2025年12月1日 ～ 実行日（今日）まで
 * 【件数】Gmailを月別＋ページ送りで全件探索（500件上限で切らない）
 * 【1行】入会メール1通＝1行（メールIDで重複除外。スレッド返信の二重は除外）
 */

var KYODO_ENROLL_SUBJECT = "【JOYFIT24経堂】ご入会ありがとうございます。※必ず一読ください";
var KYODO_ENROLL_FROM = "info@joyfit-service.jp";
var KYODO_ENROLL_SHEET_NAME = "シート1";
var KYODO_MATCH_SHEET_NAME = "照合";
var KYODO_RED_SHEET_NAME = "赤";
var KYODO_TEST_SHEET_NAME = "テスト";
var KYODO_ENROLL_LOG_SHEET_NAME = "_取込ログ";
var KYODO_ENROLL_PAGE_SIZE = 100;
var KYODO_ENROLL_RANGE_START = new Date(2025, 11, 1); // 2025/12/1
/** 運用ブック ID（マスタ/各種リンク.md と同期） */
var KYODO_MASTER_SPREADSHEET_ID = "1S1Noom6Y2LYovkBwUO_SuzSFvC4W0VB3zM0Hjd5kkK0";

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("入会メール取込")
    .addItem("Gmailから取込（全件）", "kyodoImportEnrollmentEmails")
    .addSeparator()
    .addItem("照合リストにメール反映", "kyodoHookupEmailsByNameAndDatetime")
    .addSeparator()
    .addItem("「赤」シートへ抽出（赤い行）", "kyodoExportRedMarkedRows")
    .addSeparator()
    .addItem("キャンペーン確認メール（テスト1件・手入力）", "kyodoSendCampaignMailTest")
    .addItem("キャンペーン確認メール（テストシート送信）", "kyodoSendCampaignMailToTestSheet")
    .addItem("キャンペーン確認メール（赤・一斉送信）", "kyodoSendCampaignMailToRed")
    .addToUi();
}

/** メニュー／スプレッドシート上のボタンに割り当て可能 */
function kyodoExportRedMarkedRows() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = kyodoExportRedMarkedRowsCore_();
    ui.alert(
      "赤",
      "抽出: " + result.count + " 件\nシート「" + KYODO_RED_SHEET_NAME + "」に出力しました。",
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log(e);
    ui.alert("赤", "エラー: " + e.message, ui.ButtonSet.OK);
  }
}

/** シート1で背景が赤の行だけ「赤」シートにコピー */
function kyodoExportRedMarkedRowsCore_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var src = kyodoGetOutputSheet_(ss);
  var lastRow = src.getLastRow();
  var lastCol = src.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    throw new Error("シート1にデータがありません。");
  }

  var header = src.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerBg = src.getRange(1, 1, 1, lastCol).getBackgrounds()[0];
  var dataHeight = lastRow - 1;
  var values = src.getRange(2, 1, dataHeight, lastCol).getValues();
  var backgrounds = src.getRange(2, 1, dataHeight, lastCol).getBackgrounds();

  var pickedValues = [];
  var pickedBgs = [];
  for (var r = 0; r < dataHeight; r++) {
    if (!kyodoRowBackgroundsContainRed_(backgrounds[r])) continue;
    pickedValues.push(values[r]);
    pickedBgs.push(backgrounds[r]);
  }

  var dest = ss.getSheetByName(KYODO_RED_SHEET_NAME);
  if (!dest) {
    dest = ss.insertSheet(KYODO_RED_SHEET_NAME);
  } else {
    dest.clear();
  }

  dest.getRange(1, 1, 1, lastCol).setValues([header]);
  dest.getRange(1, 1, 1, lastCol).setBackgrounds([headerBg]);
  dest.getRange(1, 1, 1, lastCol).setFontWeight("bold");
  dest.setFrozenRows(1);

  if (pickedValues.length) {
    dest.getRange(2, 1, pickedValues.length, lastCol).setValues(pickedValues);
    dest.getRange(2, 1, pickedValues.length, lastCol).setBackgrounds(pickedBgs);
    dest.autoResizeColumns(1, lastCol);
  }

  return { count: pickedValues.length };
}

/** 行のいずれかのセルが赤系背景か */
function kyodoRowBackgroundsContainRed_(rowBgs) {
  for (var i = 0; i < rowBgs.length; i++) {
    if (kyodoIsRedBackground_(rowBgs[i])) return true;
  }
  return false;
}

/** 手動で塗った赤・えんじ・ピンク寄り赤を判定 */
function kyodoIsRedBackground_(color) {
  var hex = kyodoNormalizeBgHex_(color);
  if (!hex) return false;

  var known = {
    "#ff0000": 1,
    "#ff5252": 1,
    "#ff1744": 1,
    "#f44336": 1,
    "#e53935": 1,
    "#d32f2f": 1,
    "#c62828": 1,
    "#b71c1c": 1,
    "#ef5350": 1,
    "#e57373": 1,
    "#ff8a80": 1,
    "#990000": 1,
    "#cc0000": 1,
    "#800000": 1,
    "#a61e1e": 1,
    "#8b0000": 1
  };
  if (known[hex]) return true;

  var r = parseInt(hex.substring(1, 3), 16);
  var g = parseInt(hex.substring(3, 5), 16);
  var b = parseInt(hex.substring(5, 7), 16);
  if (isNaN(r) || isNaN(g) || isNaN(b)) return false;
  return r >= 150 && r > g + 40 && r > b + 40;
}

function kyodoNormalizeBgHex_(color) {
  var s = String(color || "").trim().toLowerCase();
  if (!s || s === "#ffffff" || s === "#fff" || s === "white") return "";
  if (s.charAt(0) !== "#") s = "#" + s;
  if (s.length === 4) {
    s = "#" + s.charAt(1) + s.charAt(1) + s.charAt(2) + s.charAt(2) + s.charAt(3) + s.charAt(3);
  }
  return s.length === 7 ? s : "";
}

function kyodoImportEnrollmentEmails() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = kyodoImportEnrollmentEmailsCore_();
    var lines = [
      "取込完了",
      "期間: " + result.rangeLabel,
      "総数: " + result.count + " 件（入会メール1通＝1行）",
      "Gmailメッセージ: " + result.messageCount + " 通",
      "名前未取得: " + result.nameMissing + " 件",
      "スレッド返信など除外: " + result.skippedDuplicate + " 件"
    ];
    if (result.hitPageLimit) {
      lines.push("※ 一部チャンクで件数上限。再実行するか _取込ログ を確認");
    }
    ui.alert("入会メール取込", lines.join("\n"), ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(e);
    ui.alert("入会メール取込", "エラー: " + e.message, ui.ButtonSet.OK);
  }
}

function kyodoHookupEmailsByNameAndDatetime() {
  var ui = SpreadsheetApp.getUi();
  try {
    var result = kyodoHookupEmailsCore_();
    ui.alert(
      "照合",
      "反映完了\n一致: " + result.matched + " 件\n不一致: " + result.unmatched + " 件\n要確認: " + result.ambiguous + " 件",
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log(e);
    ui.alert("照合", "エラー: " + e.message, ui.ButtonSet.OK);
  }
}

function kyodoImportEnrollmentEmailsCore_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = kyodoGetOutputSheet_(ss);
  if (!sh) throw new Error("シートが見つかりません。");

  var rangeEnd = kyodoEndOfToday_();
  var fetchResult = kyodoFetchAllEnrollmentRows_(KYODO_ENROLL_RANGE_START, rangeEnd);
  var rows = fetchResult.rows;

  rows.sort(function (a, b) {
    return b.date.getTime() - a.date.getTime();
  });

  var tz = Session.getScriptTimeZone();
  sh.clear();
  sh.getRange(1, 1, 1, 5).setValues([
    ["タイムスタンプ", "名前", "メールアドレス", "受付番号", "備考"]
  ]);
  sh.getRange(1, 1, 1, 5).setFontWeight("bold");

  var nameMissing = 0;
  if (rows.length) {
    var out = [];
    for (var i = 0; i < rows.length; i++) {
      if (!rows[i].name) nameMissing++;
      out.push([
        Utilities.formatDate(rows[i].date, tz, "yyyy/MM/dd HH:mm"),
        rows[i].name,
        rows[i].email,
        rows[i].receptionNo,
        rows[i].note
      ]);
    }
    sh.getRange(2, 1, out.length, 5).setValues(out);
    sh.autoResizeColumns(1, 5);
  }

  kyodoWriteImportLog_(ss, fetchResult, rows.length, nameMissing, tz);

  return {
    count: rows.length,
    sheetName: sh.getName(),
    rangeLabel: kyodoFormatRangeLabel_(KYODO_ENROLL_RANGE_START, rangeEnd, tz),
    messageCount: fetchResult.messageCount,
    nameMissing: nameMissing,
    skippedDuplicate: fetchResult.skippedDuplicate,
    hitPageLimit: fetchResult.hitPageLimit
  };
}

/** 月別チャンク × ページ送りで Gmail 全探索 */
function kyodoFetchAllEnrollmentRows_(rangeStart, rangeEnd) {
  var chunks = kyodoBuildMonthChunks_(rangeStart, rangeEnd);
  var seenMessageId = {};
  var seenEnrollmentKey = {};
  var rows = [];
  var messageCount = 0;
  var skippedDuplicate = 0;
  var hitPageLimit = false;
  var threadsScanned = 0;

  for (var c = 0; c < chunks.length; c++) {
    var chunk = chunks[c];
    var start = 0;
    while (true) {
      var query = kyodoBuildGmailSearchQuery_(chunk.after, chunk.before);
      var threads = GmailApp.search(query, start, KYODO_ENROLL_PAGE_SIZE);
      if (!threads.length) break;

      threadsScanned += threads.length;

      for (var t = 0; t < threads.length; t++) {
        var messages = threads[t].getMessages();
        for (var m = 0; m < messages.length; m++) {
          var msg = messages[m];
          if (!kyodoIsEnrollmentMessage_(msg)) continue;

          var date = msg.getDate();
          if (!kyodoIsInEnrollRange_(date, rangeStart, rangeEnd)) continue;

          var msgId = msg.getId();
          if (seenMessageId[msgId]) continue;
          seenMessageId[msgId] = true;
          messageCount++;

          var email = kyodoGetRecipientEmail_(msg);
          if (!email) continue;

          var minuteKey =
            email +
            "\x1f" +
            Utilities.formatDate(
              kyodoTruncateToMinute_(date),
              Session.getScriptTimeZone(),
              "yyyy/MM/dd HH:mm"
            );
          if (seenEnrollmentKey[minuteKey]) {
            skippedDuplicate++;
            continue;
          }
          seenEnrollmentKey[minuteKey] = true;

          var name = kyodoExtractMemberName_(msg);
          var receptionNo = kyodoExtractReceptionNo_(msg);
          var note = "";
          if (!name) note = "名前未取得（要確認）";

          rows.push({
            date: date,
            name: name,
            email: email,
            receptionNo: receptionNo,
            note: note
          });
        }
      }

      start += threads.length;
      if (threads.length < KYODO_ENROLL_PAGE_SIZE) break;
      if (start >= 500) {
        hitPageLimit = true;
        Logger.log("Gmail search page limit at chunk " + chunk.label + " start=" + start);
        break;
      }
      Utilities.sleep(150);
    }
    Utilities.sleep(100);
  }

  return {
    rows: rows,
    messageCount: messageCount,
    skippedDuplicate: skippedDuplicate,
    hitPageLimit: hitPageLimit,
    threadsScanned: threadsScanned,
    chunks: chunks.length
  };
}

function kyodoBuildGmailSearchQuery_(afterDate, beforeDate) {
  var afterStr = Utilities.formatDate(afterDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
  var beforeStr = Utilities.formatDate(beforeDate, Session.getScriptTimeZone(), "yyyy/MM/dd");
  return (
    "from:" +
    KYODO_ENROLL_FROM +
    ' subject:"' +
    KYODO_ENROLL_SUBJECT +
    '" after:' +
    afterStr +
    " before:" +
    beforeStr
  );
}

function kyodoBuildMonthChunks_(rangeStart, rangeEnd) {
  var chunks = [];
  var cur = new Date(rangeStart.getFullYear(), rangeStart.getMonth(), 1);
  var endLimit = new Date(rangeEnd.getFullYear(), rangeEnd.getMonth() + 1, 1);

  while (cur.getTime() < endLimit.getTime()) {
    var next = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
    var chunkStart = cur.getTime() < rangeStart.getTime() ? rangeStart : cur;
    var chunkEnd = next.getTime() > endLimit.getTime() ? endLimit : next;
    chunks.push({
      after: chunkStart,
      before: chunkEnd,
      label: Utilities.formatDate(cur, Session.getScriptTimeZone(), "yyyy/MM")
    });
    cur = next;
  }
  return chunks;
}

function kyodoIsInEnrollRange_(date, rangeStart, rangeEnd) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return false;
  var t = date.getTime();
  return t >= rangeStart.getTime() && t <= rangeEnd.getTime();
}

function kyodoEndOfToday_() {
  var now = new Date();
  return new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59, 999);
}

function kyodoFormatRangeLabel_(start, end, tz) {
  return (
    Utilities.formatDate(start, tz, "yyyy/MM/dd") +
    " ～ " +
    Utilities.formatDate(end, tz, "yyyy/MM/dd")
  );
}

function kyodoWriteImportLog_(ss, fetchResult, rowCount, nameMissing, tz) {
  var sh = ss.getSheetByName(KYODO_ENROLL_LOG_SHEET_NAME);
  if (!sh) sh = ss.insertSheet(KYODO_ENROLL_LOG_SHEET_NAME);
  sh.clear();
  var now = new Date();
  sh.getRange(1, 1, 8, 2).setValues([
    ["最終取込", Utilities.formatDate(now, tz, "yyyy/MM/dd HH:mm:ss")],
    ["対象期間", kyodoFormatRangeLabel_(KYODO_ENROLL_RANGE_START, kyodoEndOfToday_(), tz)],
    ["一覧行数（総数）", rowCount],
    ["Gmailメッセージ数", fetchResult.messageCount],
    ["名前未取得", nameMissing],
    ["重複除外", fetchResult.skippedDuplicate],
    ["探索スレッド数", fetchResult.threadsScanned],
    ["月別チャンク数", fetchResult.chunks]
  ]);
  try {
    sh.hideSheet();
  } catch (e) {}
}

function kyodoHookupEmailsCore_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var enrollSh = kyodoGetOutputSheet_(ss);
  if (!enrollSh || enrollSh.getLastRow() < 2) {
    throw new Error("先に「Gmailから取込（全件）」でシート1にデータを入れてください。");
  }

  var matchSh = kyodoEnsureMatchSheet_(ss);
  var lastRow = matchSh.getLastRow();
  if (lastRow < 2) {
    throw new Error("「" + KYODO_MATCH_SHEET_NAME + "」に名前と入会日時を入力してください。");
  }

  var index = kyodoBuildEnrollmentIndex_(enrollSh);
  var tz = Session.getScriptTimeZone();
  var height = lastRow - 1;
  var vals = matchSh.getRange(2, 1, height, 2).getValues();
  var emails = [];
  var statuses = [];
  var matched = 0;
  var unmatched = 0;
  var ambiguous = 0;

  for (var r = 0; r < vals.length; r++) {
    var name = kyodoNormalizePersonName_(vals[r][0]);
    var dt = kyodoParseSheetDateTime_(vals[r][1]);
    var emailOut = "";
    var status = "";

    if (!name) {
      status = "名前が空";
      unmatched++;
    } else if (!dt) {
      status = "入会日時が読めない";
      unmatched++;
    } else {
      var key = kyodoMakeLookupKey_(name, dt, tz);
      var hit = index[key];
      if (!hit) {
        var nameOnly = index["name:" + name];
        if (nameOnly && nameOnly.count > 0) {
          status = "日時不一致（同名は別日時に存在）";
          ambiguous++;
        } else {
          status = "該当なし";
          unmatched++;
        }
      } else if (hit.ambiguous) {
        status = "複数候補（要手動確認）";
        emailOut = hit.email || "";
        ambiguous++;
      } else {
        emailOut = hit.email;
        status = "一致";
        matched++;
      }
    }

    emails.push([emailOut]);
    statuses.push([status]);
  }

  matchSh.getRange(2, 3, height, 1).setValues(emails);
  matchSh.getRange(2, 4, height, 1).setValues(statuses);
  matchSh.autoResizeColumns(1, 4);

  return { matched: matched, unmatched: unmatched, ambiguous: ambiguous };
}

function kyodoBuildEnrollmentIndex_(enrollSh) {
  var tz = Session.getScriptTimeZone();
  var data = enrollSh.getDataRange().getValues();
  var index = {};

  for (var i = 1; i < data.length; i++) {
    var dt = kyodoParseSheetDateTime_(data[i][0]);
    var name = kyodoNormalizePersonName_(data[i][1]);
    var email = String(data[i][2] || "").trim().toLowerCase();
    if (!name || !dt || !email) continue;

    var key = kyodoMakeLookupKey_(name, dt, tz);
    if (!index[key]) {
      index[key] = { email: email, count: 0, ambiguous: false };
    }
    index[key].count++;
    if (index[key].count === 1) {
      index[key].email = email;
    } else if (index[key].email !== email) {
      index[key].ambiguous = true;
    }

    if (!index["name:" + name]) {
      index["name:" + name] = { count: 0 };
    }
    index["name:" + name].count++;
  }

  return index;
}

function kyodoMakeLookupKey_(name, date, tz) {
  return name + "\x1f" + Utilities.formatDate(kyodoTruncateToMinute_(date), tz, "yyyy/MM/dd HH:mm");
}

function kyodoEnsureMatchSheet_(ss) {
  var sh = ss.getSheetByName(KYODO_MATCH_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(KYODO_MATCH_SHEET_NAME);
  }
  var header = sh.getRange(1, 1, 1, 4).getValues()[0];
  if (!header[0]) {
    sh.getRange(1, 1, 1, 4).setValues([["名前", "入会日時", "メールアドレス", "照合結果"]]);
    sh.getRange(1, 1, 1, 4).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function kyodoParseSheetDateTime_(v) {
  if (v === "" || v === null || v === undefined) return null;
  if (v instanceof Date && !isNaN(v.getTime())) {
    return kyodoTruncateToMinute_(v);
  }
  var s = String(v).trim();
  if (!s) return null;
  var m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})[ T]?(\d{1,2})?:?(\d{1,2})?/);
  if (m) {
    var hh = m[4] !== undefined && m[4] !== "" ? parseInt(m[4], 10) : 0;
    var mm = m[5] !== undefined && m[5] !== "" ? parseInt(m[5], 10) : 0;
    return kyodoTruncateToMinute_(
      new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10), hh, mm)
    );
  }
  var d = new Date(s);
  if (!isNaN(d.getTime())) return kyodoTruncateToMinute_(d);
  return null;
}

function kyodoTruncateToMinute_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes());
}

function kyodoGetOutputSheet_(ss) {
  return (
    ss.getSheetByName(KYODO_ENROLL_SHEET_NAME) ||
    ss.getSheetByName("Sheet1") ||
    ss.getSheets()[0]
  );
}

function kyodoIsEnrollmentMessage_(msg) {
  var subj = String(msg.getSubject() || "").trim();
  if (subj !== KYODO_ENROLL_SUBJECT) return false;
  var from = kyodoParseEmailAddress_(msg.getFrom());
  if (from === KYODO_ENROLL_FROM.toLowerCase()) return true;
  return String(msg.getFrom() || "").indexOf("joyfit-service") !== -1;
}

function kyodoGetRecipientEmail_(msg) {
  var fields = [msg.getTo(), msg.getCc(), msg.getBcc()];
  for (var i = 0; i < fields.length; i++) {
    var emails = kyodoParseAllEmails_(fields[i]);
    for (var j = 0; j < emails.length; j++) {
      if (emails[j].indexOf("joyfit") === -1 && emails[j].indexOf("okamoto-group") === -1) {
        return emails[j];
      }
    }
  }
  var toOnly = kyodoParseEmailAddress_(msg.getTo());
  return toOnly || "";
}

function kyodoParseEmailAddress_(raw) {
  var list = kyodoParseAllEmails_(raw);
  return list.length ? list[0] : "";
}

function kyodoParseAllEmails_(raw) {
  var s = String(raw || "").trim();
  if (!s) return [];
  var out = [];
  var re = /[\w.+-]+@[\w.-]+\.\w+/g;
  var m;
  while ((m = re.exec(s)) !== null) {
    out.push(m[0].toLowerCase());
  }
  return out;
}

function kyodoGetMessageBodyText_(msg) {
  var plain = msg.getPlainBody();
  if (plain && String(plain).trim()) return String(plain);
  var html = msg.getBody();
  if (!html) return "";
  return String(html)
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .trim();
}

function kyodoExtractReceptionNo_(msg) {
  var body = kyodoGetMessageBodyText_(msg);
  var m = body.match(/受付番号\s*[:：]\s*(\d+)/);
  return m ? m[1] : "";
}

function kyodoExtractMemberName_(msg) {
  var body = kyodoGetMessageBodyText_(msg);

  var patterns = [
    /お名前\s*[:：]\s*([^\n\r]+)/,
    /氏名\s*[:：]\s*([^\n\r]+)/,
    /会員名\s*[:：]\s*([^\n\r]+)/
  ];
  for (var p = 0; p < patterns.length; p++) {
    var m = body.match(patterns[p]);
    if (m) {
      var n = kyodoNormalizePersonName_(m[1]);
      if (n) return n;
    }
  }

  var lines = body.split(/\r?\n/);
  for (var i = 0; i < Math.min(lines.length, 15); i++) {
    var line = String(lines[i] || "").trim();
    if (!line) continue;
    if (/^(https?|ー|－|－－|【|■|●|この度|ご利用|JOYFIT)/.test(line)) continue;
    var mTop = line.match(/^(.{1,40}?)\s*様\s*$/);
    if (mTop) {
      var n2 = kyodoNormalizePersonName_(mTop[1]);
      if (n2 && !/[@＠]/.test(n2)) return n2;
    }
  }

  return "";
}

function kyodoNormalizePersonName_(raw) {
  var s = String(raw || "")
    .replace(/\s*様\s*$/g, "")
    .replace(/[ \u3000]+/g, " ")
    .replace(/^[0-9０-９\s]+/, "")
    .trim();
  if (!s) return "";
  if (/[@＠]/.test(s) || /^https?:/i.test(s)) return "";

  var parts = s.split(" ");
  if (parts.length >= 2 && parts.length % 2 === 0) {
    var half = parts.length / 2;
    var a = parts.slice(0, half).join(" ");
    var b = parts.slice(half).join(" ");
    if (a === b) return a;
  }
  return s;
}

// --- キャンペーン在籍確認メール（「赤」シート宛） ---

var KYODO_CAMPAIGN_MAIL_SUBJECT =
  "【重要】JOYFIT24経堂より：キャンペーン在籍期間に関するご確認";
var KYODO_CAMPAIGN_REPLY_DEADLINE = "6月5日（金）";
var KYODO_CAMPAIGN_SEND_LOG_HEADER = "送信日時";
var KYODO_CAMPAIGN_SEND_DELAY_MS = 1200;

/** テスト送信（自分宛など1件） */
function kyodoSendCampaignMailTest() {
  var ui = SpreadsheetApp.getUi();
  try {
    var me = "";
    try {
      me = Session.getActiveUser().getEmail();
    } catch (e) {}
    var resp = ui.prompt(
      "テスト送信",
      "送信先メールアドレス",
      ui.ButtonSet.OK_CANCEL,
      me ? me : ""
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    var to = kyodoParseEmailAddress_(resp.getResponseText());
    if (!to) throw new Error("メールアドレスが空です。");
    var name = "テスト 太郎";
    try {
      var list = kyodoLoadTestSheetRecipients_();
      if (list.length) name = "テスト";
    } catch (ignore) {}
    kyodoSendOneCampaignMail_(to, name, true);
    ui.alert("テスト送信", "【テスト】件名で送信しました:\n" + to, ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(e);
    ui.alert("テスト送信", "エラー: " + e.message, ui.ButtonSet.OK);
  }
}

/** 「テスト」シートのメールアドレスだけに送信（件名に【テスト】付与） */
function kyodoSendCampaignMailToTestSheet() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var list = kyodoLoadTestSheetRecipients_();
    if (!list.length) {
      throw new Error(
        "「" +
          KYODO_TEST_SHEET_NAME +
          "」シートのA列にメールアドレスがありません。\nA列だけに1行1アドレスで入力してください。"
      );
    }

    var addrs = [];
    for (var i = 0; i < list.length; i++) {
      if (list[i].email) addrs.push(list[i].email);
    }

    var preview =
      "【テスト送信】\n\n件名:\n【テスト】" +
      KYODO_CAMPAIGN_MAIL_SUBJECT +
      "\n\n送信先（「" +
      KYODO_TEST_SHEET_NAME +
      "」シート・A列のみ）: " +
      addrs.length +
      " 件\n" +
      addrs.join("\n") +
      "\n\n送信しますか？";
    if (ui.alert("テストシート送信の確認", preview, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
      return;
    }

    var result = kyodoSendCampaignMailBatch_(ss, list, KYODO_TEST_SHEET_NAME, {
      isTest: true,
      skipAlreadySent: false
    });
    ui.alert(
      "テストシート送信",
      "送信: " + result.sent + " 件\n失敗: " + result.failed + " 件",
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log(e);
    ui.alert("テストシート送信", "エラー: " + e.message, ui.ButtonSet.OK);
  }
}

/** 「赤」シートの全員にHTMLメール一斉送信 */
function kyodoSendCampaignMailToRed() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var list = kyodoLoadRedSheetRecipients_();
    if (!list.length) throw new Error("「赤」シートに送信先がありません。");

    var preview =
      "件名:\n" +
      KYODO_CAMPAIGN_MAIL_SUBJECT +
      "\n\n送信先: " +
      list.length +
      " 件（赤シート）\n返信期限: 明日【" +
      KYODO_CAMPAIGN_REPLY_DEADLINE +
      "】中\n\nこの内容で送信しますか？";
    if (ui.alert("一斉送信の確認", preview, ui.ButtonSet.YES_NO) !== ui.Button.YES) {
      return;
    }

    var result = kyodoSendCampaignMailBatch_(ss, list, KYODO_RED_SHEET_NAME, {
      isTest: false,
      skipAlreadySent: true
    });
    ui.alert(
      "一斉送信",
      "送信: " + result.sent + " 件\n失敗: " + result.failed + " 件\nスキップ: " + result.skipped + " 件",
      ui.ButtonSet.OK
    );
  } catch (e) {
    Logger.log(e);
    ui.alert("一斉送信", "エラー: " + e.message, ui.ButtonSet.OK);
  }
}

function kyodoSendCampaignMailBatch_(ss, list, sheetName, options) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error("シート「" + sheetName + "」が見つかりません。");

  var opts = options || {};
  var isTest = !!opts.isTest;
  var skipAlreadySent = opts.skipAlreadySent !== false;

  var sent = 0;
  var failed = 0;
  var skipped = 0;
  var logCol = kyodoEnsureCampaignSendLogColumn_(sh);
  var tz = Session.getScriptTimeZone();

  for (var i = 0; i < list.length; i++) {
    var row = list[i];
    if (skipAlreadySent && row.alreadySent) {
      skipped++;
      continue;
    }
    if (!row.email) {
      failed++;
      if (logCol > 0) {
        sh.getRange(row.sheetRow, logCol).setValue("失敗: メールなし");
      }
      continue;
    }
    try {
      kyodoSendOneCampaignMail_(row.email, row.name, isTest);
      sent++;
      if (logCol > 0) {
        sh.getRange(row.sheetRow, logCol).setValue(
          Utilities.formatDate(new Date(), tz, "yyyy/MM/dd HH:mm") + (isTest ? "（テスト）" : "")
        );
      }
      Utilities.sleep(KYODO_CAMPAIGN_SEND_DELAY_MS);
    } catch (err) {
      failed++;
      Logger.log(err);
      if (logCol > 0) {
        sh.getRange(row.sheetRow, logCol).setValue("失敗: " + String(err.message || err));
      }
    }
  }
  return { sent: sent, failed: failed, skipped: skipped };
}

function kyodoSendOneCampaignMail_(to, name, isTest) {
  var displayName = String(name || "").trim() || "お客";
  var plain = kyodoBuildCampaignMailPlain_(displayName);
  var html = kyodoBuildCampaignMailHtml_(displayName);
  var subject = KYODO_CAMPAIGN_MAIL_SUBJECT;
  if (isTest) subject = "【テスト】" + subject;
  MailApp.sendEmail({
    to: to,
    subject: subject,
    body: plain,
    htmlBody: html
  });
}

function kyodoBuildCampaignMailPlain_(name) {
  var lines = [];
  lines.push(name + " 様");
  lines.push("");
  lines.push("いつもJOYFIT24経堂をご利用いただき、誠にありがとうございます。");
  lines.push("");
  lines.push(
    "本日は、ご入会時に適用されました「満6ヶ月キャンペーン」の在籍条件につきまして、大切なお願いがありご連絡いたしました。"
  );
  lines.push("");
  lines.push(
    "適用させていただいたキャンペーンは「満6ヶ月間のご在籍」が必須条件となっております。"
  );
  lines.push(
    "今回は期間途中でのご解約となりますため、恐れ入りますが下記【A】または【B】のいずれかをご選択いただきたく存じます。"
  );
  lines.push("お手数ですが、どちらがご希望かお教えください。");
  lines.push("");
  lines.push("【A】退会を確定し、解約金を支払う");
  lines.push("キャンペーン時の値引き相当額（途中解約金）をお支払いいただきます。");
  lines.push("");
  lines.push("【B】退会をキャンセルし、必須期間まで会員を継続する");
  lines.push("一度退会をキャンセルし、必須在籍期間まで会員としてご継続いただきます。");
  lines.push("");
  lines.push("各種手続きの都合上、大変急なお願いとなり誠に恐縮ではございますが、");
  lines.push("明日【" + KYODO_CAMPAIGN_REPLY_DEADLINE + "】中までにご返信いただけますと幸いです。");
  lines.push("");
  lines.push("■ ご返信のお願い");
  lines.push("・【A】または【B】のどちらがご希望か、本メールへご返信ください。");
  lines.push("・ご不明点がございましたら、同じく本メールよりお気軽にご連絡ください。");
  lines.push("　（本メールへの返信にて承ります）");
  lines.push("");
  lines.push("お客様にはお手数をおかけいたしますが、何卒ご理解とご協力のほどよろしくお願い申し上げます。");
  lines.push("");
  lines.push("JOYFIT24経堂");
  return lines.join("\n");
}

function kyodoBuildCampaignMailHtml_(name) {
  var n = kyodoEscapeHtml_(name);
  var deadline = kyodoEscapeHtml_(KYODO_CAMPAIGN_REPLY_DEADLINE);
  return (
    "<div style=\"background:#f5f5f5;padding:24px 0;font-family:'Meiryo','Hiragino Kaku Gothic ProN',Arial,sans-serif;color:#333;line-height:1.75;\">" +
    "<div style=\"max-width:640px;margin:0 auto;background:#fff;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;\">" +
    "<div style=\"background:linear-gradient(90deg,#c62828 0%,#e91e63 100%);padding:16px 20px;color:#fff;font-size:17px;font-weight:bold;\">" +
    "JOYFIT24経堂" +
    "</div>" +
    "<div style=\"padding:24px 20px;font-size:14px;\">" +
    "<p style=\"margin:0 0 16px 0;\">" + n + " 様</p>" +
    "<p style=\"margin:0 0 16px 0;\">いつもJOYFIT24経堂をご利用いただき、誠にありがとうございます。</p>" +
    "<p style=\"margin:0 0 16px 0;\">本日は、ご入会時に適用されました「満6ヶ月キャンペーン」の在籍条件につきまして、大切なお願いがありご連絡いたしました。</p>" +
    "<p style=\"margin:0 0 16px 0;\">適用させていただいたキャンペーンは「満6ヶ月間のご在籍」が必須条件となっております。<br>" +
    "今回は期間途中でのご解約となりますため、恐れ入りますが下記<strong>【A】または【B】</strong>のいずれかをご選択いただきたく存じます。<br>" +
    "お手数ですが、<strong>どちらがご希望か</strong>お教えください。</p>" +
    "<div style=\"margin:20px 0;padding:16px;background:#fff8f8;border-left:4px solid #c62828;border-radius:4px;\">" +
    "<p style=\"margin:0 0 12px 0;font-weight:bold;color:#b71c1c;\">【A】退会を確定し、解約金を支払う</p>" +
    "<p style=\"margin:0 0 16px 0;\">キャンペーン時の値引き相当額（途中解約金）をお支払いいただきます。</p>" +
    "<p style=\"margin:0 0 12px 0;font-weight:bold;color:#b71c1c;\">【B】退会をキャンセルし、必須期間まで会員を継続する</p>" +
    "<p style=\"margin:0;\">一度退会をキャンセルし、必須在籍期間まで会員としてご継続いただきます。</p>" +
    "</div>" +
    "<p style=\"margin:0 0 16px 0;\">各種手続きの都合上、大変急なお願いとなり誠に恐縮ではございますが、<br>" +
    "<strong>明日【" + deadline + "】中まで</strong>にご返信いただけますと幸いです。</p>" +
    "<div style=\"margin:0 0 16px 0;padding:14px 16px;background:#f3f6fb;border-radius:6px;\">" +
    "<p style=\"margin:0 0 8px 0;font-weight:bold;\">■ ご返信のお願い</p>" +
    "<p style=\"margin:0 0 6px 0;\">・<strong>【A】または【B】のどちらがご希望か</strong>、本メールへご返信ください。</p>" +
    "<p style=\"margin:0 0 6px 0;\">・ご不明点がございましたら、同じく<strong>本メールより</strong>お気軽にご連絡ください。</p>" +
    "<p style=\"margin:0;color:#555;font-size:13px;\">（本メールへの返信にて承ります）</p>" +
    "</div>" +
    "<p style=\"margin:0;\">お客様にはお手数をおかけいたしますが、何卒ご理解とご協力のほどよろしくお願い申し上げます。</p>" +
    "<p style=\"margin:24px 0 0 0;color:#777;font-size:12px;\">JOYFIT24経堂</p>" +
    "</div></div></div>"
  );
}

function kyodoEscapeHtml_(s) {
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function kyodoLoadRedSheetRecipients_() {
  return kyodoLoadSheetRecipients_(KYODO_RED_SHEET_NAME, "「赤」シートにデータがありません。");
}

/**
 * 「テスト」シートの A 列だけ読む（メールアドレスのみ・1行1件）
 * ヘッダー行（@なし）は自動スキップ
 */
function kyodoLoadTestSheetRecipients_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(KYODO_TEST_SHEET_NAME);
  if (!sh) {
    throw new Error("「" + KYODO_TEST_SHEET_NAME + "」シートがありません。タブ名を「テスト」にしてください。");
  }

  var lastRow = sh.getLastRow();
  if (lastRow < 1) {
    throw new Error("「" + KYODO_TEST_SHEET_NAME + "」シートのA列にメールアドレスを入力してください。");
  }

  var vals = sh.getRange(1, 1, lastRow, 1).getValues();
  var list = [];
  var seen = {};

  for (var r = 0; r < vals.length; r++) {
    var email = kyodoParseEmailAddress_(vals[r][0]);
    if (!email) continue;
    if (seen[email]) continue;
    seen[email] = true;
    list.push({
      sheetRow: r + 1,
      name: "テスト",
      email: email,
      alreadySent: false
    });
  }
  return list;
}

/** 指定シートから名前・メールを読み込む（2行目以降・メールがある行のみ） */
function kyodoLoadSheetRecipients_(sheetName, emptyError) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) {
    throw new Error(emptyError || "シートにデータがありません。");
  }

  var data = sh.getDataRange().getValues();
  var header = data[0];
  var nameCol = kyodoFindHeaderColIndex_(header, ["名前"]);
  var emailCol = kyodoFindHeaderColIndex_(header, ["メールアドレス", "メール"]);
  var sentCol = kyodoFindHeaderColIndex_(header, [KYODO_CAMPAIGN_SEND_LOG_HEADER, "送信済"]);
  if (nameCol < 0) nameCol = 1;
  if (emailCol < 0) emailCol = 2;

  var list = [];
  for (var r = 1; r < data.length; r++) {
    var name = kyodoNormalizePersonName_(data[r][nameCol]);
    var email = kyodoParseEmailAddress_(data[r][emailCol]);
    if (!email) continue;

    var alreadySent = false;
    if (sentCol >= 0) {
      var logVal = String(data[r][sentCol] || "").trim();
      if (logVal && logVal.indexOf("失敗") !== 0) {
        alreadySent = true;
      }
    }
    list.push({
      sheetRow: r + 1,
      name: name || "テスト",
      email: email,
      alreadySent: alreadySent
    });
  }
  return list;
}

function kyodoFindHeaderColIndex_(headerRow, labels) {
  for (var c = 0; c < headerRow.length; c++) {
    var h = String(headerRow[c] || "").trim();
    for (var i = 0; i < labels.length; i++) {
      if (h === labels[i]) return c;
    }
  }
  return -1;
}

function kyodoEnsureCampaignSendLogColumn_(sh) {
  var lastCol = sh.getLastColumn();
  var header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = kyodoFindHeaderColIndex_(header, [KYODO_CAMPAIGN_SEND_LOG_HEADER]);
  if (idx >= 0) return idx + 1;
  var newCol = lastCol + 1;
  sh.getRange(1, newCol).setValue(KYODO_CAMPAIGN_SEND_LOG_HEADER);
  sh.getRange(1, newCol).setFontWeight("bold");
  return newCol;
}
