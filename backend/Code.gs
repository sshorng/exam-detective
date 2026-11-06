/**
 * 考場抓鬼大師・錯題熱點後端（Google Apps Script）
 *
 * 一份試算表服務所有考卷：每份考卷用自己的 examId 與題庫版本區分，不必重新部署。
 *
 * 部署四步：
 *   1. 開一份新的 Google 試算表，命名例如「考場抓鬼大師回報」。
 *   2. 在試算表選「擴充功能 → Apps Script」，把本檔全部內容貼進 Code.gs 並存檔。
 *   3. 部署 → 新增部署作業 → 類型選「網頁應用程式」，
 *      執行身分選「我」，具有存取權的使用者選「任何人」，按部署並授權。
 *   4. 複製產生的 /exec 網址，回到網頁的 index.html，填進 Hotspot 的 GAS_URL 常數。
 *
 * 工作表會自動建立：
 *   回報：時間、批次、考卷、版本、班級、座號、題號、客戶端
 *   歷史：清空時間、層級、時間、批次、考卷、版本、班級、座號、題號、客戶端
 *   設定：B1 教師密碼（首次自動產生 8 碼隨機值，可自行修改）
 */

const SHEET_REPORT = "回報";
const SHEET_HISTORY = "歷史";
const SHEET_SETTING = "設定";
const DEDUPE_MS = 10000;
const REPORT_HEADERS = ["時間", "批次", "考卷", "版本", "班級", "座號", "題號", "客戶端"];
const HISTORY_HEADERS = ["清空時間", "層級", "時間", "批次", "考卷", "版本", "班級", "座號", "題號", "客戶端"];

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) || "{}";
    const body = JSON.parse(raw);
    const action = String(body.action || "");
    const book = SpreadsheetApp.getActiveSpreadsheet();
    if (!book) return json_({ ok: false, error: "找不到試算表，請以容器綁定方式建立腳本" });
    ensureSheets_(book);

    if (action === "report") return json_(report_(book, body));
    if (action === "auth") return json_(auth_(book, body.key));
    if (action === "counts") return json_(counts_(book, body));
    if (action === "clear") return json_(clear_(book, body));
    if (action === "info") return json_({ ok: true, ver: "hotspot-1" });
    return json_({ ok: false, error: "unknown-action" });
  } catch (error) {
    return json_({ ok: false, error: String((error && error.message) || error) });
  }
}

function doGet() {
  return json_({ ok: true, ver: "hotspot-1" });
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function ensureSheets_(book) {
  const report = book.getSheetByName(SHEET_REPORT) || book.insertSheet(SHEET_REPORT);
  if (report.getLastRow() === 0) report.appendRow(REPORT_HEADERS);

  const history = book.getSheetByName(SHEET_HISTORY) || book.insertSheet(SHEET_HISTORY);
  if (history.getLastRow() === 0) history.appendRow(HISTORY_HEADERS);

  const setting = book.getSheetByName(SHEET_SETTING) || book.insertSheet(SHEET_SETTING);
  if (setting.getLastRow() === 0) {
    setting.getRange("A1").setValue("教師密碼");
    setting.getRange("B1").setValue(randomPassword_());
    setting.getRange("A4").setValue("說明");
    setting.getRange("A5").setValue("B1 是教師密碼，可自行修改；修改後網頁需重新輸入一次。");
    setting.getRange("A6").setValue("密碼只在此工作表，網頁與前端程式都不保存明文。");
    setting.setColumnWidth(1, 120);
  }
  return report;
}

function randomPassword_() {
  const chars = "abcdefghijkmnpqrstuvwxyz23456789";
  let out = "";
  for (let index = 0; index < 8; index += 1) {
    out += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return out;
}

function password_(book) {
  const setting = book.getSheetByName(SHEET_SETTING);
  if (!setting || setting.getLastRow() === 0) return "";
  return String(setting.getRange("B1").getValue() || "").trim();
}

function auth_(book, key) {
  const expected = password_(book);
  if (!expected) return { ok: false, error: "尚未設定密碼" };
  if (String(key || "").trim() !== expected) return { ok: false, error: "auth" };
  return { ok: true, sheetUrl: book.getUrl() };
}

function report_(book, body) {
  const examId = String(body.examId || "").trim();
  const ver = String(body.ver || "").trim();
  const klass = String(body.klass || "").trim();
  const seat = String(body.seat || "").trim();
  const client = String(body.client || "").trim();
  const qids = Array.isArray(body.qids) ? body.qids.map(Number).filter(function (id) { return id > 0; }) : [];
  if (!examId || !qids.length) return { ok: false, error: "缺少資料" };
  if (!klass) return { ok: false, error: "缺少班級" };

  const report = book.getSheetByName(SHEET_REPORT);
  const lastRow = report.getLastRow();
  const recent = lastRow > 1 ? report.getRange(Math.max(2, lastRow - 400), 1, Math.min(400, lastRow - 1), REPORT_HEADERS.length).getValues() : [];
  const now = new Date();
  const seen = {};
  recent.forEach(function (row) {
    const stamp = row[0] instanceof Date ? row[0].getTime() : new Date(row[0]).getTime();
    if (!stamp || now.getTime() - stamp > DEDUPE_MS) return;
    seen[String(row[2]) + "|" + String(row[7]) + "|" + String(row[6])] = true;
  });

  const batch = client + "-" + now.getTime();
  const stampText = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const rows = [];
  qids.forEach(function (qid) {
    if (seen[examId + "|" + client + "|" + qid]) return;
    rows.push([stampText, batch, examId, ver, klass, seat, qid, client]);
  });
  if (!rows.length) return { ok: true, added: 0 };
  report.getRange(report.getLastRow() + 1, 1, rows.length, REPORT_HEADERS.length).setValues(rows);
  return { ok: true, added: rows.length };
}

function counts_(book, body) {
  const check = auth_(book, body.key);
  if (!check.ok) return { ok: false, error: "auth" };

  const examId = String(body.examId || "").trim();
  const ver = String(body.ver || "").trim();
  const klassFilter = String(body.klass || "").trim();
  const report = book.getSheetByName(SHEET_REPORT);
  const lastRow = report.getLastRow();
  if (lastRow < 2) return { ok: true, counts: {}, total: 0, klassCounts: {}, roster: [], since: "", sheetUrl: book.getUrl() };

  const rows = report.getRange(2, 1, lastRow - 1, REPORT_HEADERS.length).getValues();
  const counts = {};
  const batches = {};
  const klassBatches = {};
  const roster = {};
  let since = "";
  rows.forEach(function (row) {
    if (String(row[2]) !== examId) return;
    if (ver && String(row[3]) !== ver) return;
    const klass = String(row[4] || "");
    if (klassFilter && klass !== klassFilter) return;
    const qid = Number(row[6]);
    if (!qid) return;
    counts[qid] = (counts[qid] || 0) + 1;
    const batch = String(row[1] || "");
    batches[batch] = true;
    klassBatches[klass] = klassBatches[klass] || {};
    klassBatches[klass][batch] = true;
    const seat = String(row[5] || "");
    if (seat) roster[klass + "-" + seat] = true;
    const stamp = String(row[0] || "");
    if (stamp && (!since || stamp < since)) since = stamp;
  });

  const klassCounts = {};
  Object.keys(klassBatches).forEach(function (name) {
    klassCounts[name] = Object.keys(klassBatches[name]).length;
  });

  return {
    ok: true,
    counts: counts,
    total: Object.keys(batches).length,
    klassCounts: klassCounts,
    roster: Object.keys(roster).sort(),
    since: since,
    sheetUrl: book.getUrl()
  };
}

function clear_(book, body) {
  const check = auth_(book, body.key);
  if (!check.ok) return { ok: false, error: "auth" };

  const level = Number(body.level) || 1;
  const examId = String(body.examId || "").trim();
  if (level === 3 && String(body.confirm || "").trim() !== "確定清空") {
    return { ok: false, error: "confirm" };
  }

  const report = book.getSheetByName(SHEET_REPORT);
  const history = book.getSheetByName(SHEET_HISTORY);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  const lastRow = report.getLastRow();
  const rows = lastRow > 1 ? report.getRange(2, 1, lastRow - 1, REPORT_HEADERS.length).getValues() : [];
  const keep = [];
  const archived = [];
  rows.forEach(function (row) {
    const isTarget = (level === 3) || String(row[2]) === examId;
    if (!isTarget) {
      keep.push(row);
      return;
    }
    archived.push([now, level].concat(row));
  });
  if (archived.length) {
    history.getRange(history.getLastRow() + 1, 1, archived.length, HISTORY_HEADERS.length).setValues(archived);
  }

  if (lastRow > 1) report.getRange(2, 1, lastRow - 1, REPORT_HEADERS.length).clearContent();
  if (keep.length) {
    report.getRange(2, 1, keep.length, REPORT_HEADERS.length).setValues(keep);
  }

  if (level === 2) {
    const historyLast = history.getLastRow();
    if (historyLast > 1) {
      const historyRows = history.getRange(2, 1, historyLast - 1, HISTORY_HEADERS.length).getValues();
      const historyKeep = historyRows.filter(function (row) {
        return String(row[4]) !== examId;
      });
      history.getRange(2, 1, historyLast - 1, HISTORY_HEADERS.length).clearContent();
      if (historyKeep.length) {
        history.getRange(2, 1, historyKeep.length, HISTORY_HEADERS.length).setValues(historyKeep);
      }
    }
  }

  return { ok: true, archived: archived.length, level: level, since: now };
}
