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
 *   權杖：教師登入權杖、到期時間、建立時間（登入一次可持續使用，逾期自動失效）
 */

const SHEET_REPORT = "回報";
const SHEET_HISTORY = "歷史";
const SHEET_SETTING = "設定";
const SHEET_TOKEN = "權杖";
const SHEET_PROGRESS = "進度";
const DEDUPE_MS = 10000;
const TOKEN_DAYS = 90;
const CACHE_TTL = 30;
const BOARD_TTL = 10;
const REPORT_HEADERS = ["時間", "批次", "考卷", "版本", "班級", "座號", "題號", "客戶端"];
const HISTORY_HEADERS = ["清空時間", "層級", "時間", "批次", "考卷", "版本", "班級", "座號", "題號", "客戶端"];
const TOKEN_HEADERS = ["權杖", "到期時間", "建立時間"];
const PROGRESS_HEADERS = ["更新時間", "考卷", "版本", "班級", "座號", "已完成", "應完成", "客戶端"];

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) || "{}";
    const body = JSON.parse(raw);
    const action = String(body.action || "");
    const book = SpreadsheetApp.getActiveSpreadsheet();
    if (!book) return json_({ ok: false, error: "找不到試算表，請以容器綁定方式建立腳本" });
    ensureSheets_(book);
    purgeExpiredTokens_(book);

    if (action === "report") return json_(report_(book, body));
    if (action === "auth") return json_(auth_(book, body.key));
    if (action === "logout") return json_(logout_(book, body.token));
    if (action === "sheetinfo") {
      const infoCheck = guard_(book, body);
      if (!infoCheck.ok) return json_({ ok: false, error: "auth" });
      return json_(sheetInfo_(book));
    }
    if (action === "progress") return json_(progress_(book, body));
    if (action === "board") return json_(board_(book, body));
    if (action === "counts") return json_(counts_(book, body));
    if (action === "clear") return json_(clear_(book, body));
    if (action === "info") return json_({ ok: true, ver: "hotspot-5" });
    return json_({ ok: false, error: "unknown-action" });
  } catch (error) {
    return json_({ ok: false, error: String((error && error.message) || error) });
  }
}

function doGet() {
  return json_({ ok: true, ver: "hotspot-5" });
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function ensureSheets_(book) {
  const report = book.getSheetByName(SHEET_REPORT) || book.insertSheet(SHEET_REPORT);
  if (report.getLastRow() === 0) report.appendRow(REPORT_HEADERS);

  const history = book.getSheetByName(SHEET_HISTORY) || book.insertSheet(SHEET_HISTORY);
  if (history.getLastRow() === 0) history.appendRow(HISTORY_HEADERS);

  const tokens = book.getSheetByName(SHEET_TOKEN) || book.insertSheet(SHEET_TOKEN);
  if (tokens.getLastRow() === 0) tokens.appendRow(TOKEN_HEADERS);

  const progress = book.getSheetByName(SHEET_PROGRESS) || book.insertSheet(SHEET_PROGRESS);
  if (progress.getLastRow() === 0) progress.appendRow(PROGRESS_HEADERS);

  const setting = book.getSheetByName(SHEET_SETTING) || book.insertSheet(SHEET_SETTING);
  if (setting.getLastRow() === 0) {
    setting.getRange("A1").setValue("教師密碼");
    setting.getRange("B1").setValue(randomSecret_(8));
    setting.getRange("A4").setValue("說明");
    setting.getRange("A5").setValue("B1 是教師密碼，可自行修改；修改後網頁需重新輸入一次。");
    setting.getRange("A6").setValue("密碼只在試算表，網頁與前端程式都不保存明文。");
    setting.getRange("A7").setValue("登入一次會發一張權杖存在「權杖」工作表，預設 90 天後失效。");
    setting.setColumnWidth(1, 140);
  }
  return report;
}

function randomSecret_(length) {
  const chars = "abcdefghijkmnpqrstuvwxyz23456789";
  let out = "";
  for (let index = 0; index < length; index += 1) {
    out += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return out;
}

function randomToken_() {
  return (
    randomSecret_(10) +
    randomSecret_(10) +
    randomSecret_(8) +
    new Date().getTime().toString(36)
  );
}

function password_(book) {
  const setting = book.getSheetByName(SHEET_SETTING);
  if (!setting || setting.getLastRow() === 0) return "";
  return String(setting.getRange("B1").getValue() || "").trim();
}

function purgeExpiredTokens_(book) {
  const sheet = book.getSheetByName(SHEET_TOKEN);
  if (!sheet || sheet.getLastRow() < 2) return;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, TOKEN_HEADERS.length).getValues();
  const now = new Date().getTime();
  const keep = rows.filter(function (row) {
    const expiry = row[1] instanceof Date ? row[1].getTime() : new Date(row[1]).getTime();
    return row[0] && expiry && expiry > now;
  });
  sheet.getRange(2, 1, rows.length, TOKEN_HEADERS.length).clearContent();
  if (keep.length) sheet.getRange(2, 1, keep.length, TOKEN_HEADERS.length).setValues(keep);
}

function issueToken_(book) {
  const sheet = book.getSheetByName(SHEET_TOKEN);
  const now = new Date();
  const expiry = new Date(now.getTime() + TOKEN_DAYS * 24 * 60 * 60 * 1000);
  const token = randomToken_();
  sheet.appendRow([token, expiry, now]);
  return { token: token, expiry: expiry.toISOString() };
}

function validToken_(book, token) {
  if (!token) return false;
  const sheet = book.getSheetByName(SHEET_TOKEN);
  if (!sheet || sheet.getLastRow() < 2) return false;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, TOKEN_HEADERS.length).getValues();
  const now = new Date().getTime();
  for (let index = 0; index < rows.length; index += 1) {
    if (String(rows[index][0]) !== String(token)) continue;
    const expiry = rows[index][1] instanceof Date ? rows[index][1].getTime() : new Date(rows[index][1]).getTime();
    return Boolean(expiry) && expiry > now;
  }
  return false;
}

function auth_(book, key) {
  const expected = password_(book);
  if (!expected) return { ok: false, error: "尚未設定密碼" };
  if (String(key || "").trim() !== expected) return { ok: false, error: "auth" };
  const issued = issueToken_(book);
  return { ok: true, token: issued.token, expiry: issued.expiry, sheetUrl: book.getUrl() };
}

function logout_(book, token) {
  const sheet = book.getSheetByName(SHEET_TOKEN);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true };
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, TOKEN_HEADERS.length).getValues();
  const keep = rows.filter(function (row) {
    return String(row[0]) !== String(token || "");
  });
  sheet.getRange(2, 1, rows.length, TOKEN_HEADERS.length).clearContent();
  if (keep.length) sheet.getRange(2, 1, keep.length, TOKEN_HEADERS.length).setValues(keep);
  return { ok: true };
}

function cache_() {
  return CacheService.getScriptCache();
}

function epochOf_(examId) {
  const saved = cache_().get("epoch:" + examId);
  return saved || "1";
}

function bumpEpoch_(examId) {
  const next = String(Number(epochOf_(examId)) + 1);
  cache_().put("epoch:" + examId, next, 21600);
}

function progress_(book, body) {
  const examId = String(body.examId || "").trim();
  const ver = String(body.ver || "").trim();
  const klass = String(body.klass || "").trim();
  const seat = String(body.seat || "").trim();
  const client = String(body.client || "").trim();
  const done = Math.max(0, Number(body.done) || 0);
  const total = Math.max(0, Number(body.total) || 0);
  if (!examId || !klass || !seat || !total) return { ok: false, error: "缺少資料" };

  const sheet = book.getSheetByName(SHEET_PROGRESS);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const row = [now, examId, ver, klass, seat, done, total, client];
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const rows = sheet.getRange(2, 1, lastRow - 1, PROGRESS_HEADERS.length).getValues();
    for (let index = 0; index < rows.length; index += 1) {
      if (String(rows[index][1]) === examId && String(rows[index][3]) === klass && String(rows[index][4]) === seat) {
        sheet.getRange(index + 2, 1, 1, PROGRESS_HEADERS.length).setValues([row]);
        bumpEpoch_(examId);
        return { ok: true, updated: true };
      }
    }
  }
  sheet.appendRow(row);
  bumpEpoch_(examId);
  return { ok: true, updated: false };
}

function board_(book, body) {
  const check = guard_(book, body);
  if (!check.ok) return { ok: false, error: "auth" };

  const examId = String(body.examId || "").trim();
  const ver = String(body.ver || "").trim();
  const klassFilter = String(body.klass || "").trim();
  const cacheKey = "board:" + epochOf_(examId) + ":" + examId + "|" + ver + "|" + klassFilter;
  const cached = cache_().get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (parsed && parsed.ok) return parsed;
    } catch (error) {}
  }

  const sheet = book.getSheetByName(SHEET_PROGRESS);
  const lastRow = sheet ? sheet.getLastRow() : 0;
  if (lastRow < 2) {
    const empty = { ok: true, percent: 0, done: 0, total: 0, students: [], started: 0, finished: 0 };
    cache_().put(cacheKey, JSON.stringify(empty), BOARD_TTL);
    return empty;
  }

  const rows = sheet.getRange(2, 1, lastRow - 1, PROGRESS_HEADERS.length).getValues();
  const perKlass = {};
  rows.forEach(function (row) {
    if (String(row[1]) !== examId) return;
    if (ver && String(row[2]) !== ver) return;
    const klass = String(row[3] || "");
    const seat = String(row[4] || "");
    if (!klass || !seat) return;
    if (!perKlass[klass]) perKlass[klass] = {};
    perKlass[klass][seat] = {
      seat: seat,
      done: Number(row[5]) || 0,
      total: Number(row[6]) || 0
    };
  });

  const klassList = Object.keys(perKlass).sort().map(function (klass) {
    const seats = Object.keys(perKlass[klass]);
    let klassDone = 0;
    let klassTotal = 0;
    seats.forEach(function (seat) {
      klassDone += perKlass[klass][seat].done;
      klassTotal += perKlass[klass][seat].total;
    });
    return { klass: klass, percent: klassTotal ? Math.round(klassDone / klassTotal * 100) : 0, students: seats.length };
  });

  const students = {};
  Object.keys(perKlass).forEach(function (klass) {
    if (klassFilter && klass !== klassFilter) return;
    Object.keys(perKlass[klass]).forEach(function (seat) {
      students[klass + "-" + seat] = perKlass[klass][seat];
    });
  });

  let done = 0;
  let total = 0;
  let started = 0;
  let finished = 0;
  const list = [];
  Object.keys(students).forEach(function (key) {
    const item = students[key];
    done += item.done;
    total += item.total;
    if (item.done > 0) started += 1;
    if (item.total > 0 && item.done >= item.total) finished += 1;
    list.push({
      seat: item.seat,
      rate: item.total ? Math.round(item.done / item.total * 100) : 0
    });
  });
  list.sort(function (left, right) {
    return Number(left.seat) - Number(right.seat) || String(left.seat).localeCompare(String(right.seat));
  });

  const payload = {
    ok: true,
    percent: total ? Math.round(done / total * 100) : 0,
    done: done,
    total: total,
    students: list,
    klassList: klassList,
    klass: klassFilter,
    started: started,
    finished: finished
  };
  cache_().put(cacheKey, JSON.stringify(payload), BOARD_TTL);
  return payload;
}

function sheetInfo_(book) {
  return { ok: true, title: book.getName(), url: book.getUrl() };
}

function guard_(book, body) {
  if (body && body.token && validToken_(book, body.token)) return { ok: true };
  if (body && body.key && String(body.key).trim() === password_(book)) return { ok: true };
  return { ok: false, error: "auth" };
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
  bumpEpoch_(examId);
  return { ok: true, added: rows.length };
}

function counts_(book, body) {
  const check = guard_(book, body);
  if (!check.ok) return { ok: false, error: "auth" };

  const examId = String(body.examId || "").trim();
  const ver = String(body.ver || "").trim();
  const klassFilter = String(body.klass || "").trim();
  const cacheKey = "counts:" + epochOf_(examId) + ":" + examId + "|" + ver + "|" + klassFilter;
  const cached = cache_().get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      if (parsed && parsed.ok) return parsed;
    } catch (error) {}
  }

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

  const payload = {
    ok: true,
    counts: counts,
    total: Object.keys(batches).length,
    klassCounts: klassCounts,
    roster: Object.keys(roster).sort(),
    since: since,
    sheetUrl: book.getUrl()
  };
  cache_().put(cacheKey, JSON.stringify(payload), CACHE_TTL);
  return payload;
}

function clear_(book, body) {
  const check = guard_(book, body);
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

  const progress = book.getSheetByName(SHEET_PROGRESS);
  if (progress && progress.getLastRow() > 1) {
    const progressLast = progress.getLastRow();
    const progressRows = progress.getRange(2, 1, progressLast - 1, PROGRESS_HEADERS.length).getValues();
    const progressKeep = progressRows.filter(function (row) {
      if (level === 3) return false;
      return String(row[1]) !== examId;
    });
    progress.getRange(2, 1, progressLast - 1, PROGRESS_HEADERS.length).clearContent();
    if (progressKeep.length) {
      progress.getRange(2, 1, progressKeep.length, PROGRESS_HEADERS.length).setValues(progressKeep);
    }
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

  bumpEpoch_(examId);
  return { ok: true, archived: archived.length, level: level, since: now };
}
