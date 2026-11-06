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
const TOKEN_DAYS = 90;
const CACHE_TTL = 30;
const BOARD_TTL = 10;
const REPORT_HEADERS = ["時間", "批次", "考卷", "版本", "班級", "座號", "題號", "客戶端"];
const HISTORY_HEADERS = ["清空時間", "層級", "時間", "批次", "考卷", "版本", "班級", "座號", "題號", "客戶端"];
const TOKEN_HEADERS = ["權杖", "到期時間", "建立時間"];
const PROGRESS_HEADERS = ["更新時間", "考卷", "版本", "班級", "座號", "已完成", "應完成", "客戶端"];

// 班級與座號先正規化（去空白、全形轉半形），避免同一位學生被寫成兩列
function squash_(value) {
  return String(value == null ? "" : value)
    .replace(/[\uFF01-\uFF5E]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
    .replace(/\s+/g, "");
}

// 座號再補一位數（5 與 05 視為同一個人），避免同一位學生被記成兩列
function squashSeat_(value) {
  const seat = squash_(value);
  if (/^\d{1,2}$/.test(seat)) return ("0" + seat).slice(-2);
  return seat;
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) || "{}";
    const body = JSON.parse(raw);
    const action = String(body.action || "");
    const book = SpreadsheetApp.getActiveSpreadsheet();
    if (!book) return json_({ ok: false, error: "找不到試算表，請以容器綁定方式建立腳本" });
    ensureOnce_(book);
    purgeOnce_(book);

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
    if (action === "migrate") {
      const migrateCheck = guard_(book, body);
      if (!migrateCheck.ok) return json_({ ok: false, error: "auth" });
      return json_(migrate_(book, body));
    }
    if (action === "info") return json_({ ok: true, ver: "hotspot-9" });
    return json_({ ok: false, error: "unknown-action" });
  } catch (error) {
    return json_({ ok: false, error: String((error && error.message) || error) });
  }
}

function doGet() {
  return json_({ ok: true, ver: "hotspot-9" });
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

function ensureOnce_(book) {
  if (cache_().get("init:ok")) return;
  ensureSheets_(book);
  cache_().put("init:ok", "1", 21600);
}

function purgeOnce_(book) {
  if (cache_().get("purge:ok")) return;
  purgeExpiredTokens_(book);
  cache_().put("purge:ok", "1", 600);
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
  const klass = squash_(body.klass);
  const seat = squashSeat_(body.seat);
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
    let hit = -1;
    const extra = [];
    for (let index = 0; index < rows.length; index += 1) {
      if (String(rows[index][1]) !== examId) continue;
      if (squash_(rows[index][3]) !== klass) continue;
      if (squashSeat_(rows[index][4]) !== seat) continue;
      if (hit < 0) hit = index;
      else extra.push(index);
    }
    if (hit >= 0) {
      const prevDone = Math.max(0, Number(rows[hit][5]) || 0);
      const prevTotal = Math.max(0, Number(rows[hit][6]) || 0);
      const sameRound = prevTotal === total;
      const nextDone = sameRound ? Math.max(prevDone, done) : done;
      row[5] = nextDone;
      sheet.getRange(hit + 2, 1, 1, PROGRESS_HEADERS.length).setValues([row]);
      // 同一位學生若因座號寫法不同（1 與 01、全形數字）留下多列，順手整併成一列
      if (extra.length) {
        extra.sort(function (left, right) { return right - left; }).forEach(function (index) {
          sheet.deleteRow(index + 2);
        });
      }
      bumpEpoch_(examId);
      return { ok: true, updated: true, done: nextDone, kept: nextDone !== done, merged: extra.length };
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
    const klass = squash_(row[3]);
    const seat = squashSeat_(row[4]);
    if (!klass || !seat) return;
    if (!perKlass[klass]) perKlass[klass] = {};
    const done = Number(row[5]) || 0;
    const total = Number(row[6]) || 0;
    const at = String(row[0] || "");
    const prev = perKlass[klass][seat];
    if (prev) {
      const prevRate = prev.total ? prev.done / prev.total : 0;
      const nextRate = total ? done / total : 0;
      if (nextRate < prevRate) return;
      if (nextRate === prevRate && at < prev.at) return;
    }
    perKlass[klass][seat] = {
      klass: klass,
      seat: seat,
      done: done,
      total: total,
      at: at
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
      klass: item.klass,
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
  // 以「考卷＋版本＋客戶端＋題號」去重：同一個學生同一題只記一次，
  // 之後重新整理、隔幾分鐘再按一次卷宗都不會重複計數。
  const seen = {};
  if (lastRow > 1) {
    const rows = report.getRange(2, 1, lastRow - 1, REPORT_HEADERS.length).getValues();
    rows.forEach(function (row) {
      if (String(row[2]) !== examId) return;
      if (String(row[3]) !== ver) return;
      seen[String(row[7]) + "|" + String(row[6])] = true;
    });
  }

  const now = new Date();
  const batch = client + "-" + now.getTime();
  const stampText = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const rows = [];
  const queued = {};
  qids.forEach(function (qid) {
    if (queued[qid]) return;
    if (seen[client + "|" + qid]) return;
    queued[qid] = true;
    rows.push([stampText, batch, examId, ver, klass, seat, qid, client]);
  });
  if (!rows.length) return { ok: true, added: 0, dedup: true };
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
  const people = {};
  const klassPeople = {};
  const roster = {};
  const seen = {};
  let since = "";
  rows.forEach(function (row) {
    if (String(row[2]) !== examId) return;
    if (ver && String(row[3]) !== ver) return;
    const klass = String(row[4] || "");
    if (klassFilter && klass !== klassFilter) return;
    const qid = Number(row[6]);
    if (!qid) return;
    const seat = String(row[5] || "");
    const person = String(row[7] || "") || (klass + "-" + seat);
    const dedupeKey = person + "|" + qid;
    if (seen[dedupeKey]) return;
    seen[dedupeKey] = true;
    counts[qid] = (counts[qid] || 0) + 1;
    people[person] = true;
    klassPeople[klass] = klassPeople[klass] || {};
    klassPeople[klass][person] = true;
    if (seat) roster[klass + "-" + seat] = true;
    const stamp = String(row[0] || "");
    if (stamp && (!since || stamp < since)) since = stamp;
  });

  const klassCounts = {};
  Object.keys(klassPeople).forEach(function (name) {
    klassCounts[name] = Object.keys(klassPeople[name]).length;
  });

  const payload = {
    ok: true,
    counts: counts,
    total: Object.keys(people).length,
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

// --- 一次性整理：把舊資料轉成現在的去重規則 ------------------------------------
// 在試算表用選單「考場抓鬼大師 → 整理資料」執行，或呼叫 doPost 的 migrate（需教師權杖）。
function migrate_(book, body) {
  const dryRun = !body || body.dryRun !== false;
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const result = {
    ok: true,
    dryRun: dryRun,
    report: { total: 0, keep: 0, drop: 0, backup: "" },
    progress: { total: 0, keep: 0, drop: 0, backup: "" }
  };

  const report = book.getSheetByName(SHEET_REPORT);
  if (report && report.getLastRow() > 1) {
    const rows = report.getRange(2, 1, report.getLastRow() - 1, REPORT_HEADERS.length).getValues();
    const seen = {};
    const keep = [];
    rows.forEach(function (row) {
      const qid = String(row[6] || "").trim();
      if (!qid) {
        keep.push(row);
        return;
      }
      const key = [String(row[2] || ""), String(row[3] || ""), String(row[7] || ""), qid].join("|");
      if (seen[key]) return;
      seen[key] = true;
      keep.push(row);
    });
    result.report.total = rows.length;
    result.report.keep = keep.length;
    result.report.drop = rows.length - keep.length;
    if (!dryRun && result.report.drop > 0) {
      result.report.backup = backupSheet_(book, report, SHEET_REPORT + "_備份_" + stamp);
      report.getRange(2, 1, rows.length, REPORT_HEADERS.length).setValues(keep);
      report.deleteRows(keep.length + 2, result.report.drop);
    }
  }

  const progress = book.getSheetByName(SHEET_PROGRESS);
  if (progress && progress.getLastRow() > 1) {
    const rows = progress.getRange(2, 1, progress.getLastRow() - 1, PROGRESS_HEADERS.length).getValues();
    const map = {};
    const order = [];
    rows.forEach(function (row) {
      const examId = String(row[1] || "").trim();
      const klass = squash_(row[3]);
      const seat = squashSeat_(row[4]);
      if (!examId || !klass || !seat) return;
      const key = examId + "|" + klass + "|" + seat;
      const candidate = [row[0], examId, row[2], klass, seat, Number(row[5]) || 0, Number(row[6]) || 0, row[7]];
      const prev = map[key];
      if (!prev) {
        map[key] = candidate;
        order.push(key);
        return;
      }
      const prevTotal = Number(prev[6]) || 0;
      const nextTotal = Number(candidate[6]) || 0;
      const prevRate = prevTotal ? (Number(prev[5]) || 0) / prevTotal : 0;
      const nextRate = nextTotal ? (Number(candidate[5]) || 0) / nextTotal : 0;
      if (nextRate > prevRate || (nextRate === prevRate && String(candidate[0] || "") >= String(prev[0] || ""))) {
        map[key] = candidate;
      }
    });
    const keep = order.map(function (key) { return map[key]; });
    result.progress.total = rows.length;
    result.progress.keep = keep.length;
    result.progress.drop = rows.length - keep.length;
    if (!dryRun && result.progress.drop > 0) {
      result.progress.backup = backupSheet_(book, progress, SHEET_PROGRESS + "_備份_" + stamp);
      progress.getRange(2, 1, rows.length, PROGRESS_HEADERS.length).setValues(keep);
      progress.deleteRows(keep.length + 2, result.progress.drop);
      const exams = {};
      keep.forEach(function (row) { exams[String(row[1] || "")] = true; });
      Object.keys(exams).forEach(function (id) { if (id) bumpEpoch_(id); });
    }
  }

  return result;
}

function backupSheet_(book, sheet, name) {
  const copy = sheet.copyTo(book);
  copy.setName(name);
  return name;
}

function migrateHint_(result) {
  const lines = [
    "回報表：現有 " + result.report.total + " 列，整理後保留 " + result.report.keep + " 列，刪除 " + result.report.drop + " 列",
    "進度表：現有 " + result.progress.total + " 列，整理後保留 " + result.progress.keep + " 列，刪除 " + result.progress.drop + " 列"
  ];
  if (result.dryRun) lines.push("（試跑，未改動任何資料）");
  else lines.push("已先複製備份表，再完成整理。");
  return lines.join("\n");
}

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("考場抓鬼大師")
      .addItem("整理資料（試跑，不改動）", "uiMigrateDryRun")
      .addItem("整理資料（實際執行）", "uiMigrateRun")
      .addToUi();
  } catch (error) {}
}

function uiMigrateDryRun() {
  const ui = SpreadsheetApp.getUi();
  const result = migrate_(SpreadsheetApp.getActiveSpreadsheet(), { dryRun: true });
  ui.alert("整理試跑", migrateHint_(result), ui.ButtonSet.OK);
}

function uiMigrateRun() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const preview = migrate_(book, { dryRun: true });
  const answer = ui.alert("整理資料", migrateHint_(preview) + "\n\n會先複製備份表，確定執行嗎？", ui.ButtonSet.YES_NO);
  if (answer !== ui.Button.YES) return;
  const result = migrate_(book, { dryRun: false });
  ui.alert("整理完成", migrateHint_(result), ui.ButtonSet.OK);
}
