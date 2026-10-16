/**
 * nurseTF Google Sheets API
 */
const SHEET_ID = "1gu5xf9k3EKPqqWjCkX4X92wqHDfmIZmJAkpwLUueKfU";
const SHEET_NAME = "actions";
const HEADERS = ["id", "cat", "start", "end", "name", "done", "note", "ord"];

function doGet(e) {
  const callback = (e && e.parameter && e.parameter.callback) || "";
  let result;
  try {
    const action = (e && e.parameter && e.parameter.action) || "list";
    let payload = {};
    if (e && e.parameter && e.parameter.payload) {
      payload = JSON.parse(e.parameter.payload);
    }
    
    if (action === "list") {
      result = { ok: true, rows: listRows_() };
    } else if (action === "saveAll") {
      result = saveAllRows_(payload.rows || []);
    } else {
      result = { ok: false, error: "지원하지 않는 작업입니다." };
    }
  } catch (err) {
    result = { ok: false, error: String(err && err.message || err) };
  }
  
  const json = JSON.stringify(result).replace(/</g, "\\u003c");
  
  // JSONP 콜백 요청이 들어오면 함수 형태로 감싸서 응답
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  if (!SHEET_ID) throw new Error("SHEET_ID를 설정하세요.");
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function listRows_() {
  const sh = sheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(r => r[0] !== "")
    .map(r => ({
      id: String(r[0]),
      cat: String(r[1] || "선전"),
      start: dateString_(r[2]),
      end: dateString_(r[3]),
      name: String(r[4] || ""),
      done: r[5] === true || String(r[5]).toLowerCase() === "true",
      note: String(r[6] || ""),
      ord: Number(r[7]) || 0
    }));
}

function saveAllRows_(rows) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const last = sh.getLastRow();
    if (last > 1) {
      sh.getRange(2, 1, last - 1, HEADERS.length).clearContent();
    }
    
    if (rows.length === 0) return { ok: true };

    const outData = rows.map(r => [
      String(r.id || Utilities.getUuid()),
      String(r.cat || "선전"),
      validDate_(r.start),
      validDate_(r.end),
      String(r.name || ""),
      Boolean(r.done),
      String(r.note || ""),
      Number(r.ord) || Date.now()
    ]);

    sh.getRange(2, 1, outData.length, HEADERS.length).setValues(outData);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function validDate_(v) {
  const s = String(v || "").trim();
  if (!s) return "";
  const m = s.match(/(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : "";
}

function dateString_(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v))
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const s = String(v).trim();
  const m = s.match(/(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : "";
}
