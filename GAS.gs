/**
 * nurseTF Google Sheets API
 * Apps Script에 붙여 넣고 SHEET_ID를 설정한 뒤 웹 앱으로 배포하세요.
 * 공개 웹 앱: 링크를 아는 누구나 데이터 읽기/추가/수정/삭제가 가능합니다.
 */
const SHEET_ID = "1gu5xf9k3EKPqqWjCkX4X92wqHDfmIZmJAkpwLUueKfU";
const SHEET_NAME = "actions";
const HEADERS = ["id", "cat", "start", "end", "name", "done", "note", "ord"];

function doGet(e) {
  const callback = (e && e.parameter && e.parameter.callback) || "";
  let result;
  try {
    const action = (e && e.parameter && e.parameter.action) || "list";
    if (action === "list") result = { ok: true, rows: listRows_() };
    else result = mutate_(action, JSON.parse((e.parameter && e.parameter.payload) || "{}"));
  } catch (err) { result = { ok: false, error: String(err && err.message || err) }; }
  const json = JSON.stringify(result).replace(/</g, "\\u003c");
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  if (!SHEET_ID || SHEET_ID.indexOf("PASTE_") === 0) throw new Error("Apps Script의 SHEET_ID를 설정하세요.");
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sh.setFrozenRows(1);
  } else {
    const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), HEADERS.length)).getValues()[0].map(String);
    if (HEADERS.some((h, i) => existing[i] !== h)) {
      throw new Error("actions 시트의 첫 행 헤더를 다음 순서로 맞추세요: " + HEADERS.join(", "));
    }
  }
  return sh;
}

function listRows_() {
  const sh = sheet_(), last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getValues()
    .filter(r => r[0] !== "")
    .map(r => ({
      id: String(r[0]), cat: String(r[1] || "선전"),
      start: dateString_(r[2]), end: dateString_(r[3]),
      name: String(r[4] || ""), done: r[5] === true || String(r[5]).toLowerCase() === "true",
      note: String(r[6] || ""), ord: Number(r[7]) || 0
    }));
}

function mutate_(action, p) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    if (action === "add") {
      const id = Utilities.getUuid();
      sh.appendRow([id, safeCat_(p.cat), validDate_(p.start), validDate_(p.end),
        String(p.name || ""), Boolean(p.done), String(p.note || ""), Number(p.ord) || Date.now()]);
      return { ok: true, id };
    }
    const id = String(p.id || "");
    if (!id) throw new Error("행 ID가 없습니다.");
    const last = sh.getLastRow();
    if (last < 2) throw new Error("수정할 행이 없습니다.");
    const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues().flat();
    const idx = ids.indexOf(id);
    if (idx < 0) throw new Error("해당 행을 찾을 수 없습니다. 새로고침 후 다시 시도하세요.");
    const rowNum = idx + 2;
    if (action === "delete") { sh.deleteRow(rowNum); return { ok: true }; }
    if (action === "update") {
      const field = String(p.field || "");
      const cols = { cat: 2, start: 3, end: 4, name: 5, done: 6, note: 7 };
      if (!Object.prototype.hasOwnProperty.call(cols, field)) throw new Error("허용되지 않은 필드입니다.");
      let value = p.value;
      if (field === "cat") value = safeCat_(value);
      if (field === "start" || field === "end") value = validDate_(value);
      if (field === "done") value = value === true || String(value) === "true";
      if (field === "name" || field === "note") value = String(value || "");
      // 단일 셀만 갱신: 다른 날짜/행 데이터는 건드리지 않습니다.
      sh.getRange(rowNum, cols[field]).setValue(value);
      return { ok: true };
    }
    throw new Error("지원하지 않는 작업입니다.");
  } finally { lock.releaseLock(); }
}

function safeCat_(v) {
  const s = String(v || "선전");
  if (["선전", "정책", "조직", "일정"].indexOf(s) < 0) throw new Error("분류 값이 올바르지 않습니다.");
  return s;
}
function validDate_(v) {
  const s = String(v || "").trim();
  if (!s) return "";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error("날짜는 YYYY-MM-DD 형식이어야 합니다.");
  return s;
}
function dateString_(v) {
  if (!v) return "";
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v))
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const s = String(v).trim(), m = s.match(/^(\d{4})[-/.]\s*(\d{1,2})[-/.]\s*(\d{1,2})/);
  return m ? m[1] + "-" + String(m[2]).padStart(2, "0") + "-" + String(m[3]).padStart(2, "0") : "";
}
