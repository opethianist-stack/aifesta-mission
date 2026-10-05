/**
 * AI 페스타 관람 미션 — 제출 수신용 Apps Script
 *
 * 1) 구글 시트 새로 만들기 → 확장 프로그램 → Apps Script
 * 2) 이 파일 내용을 Code.gs에 붙여넣고 저장
 * 3) 배포 → 새 배포 → 유형: 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 4) 발급된 …/exec URL을 index.html의 API_URL에 넣는다
 *
 * 코드를 고친 뒤에는 "배포 관리 → 편집 → 새 버전"으로 다시 배포해야 반영된다.
 */

const SHEET_NAME = "제출";

const COLUMNS = [
  ["ts",     "제출시각"],
  ["gi",     "기수"],
  ["name",   "성명"],
  ["school", "소속교"],
  ["level",  "학교급"],
  ["best",   "인상 깊은 특별관"],
  ["idea",   "학교 경영에 가져갈 것"],
  ["stamps", "스탬프"],
  ["total",  "전체 관"],
  ["boothN", "부스 수"],
  ["booths", "들른 부스"],
  ["memos",  "관별 메모"],
  ["ua",     "기기"]
];

const MAX_LEN = 5000;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (!data.name || !data.school) return reply({ ok: false, error: "필수 항목 누락" });

    const sheet = getSheet();
    data.ts = new Date();
    const row = COLUMNS.map(([key]) => clean(data[key]));
    sheet.appendRow(row);
    return reply({ ok: true });
  } catch (err) {
    return reply({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function doGet() {
  return reply({ ok: true, service: "aifesta-mission" });
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS.map(([, label]) => label));
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight("bold");
  }
  return sheet;
}

/* 길이 제한 + 수식 주입 방지(=, +, -, @로 시작하는 문자열) */
function clean(v) {
  if (v instanceof Date || typeof v === "number") return v;
  if (v === undefined || v === null) return "";
  let s = String(v).slice(0, MAX_LEN);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function reply(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
