/**
 * AI 페스타 참관 미션 — 제출 수신용 Apps Script
 *
 * 1) 구글 시트 「AI페스타 참관미션 제출」 만들기 → 확장 프로그램 → Apps Script
 * 2) 이 파일 내용을 Code.gs에 붙여넣고 저장
 * 3) 편집기 상단에서 setup 함수를 한 번 실행 → 「제출」「현황」「명단」 탭이 만들어진다
 * 4) 배포 → 새 배포 → 유형: 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: 모든 사용자
 * 5) 발급된 …/exec URL을 index.html의 API_URL에 넣는다
 *
 * 코드를 고친 뒤에는 "배포 관리 → 편집(연필) → 버전: 새 버전 → 배포"로 다시 배포해야 반영된다.
 * 이 방법으로 재배포하면 …/exec URL은 바뀌지 않는다.
 *
 * ── 열 구성은 앱이 정한다 ──
 * 앱(index.html)이 제출할 때 fields(항목 키·이름 목록)와 halls(특별관 이름 목록)를 함께 보낸다.
 * - 처음 보는 항목은 「제출」 탭 오른쪽 끝에 열로 추가된다.
 * - 항목 이름이 바뀌면 머리행 이름만 바뀐다. 열 위치와 기존 데이터는 그대로다.
 * - 각 열의 키는 머리행 셀의 메모에 적혀 있다. 메모를 지우거나 열 순서를 손으로 바꾸지 않는다.
 * - 특별관 목록이나 열 구성이 바뀌면 「현황」 탭을 자동으로 다시 만든다.
 * 그래서 문항·특별관을 바꿀 때는 index.html만 고치면 되고, 이 코드는 다시 배포하지 않아도 된다.
 *
 * 같은 기기에서 다시 제출하면 제출ID가 같으므로 새 행을 만들지 않고 기존 행을 덮어쓴다.
 * 「명단」 탭 A열(성명), B열(소속교)에 참가자 명단을 붙여넣으면 「현황」 탭에 미제출자가 나온다(성명 기준).
 */

const SHEET_NAME = "제출";
const DASH_NAME = "현황";
const ROSTER_NAME = "명단";

/* 서버가 직접 채우는 열 — 앱이 보내는 값으로 덮어쓰지 않는다 */
const SYSTEM = { ts: "제출시각(최종)", sid: "제출ID", rev: "제출 횟수" };

/* 머리행 메모가 없던 이전 시트를 이어 쓰기 위한 이름 → 키 대응 */
const LEGACY = {
  "제출시각(최종)": "ts", "기수": "gi", "성명": "name", "소속교": "school", "스탬프": "stamps",
  "전체 관": "total", "방문 부스 수": "boothN", "방문 부스": "booths", "학교 적용가능성": "rates",
  "관별 메모": "memos", "가져갈 기술": "take", "교직원에게 한 문장": "staff",
  "1년 내 들어올 기술·이유": "oneyear", "기기": "ua", "관람한 관(코드)": "visited",
  "적용가능성(코드)": "rateCodes", "제출ID": "sid", "제출 횟수": "rev"
};

/* 앱이 halls를 보내기 전에 setup을 실행할 때 쓰는 기본 특별관 목록 */
const DEFAULT_HALLS = [
  "AI 경진대회 특별관", "국방 AI 특별관", "독자 AI 파운데이션 모델 특별관", "정부특별관",
  "AI 인프라 특별관", "K-AI 파트너십 특별관", "피지컬 AI & 로봇 특별관", "보안 & 양자 특별관",
  "AI 반도체 특별관", "에이전틱 AI+X 특별관"
];

const MAX_LEN = 5000;
const KEY_RE = /^[A-Za-z][A-Za-z0-9_]{0,30}$/;
const MAX_FIELDS = 40;

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (!data.name || !data.school) return reply({ ok: false, error: "필수 항목 누락" });

    const fields = cleanFields(data.fields);
    const sheet = getSheet();
    const before = sheet.getLastColumn();
    const keys = ensureColumns(sheet, fields);

    const sid = String(data.sid || "");
    const at = sid ? findRow(sheet, keys, sid) : 0;
    const row = at
      ? sheet.getRange(at, 1, 1, keys.length).getValues()[0]
      : new Array(keys.length).fill("");

    fields.forEach(([key]) => { row[keys.indexOf(key)] = clean(data[key]); });
    row[keys.indexOf("ts")] = new Date();
    row[keys.indexOf("sid")] = clean(sid);
    const revAt = keys.indexOf("rev");
    row[revAt] = at ? (Number(row[revAt]) || 1) + 1 : 1;

    if (at) sheet.getRange(at, 1, 1, keys.length).setValues([row]);
    else sheet.appendRow(row);

    syncDashboard(cleanHalls(data.halls), sheet.getLastColumn() !== before);
    return reply({ ok: true, updated: !!at });
  } catch (err) {
    return reply({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function doGet() {
  return reply({ ok: true, service: "aifesta-mission" });
}

/* 편집기에서 한 번 실행. 다시 실행해도 제출 데이터와 명단은 지워지지 않는다 */
function setup() {
  getSheet();
  getRoster();
  buildDashboard(storedHalls() || DEFAULT_HALLS);
}

/* ── 제출 탭 ── */

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME, 0);
  ensureColumns(sheet, []);
  return sheet;
}

/* 머리행 키 목록을 돌려준다. 시스템 열과 fields 중 없는 열은 오른쪽에 추가하고, 이름이 바뀐 열은 머리행만 고친다 */
function ensureColumns(sheet, fields) {
  const width = sheet.getLastColumn();
  const head = width ? sheet.getRange(1, 1, 1, width) : null;
  const labels = head ? head.getValues()[0].map(String) : [];
  const notes = head ? head.getNotes()[0].map(String) : [];

  let touched = false;
  const keys = labels.map((label, i) => {
    if (notes[i]) return notes[i];
    const k = LEGACY[label] || "";
    if (k) { notes[i] = k; touched = true; }
    return k;
  });

  const want = [["ts", SYSTEM.ts]].concat(fields, [["sid", SYSTEM.sid], ["rev", SYSTEM.rev]]);
  want.forEach(([key, label]) => {
    const i = keys.indexOf(key);
    if (i < 0) {
      keys.push(key); labels.push(label); notes.push(key); touched = true;
    } else if (label && labels[i] !== label && !(key in SYSTEM)) {
      labels[i] = label; touched = true;
    }
  });

  if (touched) {
    const r = sheet.getRange(1, 1, 1, keys.length);
    r.setValues([labels]).setNotes([notes]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return keys;
}

function findRow(sheet, keys, sid) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const col = keys.indexOf("sid") + 1;
  const hit = sheet.getRange(2, col, last - 1, 1).createTextFinder(sid).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function headerKeys() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const w = sheet ? sheet.getLastColumn() : 0;
  return w ? sheet.getRange(1, 1, 1, w).getNotes()[0].map(String) : [];
}

/* ── 명단 탭 ── */

function getRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ROSTER_NAME);
  if (!sh) {
    sh = ss.insertSheet(ROSTER_NAME);
    sh.getRange(1, 1, 1, 2).setValues([["성명", "소속교"]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ── 현황 탭 ── */

function storedHalls() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("halls") || "null"); }
  catch (_) { return null; }
}

/* 특별관 목록이 바뀌었거나 열이 늘었을 때만 현황 탭을 다시 만든다 */
function syncDashboard(halls, columnsChanged) {
  const props = PropertiesService.getScriptProperties();
  const prev = props.getProperty("halls");
  const next = halls ? JSON.stringify(halls) : prev;
  const hasDash = !!SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DASH_NAME);
  if (next !== prev || columnsChanged || !hasDash) {
    if (next) props.setProperty("halls", next);
    getRoster();
    buildDashboard(next ? JSON.parse(next) : DEFAULT_HALLS);
  }
}

function colLetter(n) {
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function buildDashboard(halls) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DASH_NAME);
  if (!sh) sh = ss.insertSheet(DASH_NAME, 1);
  sh.clear();

  const keys = headerKeys();
  const S = `'${SHEET_NAME}'!`, R = `'${ROSTER_NAME}'!`;
  const col = (key) => {
    const i = keys.indexOf(key);
    if (i < 0) return "";
    const L = colLetter(i + 1);
    return `${S}${L}2:${L}`;
  };
  const name = col("name"), ts = col("ts"), stamps = col("stamps"), booth = col("boothN");
  const visited = col("visited"), rates = col("rateCodes");
  const rName = R + "A2:A", rBoth = R + "A2:B";
  const need = (...cols) => cols.every(Boolean);

  sh.getRange("A1").setValue("AI 페스타 참관 미션 제출 현황").setFontWeight("bold").setFontSize(14);

  const summary = [
    ["제출 인원(성명 기준, 중복 제외)", need(name) ? `=IF(COUNTA(${name})=0,0,COUNTUNIQUE(FILTER(${name},${name}<>"")))` : ""],
    ["제출 행 수(기기 기준)",            need(name) ? `=COUNTA(${name})` : ""],
    ["명단 인원",                        `=COUNTA(${rName})`],
    ["미제출 인원",                      need(name) ? `=IF(B5=0,"명단 탭 입력 필요",SUMPRODUCT((${rName}<>"")*ISNA(MATCH(${rName},${name},0))))` : ""],
    ["마지막 제출",                      need(ts) ? `=IF(B4=0,"",MAX(${ts}))` : ""],
    ["평균 스탬프",                      need(stamps) ? `=IFERROR(ROUND(AVERAGE(${stamps}),1),"")` : ""],
    ["평균 방문 부스 수",                need(booth) ? `=IFERROR(ROUND(AVERAGE(${booth}),1),"")` : ""]
  ];
  sh.getRange(3, 1, summary.length, 2).setValues(summary);
  sh.getRange("B7").setNumberFormat("m/d hh:mm");
  sh.getRange("A3:A9").setFontWeight("bold");

  const top = 11;
  sh.getRange(top, 1, 1, 5).setValues([["특별관", "관람", "적용가능성 상", "적용가능성 중", "적용가능성 하"]])
    .setFontWeight("bold").setBackground("#EEEEEE");
  const cnt = (tag, range) => range ? `=SUMPRODUCT(--ISNUMBER(FIND("[${tag}]",${range})))` : "";
  const rows = halls.map((h, i) => {
    const code = String(i + 1).padStart(2, "0");
    return [`${code} ${h}`, cnt(code, visited), cnt(code + "상", rates), cnt(code + "중", rates), cnt(code + "하", rates)];
  });
  if (rows.length) sh.getRange(top + 1, 1, rows.length, 5).setValues(rows);

  sh.getRange("G2").setValue("미제출자(명단 대조)").setFontWeight("bold");
  sh.getRange("G3").setValue(need(name)
    ? `=IF(COUNTA(${rName})=0,"명단 탭 A열 성명, B열 소속교를 붙여넣으면 여기에 표시됩니다",` +
      `IFERROR(FILTER(${rBoth},${rName}<>"",ISNA(MATCH(${rName},${name},0))),"모두 제출"))`
    : "");

  sh.setColumnWidth(1, 260);
  sh.setColumnWidths(2, 4, 100);
  sh.setColumnWidth(7, 110);
  sh.setColumnWidth(8, 180);
  sh.setFrozenRows(1);
}

/* ── 입력 정리 ── */

function cleanFields(list) {
  if (!Array.isArray(list)) return [];
  const seen = {};
  return list.filter((f) =>
    Array.isArray(f) && KEY_RE.test(String(f[0])) && !(f[0] in SYSTEM) && !seen[f[0]] && (seen[f[0]] = true)
  ).slice(0, MAX_FIELDS).map(([k, label]) => [String(k), String(label || k).slice(0, 40)]);
}

function cleanHalls(list) {
  if (!Array.isArray(list) || !list.length) return null;
  return list.slice(0, 30).map((h) => String(h).slice(0, 40));
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
