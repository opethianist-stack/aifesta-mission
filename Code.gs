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
 * 「명단」 탭 A열(성명), B열(소속교)에 참가자 명단을 붙여넣으면 「현황」 탭에 미제출자가 나온다.
 * 대조는 성명 기준이다.
 */

const SHEET_NAME = "제출";
const DASH_NAME = "현황";
const ROSTER_NAME = "명단";

/* 열 순서를 바꾸면 기존 행과 어긋난다. 새 항목은 맨 뒤에 붙인다 */
const COLUMNS = [
  ["ts",        "제출시각"],
  ["gi",        "기수"],
  ["name",      "성명"],
  ["school",    "소속교"],
  ["stamps",    "스탬프"],
  ["total",     "전체 관"],
  ["boothN",    "부스 수"],
  ["booths",    "들른 부스"],
  ["rates",     "학교 적용 가능성"],
  ["memos",     "관별 메모"],
  ["take1",     "가져갈 기술 1"],
  ["take2",     "가져갈 기술 2"],
  ["take3",     "가져갈 기술 3"],
  ["staff",     "교직원에게 한 문장"],
  ["oneyear",   "1년 내 들어올 기술·이유"],
  ["ua",        "기기"],
  ["visited",   "관람한 관(코드)"],
  ["rateCodes", "적용 가능성(코드)"]
];

/* index.html의 HALLS 순서와 같아야 한다 */
const HALLS = [
  "AI 경진대회 특별관",
  "국방 AI 특별관",
  "독자 AI 파운데이션 모델 특별관",
  "정부특별관",
  "AI 인프라 특별관",
  "K-AI 파트너십 특별관",
  "피지컬 AI & 로봇 특별관",
  "보안 & 양자 특별관",
  "AI 반도체 특별관",
  "에이전틱 AI+X 특별관"
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
    sheet.appendRow(COLUMNS.map(([key]) => clean(data[key])));
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

/* 편집기에서 한 번 실행. 다시 실행해도 제출 데이터와 명단은 지워지지 않는다 */
function setup() {
  getSheet();
  getRoster();
  buildDashboard();
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME, 0);
  const labels = COLUMNS.map(([, label]) => label);
  const head = sheet.getRange(1, 1, 1, labels.length);
  if (head.getValues()[0].join("\u0001") !== labels.join("\u0001")) {
    head.setValues([labels]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

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

function colLetter(key) {
  const idx = COLUMNS.findIndex(([k]) => k === key);
  let n = idx + 1, s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function buildDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DASH_NAME);
  if (!sh) sh = ss.insertSheet(DASH_NAME, 1);
  sh.clear();

  const S = `'${SHEET_NAME}'!`, R = `'${ROSTER_NAME}'!`;
  const name = S + colLetter("name") + "2:" + colLetter("name");
  const ts = S + colLetter("ts") + "2:" + colLetter("ts");
  const stamps = S + colLetter("stamps") + "2:" + colLetter("stamps");
  const booth = S + colLetter("boothN") + "2:" + colLetter("boothN");
  const visited = S + colLetter("visited") + "2:" + colLetter("visited");
  const rates = S + colLetter("rateCodes") + "2:" + colLetter("rateCodes");
  const rName = R + "A2:A", rBoth = R + "A2:B";

  sh.getRange("A1").setValue("AI 페스타 참관 미션 제출 현황").setFontWeight("bold").setFontSize(14);

  const summary = [
    ["제출 인원(성명 기준, 중복 제외)", `=IFERROR(COUNTUNIQUE(FILTER(${name},${name}<>"")),0)`],
    ["총 제출 건수(다시 제출 포함)",     `=COUNTA(${name})`],
    ["명단 인원",                        `=COUNTA(${rName})`],
    ["미제출 인원",                      `=IF(B5=0,"명단 탭 입력 필요",SUMPRODUCT((${rName}<>"")*ISNA(MATCH(${rName},${name},0))))`],
    ["마지막 제출",                      `=IF(B4=0,"",MAX(${ts}))`],
    ["평균 스탬프(10곳 중)",             `=IFERROR(ROUND(AVERAGE(${stamps}),1),"")`],
    ["평균 들른 부스 수",                `=IFERROR(ROUND(AVERAGE(${booth}),1),"")`]
  ];
  sh.getRange(3, 1, summary.length, 2).setValues(summary);
  sh.getRange("B7").setNumberFormat("m/d hh:mm");
  sh.getRange("A3:A9").setFontWeight("bold");

  const top = 11;
  sh.getRange(top, 1, 1, 5).setValues([["특별관", "관람", "적용 상", "적용 중", "적용 하"]])
    .setFontWeight("bold").setBackground("#EEEEEE");
  const rows = HALLS.map((h, i) => {
    const code = String(i + 1).padStart(2, "0");
    const cnt = (tag) => `=SUMPRODUCT(--ISNUMBER(FIND("[${tag}]",${tag.length > 2 ? rates : visited})))`;
    return [`${code} ${h}`, cnt(code), cnt(code + "상"), cnt(code + "중"), cnt(code + "하")];
  });
  sh.getRange(top + 1, 1, rows.length, 5).setValues(rows);

  sh.getRange("G2").setValue("미제출자(명단 대조)").setFontWeight("bold");
  sh.getRange("G3").setValue(
    `=IF(COUNTA(${rName})=0,"명단 탭 A열 성명, B열 소속교를 붙여넣으면 여기에 표시됩니다",` +
    `IFERROR(FILTER(${rBoth},${rName}<>"",ISNA(MATCH(${rName},${name},0))),"모두 제출"))`
  );

  sh.setColumnWidth(1, 260);
  sh.setColumnWidths(2, 4, 80);
  sh.setColumnWidth(7, 110);
  sh.setColumnWidth(8, 180);
  sh.setFrozenRows(1);
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
