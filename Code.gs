/* ====================================================================
   브릿지커피 단골 적립 — 구글 시트 연결 코드 (Apps Script)

   ■ 이 파일은 "구글 시트 안"에 붙여 넣는 코드예요. (웹 화면 파일이 아니에요)
     붙여 넣는 방법은 같은 폴더의 「구글시트_설정방법.md」를 따라 하세요.

   ■ 하는 일
     - 손님 화면(index.html)이 "이 번호 몇 개야?" 하고 물으면 개수만 알려 줘요.
     - 사장님 화면(admin.html)이 "적립 / 취소 / 사용 완료" 하고 시키면 표를 고쳐요.
       이때 비밀번호가 맞아야만 고칠 수 있어요. 비밀번호는 이 파일이 아니라
       "스크립트 속성"이라는 숨은 설정에 넣어요. (설정방법.md 참고)

   ■ 표(시트) 두 개를 자동으로 만들어요
     - 손님 : 번호(뒷 4자리) / 스탬프 개수 / 마지막 변경 시각
              (0개가 된 손님은 이 표에서 지워요. 개인정보를 최소로 두기 위해서예요.)
     - 기록 : 시각 / 번호 / 동작(적립·취소·사용완료) / 그때의 개수
              (실수나 이상한 일이 있었는지 나중에 확인하는 용도예요.)

   ■ ⚠ 개인정보
     - 손님 번호 뒷 4자리가 사장님 구글 계정의 이 표에 저장돼요.
     - 이 구글 시트를 다른 사람에게 "공유"하지 마세요. (공유하면 장부가 보여요.)
   ==================================================================== */

/* ---------- 고치는 곳 ---------- */
const 목표개수 = 10;                 /* 몇 개를 모으면 보상인가요? (화면 파일의 목표개수와 같게!) */
const 조회제한_분당 = 40;            /* 손님 조회를 전부 합쳐서 1분에 이 횟수까지만 받아요 (번호 훑어보기 막기) */
const 비밀번호_틀려도되는횟수 = 5;   /* 이만큼 틀리면 잠시 모든 시도를 막아요 */
const 비밀번호_잠금초 = 300;          /* 잠그는 시간(초). 300 = 5분 */

/* ---------- 아래는 보통 고치지 않아요 ---------- */
const 손님시트이름 = "손님";
const 기록시트이름 = "기록";
const 시간대 = "Asia/Seoul";

/* 처음 한 번만 직접 실행해 주세요. 표 두 개를 만들고, 구글에 "이 코드를 써도 된다"고 허락을 받아요. */
function 처음설정() {
  시트준비();
  const 비번 = PropertiesService.getScriptProperties().getProperty("비밀번호");
  if (!비번 || !/^[0-9]{4}$/.test(비번)) {
    Logger.log("표 두 개를 만들었어요. 이제 '프로젝트 설정 → 스크립트 속성'에 비밀번호(숫자 4자리)를 넣어 주세요.");
  } else {
    Logger.log("준비 끝! 비밀번호도 들어 있어요. 이제 '배포 → 새 배포 → 웹 앱'을 진행하세요.");
  }
}

function 시트준비() {
  const 문서 = SpreadsheetApp.getActiveSpreadsheet();
  let 손님 = 문서.getSheetByName(손님시트이름);
  if (!손님) {
    손님 = 문서.insertSheet(손님시트이름);
    손님.getRange(1, 1, 1, 3).setValues([["번호(뒷4자리)", "스탬프 개수", "마지막 변경"]]);
    손님.getRange("A:A").setNumberFormat("@");   /* 0012 같은 번호의 앞 0이 사라지지 않게 글자로 저장 */
    손님.setFrozenRows(1);
  }
  let 기록 = 문서.getSheetByName(기록시트이름);
  if (!기록) {
    기록 = 문서.insertSheet(기록시트이름);
    기록.getRange(1, 1, 1, 4).setValues([["시각", "번호(뒷4자리)", "동작", "그때의 개수"]]);
    기록.getRange("B:B").setNumberFormat("@");
    기록.setFrozenRows(1);
  }
  return { 손님: 손님, 기록: 기록 };
}

/* 화면에서 오는 요청을 받는 곳 */
function doPost(e) {
  let 요청;
  try { 요청 = JSON.parse(e.postData.contents); }
  catch (오류) { return 응답({ 성공: false, 오류: "요청을 읽을 수 없어요" }); }

  /* 동시에 여러 요청이 와도 순서대로 처리해서 숫자가 꼬이지 않게 해요 */
  const 잠금 = LockService.getScriptLock();
  try { 잠금.waitLock(15000); }
  catch (오류) { return 응답({ 성공: false, 오류: "지금 사용하는 사람이 많아요. 잠시 후 다시 시도해 주세요" }); }

  try { return 응답(처리(요청)); }
  catch (오류) { return 응답({ 성공: false, 오류: "서버에서 문제가 생겼어요. 잠시 후 다시 시도해 주세요" }); }
  finally { 잠금.releaseLock(); }
}

/* 주소를 그냥 열어 봤을 때: 잘 켜져 있는지만 알려 줘요 (손님 정보는 절대 안 보여 줘요) */
function doGet() { return 응답({ 성공: true, 안내: "단골 적립 서버가 잘 켜져 있어요" }); }

function 응답(내용) {
  return ContentService.createTextOutput(JSON.stringify(내용)).setMimeType(ContentService.MimeType.JSON);
}

function 처리(요청) {
  const 동작 = String(요청.동작 || "");
  const 번호 = String(요청.번호 == null ? "" : 요청.번호);

  if (동작 === "조회") return 손님조회(번호);   /* 손님 화면용: 비밀번호 없이, 개수만 */

  /* 여기서부터는 사장님 전용: 비밀번호가 맞아야 해요 */
  const 문제 = 비밀번호확인(요청.비밀번호);
  if (문제) return 문제;

  if (동작 === "로그인") return { 성공: true, 오늘횟수: 오늘횟수() };
  if (!/^[0-9]{4}$/.test(번호)) return { 성공: false, 오류: "번호는 숫자 4자리여야 해요" };

  const 시트들 = 시트준비();
  const 행 = 행찾기(시트들.손님, 번호);
  const 현재 = 개수읽기(시트들.손님, 행);

  if (동작 === "관리조회") return { 성공: true, 개수: 현재, 오늘횟수: 오늘횟수() };

  if (동작 === "적립") {
    if (현재 >= 목표개수) return { 성공: false, 오류: "이미 " + 목표개수 + "개예요. 먼저 사용 완료를 눌러 주세요", 개수: 현재, 오늘횟수: 오늘횟수() };
    개수쓰기(시트들.손님, 번호, 행, 현재 + 1);
    기록남기기(시트들.기록, 번호, "적립", 현재 + 1);
    return { 성공: true, 개수: 현재 + 1, 오늘횟수: 오늘횟수() };
  }
  if (동작 === "취소") {
    if (현재 <= 0) return { 성공: false, 오류: "뺄 스탬프가 없어요", 개수: 0, 오늘횟수: 오늘횟수() };
    개수쓰기(시트들.손님, 번호, 행, 현재 - 1);
    기록남기기(시트들.기록, 번호, "취소", 현재 - 1);
    return { 성공: true, 개수: 현재 - 1, 오늘횟수: 오늘횟수() };
  }
  if (동작 === "사용완료") {
    if (현재 < 목표개수) return { 성공: false, 오류: 목표개수 + "개를 모은 손님만 사용 완료를 할 수 있어요", 개수: 현재, 오늘횟수: 오늘횟수() };
    개수쓰기(시트들.손님, 번호, 행, 0);
    기록남기기(시트들.기록, 번호, "사용완료", 0);
    return { 성공: true, 개수: 0, 오늘횟수: 오늘횟수() };
  }
  return { 성공: false, 오류: "알 수 없는 요청이에요" };
}

/* 손님 조회: 개수만 알려 줘요. 짧은 시간에 너무 많이 물어보면 잠시 막아요. */
function 손님조회(번호) {
  if (!/^[0-9]{4}$/.test(번호)) return { 성공: false, 오류: "번호는 숫자 4자리여야 해요" };

  const 캐시 = CacheService.getScriptCache();
  const 칸 = "조회_" + Math.floor(Date.now() / 60000);
  const 횟수 = Number(캐시.get(칸) || 0) + 1;
  캐시.put(칸, String(횟수), 120);
  if (횟수 > 조회제한_분당) return { 성공: false, 오류: "지금 확인하는 손님이 많아요. 잠시 후 다시 시도해 주세요" };

  const 시트들 = 시트준비();
  return { 성공: true, 개수: 개수읽기(시트들.손님, 행찾기(시트들.손님, 번호)) };
}

/* 비밀번호 확인: 맞으면 null, 틀리면 이유를 돌려줘요 */
function 비밀번호확인(입력) {
  const 캐시 = CacheService.getScriptCache();
  const 잠금끝 = Number(캐시.get("잠금끝") || 0);
  if (Date.now() < 잠금끝) {
    return { 성공: false, 잠김: true, 남은초: Math.ceil((잠금끝 - Date.now()) / 1000), 오류: "비밀번호를 너무 많이 틀렸어요. 잠시 뒤에 다시 시도해 주세요" };
  }
  const 맞는값 = PropertiesService.getScriptProperties().getProperty("비밀번호");
  if (!맞는값 || !/^[0-9]{4}$/.test(맞는값)) {
    return { 성공: false, 오류: "구글 설정에 비밀번호(숫자 4자리)가 아직 없어요. 설정방법 3단계를 확인해 주세요" };
  }
  if (String(입력 == null ? "" : 입력) === 맞는값) { 캐시.remove("틀린횟수"); return null; }

  const 틀림 = Number(캐시.get("틀린횟수") || 0) + 1;
  if (틀림 >= 비밀번호_틀려도되는횟수) {
    캐시.put("잠금끝", String(Date.now() + 비밀번호_잠금초 * 1000), 비밀번호_잠금초);
    캐시.remove("틀린횟수");
    return { 성공: false, 잠김: true, 남은초: 비밀번호_잠금초, 오류: "비밀번호를 너무 많이 틀렸어요. 잠시 뒤에 다시 시도해 주세요" };
  }
  캐시.put("틀린횟수", String(틀림), 600);
  return { 성공: false, 오류: "비밀번호가 달라요", 남은기회: 비밀번호_틀려도되는횟수 - 틀림 };
}

/* ---------- 표 읽고 쓰기 ---------- */
function 행찾기(시트, 번호) {
  const 마지막 = 시트.getLastRow();
  if (마지막 < 2) return 0;
  const 값들 = 시트.getRange(2, 1, 마지막 - 1, 1).getValues();
  for (let i = 0; i < 값들.length; i++) {
    if (String(값들[i][0]).padStart(4, "0") === 번호) return i + 2;
  }
  return 0;
}
function 개수읽기(시트, 행) {
  return 행 ? Number(시트.getRange(행, 2).getValue()) || 0 : 0;
}
function 개수쓰기(시트, 번호, 행, 값) {
  if (값 <= 0) { if (행) 시트.deleteRow(행); return; }   /* 0개가 되면 표에서 지워요 */
  const 지금 = new Date();
  if (행) {
    시트.getRange(행, 2, 1, 2).setValues([[값, 지금]]);
  } else {
    const 새행 = 시트.getLastRow() + 1;
    시트.getRange(새행, 1).setNumberFormat("@").setValue(번호);
    시트.getRange(새행, 2, 1, 2).setValues([[값, 지금]]);
  }
}
function 기록남기기(시트, 번호, 동작, 개수) {
  const 새행 = 시트.getLastRow() + 1;
  시트.getRange(새행, 2).setNumberFormat("@").setValue(번호);
  시트.getRange(새행, 1).setValue(new Date());
  시트.getRange(새행, 3, 1, 2).setValues([[동작, 개수]]);
}

/* 오늘(한국 시간 기준) 적립한 횟수. 취소한 만큼은 빼요. */
function 오늘횟수() {
  const 기록 = 시트준비().기록;
  const 마지막 = 기록.getLastRow();
  if (마지막 < 2) return 0;
  const 시작 = Math.max(2, 마지막 - 999);   /* 하루치는 최근 1000줄 안에 충분히 들어가요 */
  const 값들 = 기록.getRange(시작, 1, 마지막 - 시작 + 1, 3).getValues();
  const 오늘 = Utilities.formatDate(new Date(), 시간대, "yyyy-MM-dd");
  let 합 = 0;
  값들.forEach(function (줄) {
    if (!(줄[0] instanceof Date)) return;
    if (Utilities.formatDate(줄[0], 시간대, "yyyy-MM-dd") !== 오늘) return;
    if (줄[2] === "적립") 합++;
    else if (줄[2] === "취소") 합--;
  });
  return Math.max(0, 합);
}
