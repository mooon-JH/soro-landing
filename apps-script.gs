// SORO 사전 신청 → 구글 시트 저장용 Apps Script
// 1. 새 구글 시트를 만들고 메뉴에서 확장 프로그램 → Apps Script 를 열어요.
// 2. 이 파일 내용을 그대로 붙여넣고 저장해요.
// 3. 배포 → 새 배포 → 유형: 웹 앱
//    - 다음 사용자로서 실행: 나
//    - 액세스 권한: 모든 사용자
// 4. 나온 웹 앱 URL(https://script.google.com/macros/s/.../exec)을 index.html 의 SCRIPT_URL 에 넣어요.

const HEADERS = ["신청일시", "이름", "이메일", "하는 일", "utm_source", "utm_medium", "utm_campaign", "utm_content", "referrer"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(HEADERS);

    const p = e.parameter;
    sheet.appendRow([
      new Date(),
      clean(p.name),
      clean(p.email),
      clean(p.role),
      clean(p.utm_source),
      clean(p.utm_medium),
      clean(p.utm_campaign),
      clean(p.utm_content),
      clean(p.referrer),
    ]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// 시트 수식 주입 방지 + 길이 제한
function clean(v) {
  const s = String(v || "").slice(0, 300);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
