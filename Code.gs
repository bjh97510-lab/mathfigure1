/**
 * 마법 원정대: 도형의 숲 — 학습 기록 수신용 Google Apps Script
 *
 * 1) 구글 스프레드시트 > 확장 프로그램 > Apps Script 에 이 코드를 붙여넣고 저장
 * 2) 배포 > 새 배포 > 유형: 웹 앱
 *      - 다음 사용자 인증 정보로 실행: 나
 *      - 액세스 권한이 있는 사용자: 모든 사용자
 * 3) 발급된 웹 앱 URL(…/exec)을 index.html 의 GOOGLE_SHEET_API_URL 에 붙여넣기
 *
 * 시트 구성
 *  - "학습로그": 모든 이벤트(로그인/정답/오답/도장)를 한 줄씩 기록
 *  - "학생현황": 학번별 최신 누적 기록(한 학생당 한 줄)
 */

const LOG_SHEET = '학습로그';
const SUMMARY_SHEET = '학생현황';
const LOG_HEADER = ['저장 시각', '학번', '이벤트', '단원', '문제 ID', '풀은 문제 수', '정답 수', '도토리', '칭찬도장', '앱 시각(timestamp)'];
const SUMMARY_HEADER = ['학번', '풀은 문제 수', '정답 수', '도토리', '칭찬도장', '마지막 접속', '마지막 이벤트'];

function getSheet_(name, header) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f3e4bf');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 웹앱에서 보내는 JSON POST 수신 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const id = String(d.studentId || '').trim();
    if (!id) return json_({ ok: false, error: 'studentId 없음' });
    const now = new Date();

    getSheet_(LOG_SHEET, LOG_HEADER).appendRow([
      now, id, d.event || '', d.unit || '', d.questionId || '',
      Number(d.solvedCount) || 0, Number(d.correctCount) || 0,
      Number(d.acorns) || 0, Number(d.stamps) || 0, d.timestamp || '',
    ]);

    const sum = getSheet_(SUMMARY_SHEET, SUMMARY_HEADER);
    const row = [id, Number(d.solvedCount) || 0, Number(d.correctCount) || 0, Number(d.acorns) || 0, Number(d.stamps) || 0, now, d.event || ''];
    const ids = sum.getRange(2, 1, Math.max(sum.getLastRow() - 1, 1), 1).getValues().map((r) => String(r[0]));
    const idx = ids.indexOf(id);
    if (idx >= 0) sum.getRange(idx + 2, 1, 1, row.length).setValues([row]);
    else sum.appendRow(row);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** 다른 기기에서 접속할 때 누적 기록 조회: GET ?studentId=10101 */
function doGet(e) {
  const id = String((e && e.parameter && e.parameter.studentId) || '').trim();
  if (!id) return json_({ ok: true, message: '마법 원정대 기록 서버가 동작 중입니다.' });
  const sum = getSheet_(SUMMARY_SHEET, SUMMARY_HEADER);
  const values = sum.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === id) {
      return json_({ ok: true, found: true, studentId: id, solvedCount: values[i][1], correctCount: values[i][2], acorns: values[i][3], stamps: values[i][4] });
    }
  }
  return json_({ ok: true, found: false });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
