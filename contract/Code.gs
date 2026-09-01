// ─────────────────────────────────────────────
//  QMS 계약 관리  |  Google Apps Script
//  스프레드시트 ID: 1_y9epi5qkhYeg5-uPxcYFMSOGjRu1Vl7u9UWg06iFG4
//  시트 gid     : 1492902388  (DB-QMS-계약관리)
//
//  [배포 방법]
//  1. https://script.google.com 에서 새 프로젝트 생성
//  2. 이 코드를 전체 복사해서 붙여넣기 후 저장
//  3. ▶ testGetInfo 함수 실행 → DriveApp / Sheets 권한 승인
//  4. 우상단 [배포] → [새 배포] → 유형: 웹 앱
//     · 다음 사용자로 실행: 나
//     · 액세스 권한: 모든 사용자 (또는 조직 내부)
//  5. 발급된 /exec URL 을 index.html 의 WEBAPP_URL_FIXED 에 붙여넣기
// ─────────────────────────────────────────────

const SHEET_ID  = '1_y9epi5qkhYeg5-uPxcYFMSOGjRu1Vl7u9UWg06iFG4';
const SHEET_GID = 1492902388;   // DB-QMS-계약관리 탭 gid

// 시트명이 바뀌어도 동작하도록 gid 로 찾기
function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const target = ss.getSheets().find(s => s.getSheetId() === SHEET_GID);
  if (!target) throw new Error('gid=' + SHEET_GID + ' 시트를 찾을 수 없습니다.');
  return target;
}

function makeOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// 날짜 셀을 yyyy-MM-dd 로 정규화
function fmtCell(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return val === null || val === undefined ? '' : val;
}

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'get';

    // ── 조회 ──
    if (action === 'get') {
      const sheet  = getSheet();
      const values = sheet.getDataRange().getValues();
      if (values.length === 0) return makeOutput({ headers: [], rows: [] });

      const headers = values[0].map(h => String(h).trim());
      const rows = values.slice(1)
        .filter(row => row.some(v => v !== '' && v !== null))   // 빈 행 제외
        .map((row, i) => {
          const obj = { _rowIndex: i + 2 };
          headers.forEach((h, j) => { obj[h] = fmtCell(row[j]); });
          return obj;
        });
      return makeOutput({ headers, rows });
    }

    // ── 추가 ──
    if (action === 'add') {
      const sheet   = getSheet();
      const rowData = JSON.parse(decodeURIComponent(e.parameter.row));
      const lastCol = sheet.getLastColumn();

      if (lastCol === 0) {
        const cols = ['협력업체','계약 내용','계약 일자','계약 만료 일자',
                      '계약서 링크 1','계약서 링크 2','계약서 링크 3'];
        sheet.appendRow(cols);
        sheet.appendRow(cols.map(h => rowData[h] || ''));
      } else {
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
        sheet.appendRow(headers.map(h => rowData[h] || ''));
      }
      return makeOutput({ ok: true });
    }

    // ── 수정 ──
    if (action === 'update') {
      const sheet   = getSheet();
      const rowData = JSON.parse(decodeURIComponent(e.parameter.row));
      const rowNum  = parseInt(e.parameter.rowIndex, 10);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
      headers.forEach((h, j) => {
        if (rowData[h] !== undefined) sheet.getRange(rowNum, j + 1).setValue(rowData[h]);
      });
      return makeOutput({ ok: true });
    }

    // ── 삭제 ──
    if (action === 'delete') {
      const sheet = getSheet();
      sheet.deleteRow(parseInt(e.parameter.rowIndex, 10));
      return makeOutput({ ok: true });
    }

    // ── 시트 최근 수정일 ──
    if (action === 'getInfo') {
      try {
        const ss   = SpreadsheetApp.openById(SHEET_ID);
        const file = DriveApp.getFileById(ss.getId());
        const d    = file.getLastUpdated();
        return makeOutput({
          lastModified: Utilities.formatDate(d, 'Asia/Seoul', 'yyyy.MM.dd HH:mm')
        });
      } catch (err) {
        return makeOutput({ lastModified: '-' });
      }
    }

    return makeOutput({ error: 'unknown action: ' + action });
  } catch (err) {
    return makeOutput({ error: err.message });
  }
}

function doPost(e) {
  return makeOutput({ error: 'POST not used. Use GET with action param.' });
}

// ── 권한 승인용 테스트 함수 (최초 1회 실행) ──
function testGetInfo() {
  const sheet = getSheet();
  Logger.log('시트명: ' + sheet.getName());
  Logger.log('헤더: ' + JSON.stringify(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]));
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const file = DriveApp.getFileById(ss.getId());
  Logger.log('최근 수정일: ' + Utilities.formatDate(file.getLastUpdated(), 'Asia/Seoul', 'yyyy.MM.dd HH:mm'));
}
