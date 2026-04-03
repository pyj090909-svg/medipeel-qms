// ─────────────────────────────────────────────
//  QMS_01 클레임 현황  |  Google Apps Script
//  스프레드시트 ID: 1Pv8mcM80TuVeY_mJa1Ivwr_BCpHaC--EedR8auib-aw
//
//  [배포 방법]
//  1. 이 코드를 Apps Script 편집기에 붙여넣기
//  2. addDocLinkColumn 함수 선택 → ▶ 실행  (최초 1회 — 컬럼 추가)
//  3. testGetInfo 함수 선택 → ▶ 실행       (최초 1회 — DriveApp 권한 승인)
//  4. 배포 → 배포 관리 → 새 버전 → 배포
// ─────────────────────────────────────────────

const SHEET_ID   = '1Pv8mcM80TuVeY_mJa1Ivwr_BCpHaC--EedR8auib-aw';
const SHEET_NAME = '시트1';

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
}

function makeOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 헤더에 없는 컬럼을 마지막에 자동 추가 ──
function ensureColumn(sheet, colName) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (!headers.includes(colName)) {
    sheet.getRange(1, lastCol + 1).setValue(colName);
    Logger.log('컬럼 추가됨: ' + colName + ' → 열 ' + (lastCol + 1));
  }
}

// ── 최초 1회 실행: 대책서링크 컬럼 추가 ──
function addDocLinkColumn() {
  const sheet = getSheet();
  ensureColumn(sheet, '대책서링크');
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log('현재 헤더: ' + JSON.stringify(headers));
  Logger.log('대책서링크 위치: 열 ' + (headers.indexOf('대책서링크') + 1));
}

function doGet(e) {
  try {
    const action = e.parameter.action || 'get';

    // ── 조회 ──
    if (action === 'get') {
      const sheet  = getSheet();
      const values = sheet.getDataRange().getValues();
      if (values.length === 0) return makeOutput({ headers: [], rows: [] });

      const headers = values[0];
      const rows = values.slice(1).map((row, i) => {
        const obj = { _rowIndex: i + 2 };
        headers.forEach((h, j) => {
          const val = row[j];
          obj[h] = (val instanceof Date)
            ? Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd')
            : (val ?? '');
        });
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
        const cols = ['접수일자','품번','품명','제조사','클레임대상협력업체','클레임구분',
                      '영업담당자','클레임항목','임시조치','후속조치','대책서(접수일)',
                      '대책서여부','대책서링크','클레임처리상태','최종작업자','최종작업일시'];
        sheet.appendRow(cols);
        sheet.appendRow(cols.map(h => rowData[h] ?? ''));
      } else {
        // 대책서링크 컬럼이 없으면 자동 추가
        ensureColumn(sheet, '대책서링크');
        const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        sheet.appendRow(headers.map(h => rowData[h] ?? ''));
      }
      return makeOutput({ ok: true });
    }

    // ── 수정 ──
    if (action === 'update') {
      const sheet   = getSheet();
      const rowData = JSON.parse(decodeURIComponent(e.parameter.row));
      const rowNum  = parseInt(e.parameter.rowIndex);

      // 대책서링크 컬럼이 없으면 자동 추가
      ensureColumn(sheet, '대책서링크');

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      headers.forEach((h, j) => {
        if (rowData[h] !== undefined) sheet.getRange(rowNum, j + 1).setValue(rowData[h]);
      });
      return makeOutput({ ok: true });
    }

    // ── 삭제 ──
    if (action === 'delete') {
      const sheet  = getSheet();
      sheet.deleteRow(parseInt(e.parameter.rowIndex));
      return makeOutput({ ok: true });
    }

    // ── 시트 최근 수정일 조회 ──
    if (action === 'getInfo') {
      try {
        const ss   = SpreadsheetApp.openById(SHEET_ID);
        const file = DriveApp.getFileById(ss.getId());
        const d    = file.getLastUpdated();
        const lastModified = Utilities.formatDate(d, 'Asia/Seoul', 'yyyy.MM.dd HH:mm');
        return makeOutput({ lastModified: lastModified });
      } catch(e) {
        return makeOutput({ lastModified: '-' });
      }
    }

    return makeOutput({ error: 'unknown action' });
  } catch(err) {
    return makeOutput({ error: err.message });
  }
}

function doPost(e) {
  return makeOutput({ error: 'POST not used. Use GET with action param.' });
}

// ── DriveApp 권한 승인용 테스트 함수 (최초 1회 실행) ──
function testGetInfo() {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const file = DriveApp.getFileById(ss.getId());
  const d    = file.getLastUpdated();
  Logger.log('최근 수정일: ' + Utilities.formatDate(d, 'Asia/Seoul', 'yyyy.MM.dd HH:mm'));
}
