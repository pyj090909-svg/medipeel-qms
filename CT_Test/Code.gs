/**
 * ══════════════════════════════════════════════════════
 *  QMS_03 포장재 상용성 시험 현황 — Google Apps Script
 * ──────────────────────────────────────────────────────
 *  배포 방법:
 *  1. 구글 시트 > 확장 프로그램 > Apps Script
 *  2. 이 코드를 붙여넣기
 *  3. 배포 > 새 배포 > 웹 앱
 *     - 실행 주체 : 나
 *     - 액세스 권한 : 모든 사용자
 *  4. 배포 URL을 index.html의 WEBAPP_URL에 붙여넣기
 * ══════════════════════════════════════════════════════
 */

const SS_ID       = '1trBtUq6kSTHgeG_7wTSwnklc3br3DCp9HowtfNvOuBY';
const SHEET_GID   = 0;
const HEADER_ROWS = 2;

const KEYS = [
  '시험의뢰번호','의뢰일자','의뢰자','의뢰구분','의뢰차수',
  '품번','품명','제조사',
  '시험일자','첨부파일','첨부파일여부','특기사항_제조사','시험결과',
  '판정일','특기사항_자사','최종판정',
  '최종작업자','최종작업일시',
  '첨부링크1','첨부링크2','첨부링크3','첨부링크4','첨부링크5'
];

/* ── 시트 찾기 (GID 기반) ── */
function getSheet() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === SHEET_GID) return sheets[i];
  }
  throw new Error('시트를 찾을 수 없습니다 (gid=' + SHEET_GID + ')');
}

/* ── 엔트리 포인트 ── */
function doGet(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var action = (e.parameter.action || '').trim();
    var sheet  = getSheet();

    if (action === 'get')    return ok(readAll(sheet));
    if (action === 'add')    return ok(addRow(sheet, JSON.parse(e.parameter.row)));
    if (action === 'update') return ok(updateRow(sheet, parseInt(e.parameter.rowIndex), JSON.parse(e.parameter.row)));
    if (action === 'delete') return ok(deleteRow(sheet, parseInt(e.parameter.rowIndex)));

    return err('Unknown action: ' + action);
  } catch (ex) {
    return err(ex.message);
  } finally {
    lock.releaseLock();
  }
}

/* ── READ ── */
function readAll(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROWS) return { rows: [] };

  var numRows = lastRow - HEADER_ROWS;
  var range   = sheet.getRange(HEADER_ROWS + 1, 1, numRows, KEYS.length);
  var data    = range.getValues();
  var rows    = [];

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (!r[0] && !r[1]) continue;               // 빈 행 건너뜀
    var obj = { _rowIndex: i + HEADER_ROWS + 1 };
    for (var j = 0; j < KEYS.length; j++) {
      var v = r[j];
      if (v instanceof Date) {
        obj[KEYS[j]] = (j === 17)                // 최종작업일시 → datetime
          ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
          : Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        obj[KEYS[j]] = (v === '' || v === null || v === undefined) ? '' : String(v);
      }
    }
    rows.push(obj);
  }
  return { rows: rows };
}

/* ── ADD ── */
function addRow(sheet, row) {
  var vals = KEYS.map(function(k) { return row[k] || ''; });
  sheet.appendRow(vals);
  return { success: true };
}

/* ── UPDATE ── */
function updateRow(sheet, rowIdx, row) {
  if (rowIdx <= HEADER_ROWS) throw new Error('Invalid row index');
  var vals = KEYS.map(function(k) { return row[k] !== undefined ? row[k] : ''; });
  sheet.getRange(rowIdx, 1, 1, KEYS.length).setValues([vals]);
  return { success: true };
}

/* ── DELETE ── */
function deleteRow(sheet, rowIdx) {
  if (rowIdx <= HEADER_ROWS) throw new Error('Invalid row index');
  sheet.deleteRow(rowIdx);
  return { success: true };
}

/* ── 응답 헬퍼 ── */
function ok(data)  { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function err(msg)  { return ContentService.createTextOutput(JSON.stringify({ error: msg })).setMimeType(ContentService.MimeType.JSON); }
