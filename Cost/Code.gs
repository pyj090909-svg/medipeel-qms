// ─────────────────────────────────────────────
//  QMS_05 RA 비용 관리  |  Google Apps Script
//  스프레드시트 ID: 1nW1VomsorXvXaXenqV51qqQ_XXnalNCKbAfdhve8TA0
//
//  [배포 방법]
//  1. 구글시트 → 확장 프로그램 → Apps Script 열기
//  2. 이 코드를 편집기에 붙여넣기
//  3. initSheets 함수 선택 → ▶ 실행  (최초 1회 — 시트 구조 생성)
//  4. testGetInfo 함수 선택 → ▶ 실행  (최초 1회 — DriveApp 권한 승인)
//  5. 배포 → 새 배포 → 웹 앱 → 액세스: 모든 사용자 → 배포
// ─────────────────────────────────────────────

const SHEET_ID   = '1nW1VomsorXvXaXenqV51qqQ_XXnalNCKbAfdhve8TA0';
const DATA_SHEET = 'RA_비용관리_DATA';
const CFG_SHEET  = 'RA_비용관리_설정';

function getDataSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(DATA_SHEET);
}
function getCfgSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(CFG_SHEET);
}

function makeOutput(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 최초 1회 실행: 시트 구조 생성 ──
function initSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // DATA 시트
  let ds = ss.getSheetByName(DATA_SHEET);
  if (!ds) {
    ds = ss.insertSheet(DATA_SHEET);
    const headers = [
      'No','대행사명','태스크명','등록일',
      '품의서번호','품의서링크1','품의서링크2','품의서링크3','품의금액','품의상태',
      '지급상태','지급예정일','지급완료일','지급금액',
      '결의서번호','결의서링크','결의서작성일','결의서상태',
      '메모'
    ];
    ds.getRange(1, 1, 1, headers.length).setValues([headers]);
    ds.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#0f1c45')
      .setFontColor('#ffffff');
    ds.setFrozenRows(1);
    Logger.log('DATA 시트 생성 완료');
  }

  // 설정 시트
  let cs = ss.getSheetByName(CFG_SHEET);
  if (!cs) {
    cs = ss.insertSheet(CFG_SHEET);
    // A열: 대행사 목록
    cs.getRange('A1').setValue('대행사목록').setFontWeight('bold');
    cs.getRange('A2').setValue('(주)에이비씨랩');
    cs.getRange('A3').setValue('(주)코스맥스');
    cs.getRange('A4').setValue('한국화학시험연구원');
    // B열: 품의상태
    cs.getRange('B1').setValue('품의상태').setFontWeight('bold');
    cs.getRange('B2').setValue('작성중');
    cs.getRange('B3').setValue('승인완료');
    cs.getRange('B4').setValue('반려');
    // C열: 지급상태
    cs.getRange('C1').setValue('지급상태').setFontWeight('bold');
    cs.getRange('C2').setValue('지급대기');
    cs.getRange('C3').setValue('지급완료');
    // D열: 결의서상태
    cs.getRange('D1').setValue('결의서상태').setFontWeight('bold');
    cs.getRange('D2').setValue('작성중');
    cs.getRange('D3').setValue('승인완료');
    cs.getRange('D4').setValue('반려');
    Logger.log('설정 시트 생성 완료');
  }
}

function doPost(e) { return handleRequest(e); }
function doGet(e)  { return handleRequest(e); }

function handleRequest(e) {
  try {
    const action = e.parameter.action || 'get';

    // ── 데이터 조회 ──
    if (action === 'get') {
      const sheet  = getDataSheet();
      if (!sheet) return makeOutput({ headers: [], rows: [] });
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

    // ── 설정 조회 (대행사 목록 등) ──
    if (action === 'getConfig') {
      const sheet = getCfgSheet();
      if (!sheet) return makeOutput({ agents: [], pStatus: [], gStatus: [], rStatus: [] });
      const data = sheet.getDataRange().getValues();
      const pick = (col) => data.slice(1).map(r => String(r[col] ?? '').trim()).filter(Boolean);
      return makeOutput({
        agents:  pick(0),
        pStatus: pick(1),
        gStatus: pick(2),
        rStatus: pick(3)
      });
    }

    // ── 추가 ──
    if (action === 'add') {
      const sheet   = getDataSheet();
      const rowData = JSON.parse(e.parameter.row);
      const lastCol = sheet.getLastColumn();

      if (lastCol === 0) {
        const headers = [
          'No','대행사명','태스크명','등록일',
          '품의서번호','품의서링크1','품의서링크2','품의서링크3','품의금액','품의상태',
          '지급상태','지급예정일','지급완료일','지급금액',
          '결의서번호','결의서링크','결의서작성일','결의서상태',
          '메모'
        ];
        sheet.appendRow(headers);
        sheet.appendRow(headers.map(h => rowData[h] ?? ''));
      } else {
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        // No 자동 채번
        const lastRow = sheet.getLastRow();
        if (!rowData['No'] && lastRow > 1) {
          const nos = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().map(Number).filter(n => !isNaN(n));
          rowData['No'] = nos.length ? Math.max(...nos) + 1 : 1;
        } else if (!rowData['No']) {
          rowData['No'] = 1;
        }
        sheet.appendRow(headers.map(h => rowData[h] ?? ''));
      }
      return makeOutput({ ok: true });
    }

    // ── 수정 ──
    if (action === 'update') {
      const sheet   = getDataSheet();
      const rowData = JSON.parse(e.parameter.row);
      const rowNum  = parseInt(e.parameter.rowIndex);
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      headers.forEach((h, j) => {
        if (rowData[h] !== undefined) sheet.getRange(rowNum, j + 1).setValue(rowData[h]);
      });
      return makeOutput({ ok: true });
    }

    // ── 삭제 ──
    if (action === 'delete') {
      const sheet = getDataSheet();
      sheet.deleteRow(parseInt(e.parameter.rowIndex));
      return makeOutput({ ok: true });
    }

    // ── 설정에 대행사 추가 ──
    if (action === 'addAgent') {
      const sheet = getCfgSheet();
      const name  = (e.parameter.name || '').trim();
      if (!name) return makeOutput({ error: '대행사명이 없습니다.' });
      const data = sheet.getRange('A:A').getValues().flat().filter(Boolean);
      if (data.includes(name)) return makeOutput({ ok: true, msg: '이미 존재' });
      sheet.getRange(data.length + 1, 1).setValue(name);
      return makeOutput({ ok: true });
    }

    // ── 시트 최근 수정일 ──
    if (action === 'getInfo') {
      try {
        const ss   = SpreadsheetApp.openById(SHEET_ID);
        const file = DriveApp.getFileById(ss.getId());
        const d    = file.getLastUpdated();
        const lastModified = Utilities.formatDate(d, 'Asia/Seoul', 'yyyy.MM.dd HH:mm');
        return makeOutput({ lastModified });
      } catch(err) {
        return makeOutput({ lastModified: '-' });
      }
    }

    return makeOutput({ error: 'unknown action' });
  } catch(err) {
    return makeOutput({ error: err.message });
  }
}

// ── DriveApp 권한 승인용 (최초 1회 실행) ──
function testGetInfo() {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const file = DriveApp.getFileById(ss.getId());
  const d    = file.getLastUpdated();
  Logger.log('최근 수정일: ' + Utilities.formatDate(d, 'Asia/Seoul', 'yyyy.MM.dd HH:mm'));
}
