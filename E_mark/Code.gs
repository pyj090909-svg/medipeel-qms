/* ─── 환경마크 서류 링크 관리 API (K~O열 활용) ─── */

function doGet(e) {
  var action = e.parameter.action;

  if (action === 'save_links') {
    try {
      var rowNo = String(e.parameter.rowNo).trim();
      var links = JSON.parse(e.parameter.links || '[]');
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheets()[0];
      var data = sh.getDataRange().getValues();

      var targetRow = -1;
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === rowNo) {
          targetRow = i + 1;
          break;
        }
      }

      if (targetRow === -1) {
        return ContentService
          .createTextOutput(JSON.stringify({ success: false, error: '해당 행을 찾을 수 없습니다. rowNo=' + rowNo }))
          .setMimeType(ContentService.MimeType.JSON);
      }

      // K~O열 = 11~15열 (1-based)
      for (var c = 0; c < 5; c++) {
        var val = links[c] || '';
        sh.getRange(targetRow, 11 + c).setValue(val);
      }

      return ContentService
        .createTextOutput(JSON.stringify({ success: true, row: targetRow }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ success: true, message: 'E-mark Link API ready' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return doGet(e);
}
