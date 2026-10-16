function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('actions');
  var lastRow = sheet.getLastRow();
  var rows = [];
  
  if (lastRow > 1) {
    var values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    for (var i = 0; i < values.length; i++) {
      rows.push({
        id: String(values[i][0]),
        cat: String(values[i][1] || '선전'),
        start: values[i][2] ? formatDate(values[i][2]) : '',
        end: values[i][3] ? formatDate(values[i][3]) : '',
        name: String(values[i][4] || ''),
        done: Boolean(values[i][5]),
        note: String(values[i][6] || '')
      });
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('actions');
  var postData = JSON.parse(e.postData.contents);
  var action = postData.action;
  
  if (action === 'save_all') {
    var newRows = postData.rows || [];
    
    // 헤더 제외 기존 데이터 삭제 후 일괄 저장
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
    }
    
    if (newRows.length > 0) {
      var dataToSet = newRows.map(function(r) {
        return [r.id, r.cat, r.start, r.end, r.name, r.done, r.note];
      });
      sheet.getRange(2, 1, dataToSet.length, 7).setValues(dataToSet);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ success: false }))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDate(dateVal) {
  if (dateVal instanceof Date) {
    var y = dateVal.getFullYear();
    var m = String(dateVal.getMonth() + 1).padStart(2, '0');
    var d = String(dateVal.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  return String(dateVal).slice(0, 10);
}
