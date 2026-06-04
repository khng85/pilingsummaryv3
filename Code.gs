// ================================
const SHEET_NAME = 'Sheet1';
const HEADERS = ['id','location','pileRef','pileType','date','init12m','ext12m','ext9m','ext6m','ext3m','actualDriven'];
const COL_ID   = 0;
const COL_LOC  = 1;
const COL_REF  = 2;
const COL_TYPE = 3;
const COL_DATE = 4;
const COL_INIT = 5;
const COL_E12  = 6;
const COL_E9   = 7;
const COL_E6   = 8;
const COL_E3   = 9;
const COL_ACT  = 10;
// ================================

function doGet(e) {
  e = e || {};                              // ← guard against null event
  var params = e.parameter || {};           // ← guard against missing parameter
  var callback = params.callback || null;   // ← safe JSONP callback read
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  // Ensure header row exists
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    var empty = JSON.stringify({ status: 'success', data: [] });
    return callback
      ? ContentService.createTextOutput(callback + '(' + empty + ')').setMimeType(ContentService.MimeType.JAVASCRIPT)
      : ContentService.createTextOutput(empty).setMimeType(ContentService.MimeType.JSON);
  }

  var rows = sheet.getDataRange().getValues();
  var data = [];

  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (row.join('').trim() === '') continue;
    data.push({
      id:           String(row[COL_ID]),
      location:     String(row[COL_LOC]),
      pileRef:      Number(row[COL_REF]),
      pileType:     String(row[COL_TYPE]),
      date:         String(row[COL_DATE]),
      init12m:      Number(row[COL_INIT]),
      ext12m:       Number(row[COL_E12]),
      ext9m:        Number(row[COL_E9]),
      ext6m:        Number(row[COL_E6]),
      ext3m:        Number(row[COL_E3]),
      actualDriven: Number(row[COL_ACT])
    });
  }

  var result = JSON.stringify({ status: 'success', data: data });
  return callback
    ? ContentService.createTextOutput(callback + '(' + result + ')').setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var payload;
  try {
    payload = JSON.parse(e.postData.contents);
  } catch(err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: 'Invalid JSON: ' + e.postData.contents })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var records = payload.data || [];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }

  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }

  if (records.length === 0) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', rowsUpdated: 0 }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var newRows = records.map(function(r) {
    return [
      r.id           || '',
      r.location     || '',
      r.pileRef      || 0,
      r.pileType     || '',
      r.date         || '',
      r.init12m      || 0,
      r.ext12m       || 0,
      r.ext9m        || 0,
      r.ext6m        || 0,
      r.ext3m        || 0,
      r.actualDriven || 0
    ];
  });

  sheet.getRange(2, 1, newRows.length, 11).setValues(newRows);

  return ContentService.createTextOutput(JSON.stringify({ status: 'success', rowsUpdated: newRows.length }))
    .setMimeType(ContentService.MimeType.JSON);
}