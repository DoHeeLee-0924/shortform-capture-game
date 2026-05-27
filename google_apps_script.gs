const SHEET_NAME = 'responses';
const ADMIN_TOKEN = 'change-this-admin-token';

function doPost(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss);
  const data = JSON.parse(e.postData.contents || '{}');
  appendObject_(sheet, data);
  return json_({ ok: true });
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.action !== 'list') return jsonp_({ ok: false, error: 'invalid action' }, p.callback);
  if (p.token !== ADMIN_TOKEN) return jsonp_({ ok: false, error: 'unauthorized' }, p.callback);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet_(ss);
  return jsonp_({ ok: true, rows: readObjects_(sheet) }, p.callback);
}

function getSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

function appendObject_(sheet, obj) {
  const existingHeaders = sheet.getLastColumn() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].filter(String)
    : [];
  const keys = Object.keys(obj);
  const headers = existingHeaders.slice();
  keys.forEach(k => { if (!headers.includes(k)) headers.push(k); });
  if (headers.length !== existingHeaders.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  const row = headers.map(h => obj[h] == null ? '' : String(obj[h]));
  sheet.appendRow(row);
}

function readObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(r => r.some(v => v !== '')).map(r => {
    const o = {};
    headers.forEach((h, i) => { if (h) o[h] = r[i]; });
    return o;
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function jsonp_(obj, callback) {
  const body = callback ? `${callback}(${JSON.stringify(obj)});` : JSON.stringify(obj);
  return ContentService.createTextOutput(body).setMimeType(callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
