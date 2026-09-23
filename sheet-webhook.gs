/**
 * Neev enquiry form → Google Sheet
 *
 * 1. Create a Google Sheet named "Neev Enquiries".
 * 2. Extensions → Apps Script. Replace Code.gs with this file. Save.
 * 3. Deploy → New deployment → Web app
 *      Execute as: Me
 *      Who has access: Anyone
 * 4. Copy the Web app URL into sheet-config.js
 *
 * If you edit this file later: Deploy → Manage deployments → pencil → New version.
 */

var SPREADSHEET_ID = '12ZIosKZGStTP_qgDqXgYve_1WW3alKj3yIeqn4yYQSM';
var SHEET_NAME = 'Enquiries';
var HEADERS = [
  'Timestamp',
  'Full name',
  'Mobile',
  'Email',
  'Enquiring for',
  'Location',
  'Reason',
  'Details',
  'Preferred contact'
];

function doPost(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var data = parseBody_(e);
      var sheet = getEnquirySheet_();
      ensureHeaders_(sheet);
      sheet.appendRow([
        new Date(),
        data.fname || '',
        data.fphone || '',
        data.femail || '',
        data['for'] || '',
        data.loc || '',
        data.reason || '',
        data.fdetail || '',
        data.pref || ''
      ]);
    } finally {
      lock.releaseLock();
    }
    return json_({ status: 'ok' });
  } catch (err) {
    return json_({ status: 'error', message: String(err) });
  }
}

function doGet() {
  return json_({ status: 'ok', service: 'neev-enquiries' });
}

function parseBody_(e) {
  if (!e) return {};
  if (e.parameter && (e.parameter.fname || e.parameter.fphone || e.parameter.fdetail)) {
    return e.parameter;
  }
  if (e.postData && e.postData.contents) {
    var raw = String(e.postData.contents).trim();
    if (!raw) return {};
    try {
      return JSON.parse(raw);
    } catch (err) {
      return {};
    }
  }
  return {};
}

function getEnquirySheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss && SPREADSHEET_ID) {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  if (!ss) {
    throw new Error('Could not open the Neev Enquiries spreadsheet.');
  }
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  return sheet;
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
