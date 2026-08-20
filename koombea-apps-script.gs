/**
 * Koombea loan survey — Google Apps Script backend.
 *
 * Receives POST submissions from koombea.html and appends one row per response
 * to the Koombea results Sheet. Kept as a SEPARATE Apps Script project from the
 * other surveys so it writes only to Koombea's own Sheet.
 *
 * SETUP
 * -----
 * 1. Open the Koombea results Sheet:
 *      https://docs.google.com/spreadsheets/d/18pMDZtOF48SaX9Mm3SYGtCesRklSn_TD4Yc0M1g0Odg/edit
 * 2. Extensions > Apps Script. Delete the starter code and paste this whole file.
 * 3. Deploy > New deployment > type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 4. Copy the resulting web app URL (ends in /exec) and paste it into
 *    koombea.html as the SCRIPT_URL value.
 * 5. For later edits, use Manage deployments > (pencil) > Version: New version
 *    so the /exec URL stays the same.
 */

// The Koombea results Sheet this script writes to.
var SHEET_ID = '18pMDZtOF48SaX9Mm3SYGtCesRklSn_TD4Yc0M1g0Odg';
var TAB_NAME = 'Responses';

// Preferred column order. Any field not listed here is appended after these,
// so the sheet still captures new fields without code changes.
var FIELD_ORDER = [
  'timestamp',
  'company',
  'language',
  'country',
  'credit_applied',
  'interest_rate',
  'usd_income_barrier',
  'employer_advance',
  'interest_level',
  'loan_amount',
  'loan_purpose',
  'auto_deduction',
  'stablecoin_familiarity',
  'switch_willingness',
  'interest_level_post',
  'pilot_optin',
  'name',
  'email',
  'comments'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var params = (e && e.parameter) ? e.parameter : {};
    var data = {};
    Object.keys(params).forEach(function (k) { data[k] = params[k]; });
    data.timestamp = new Date();

    var sheet = getSheet_();
    var headers = ensureHeaders_(sheet, data);

    var row = headers.map(function (h) {
      var v = data[h];
      return (v === undefined || v === null) ? '' : v;
    });
    sheet.appendRow(row);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

// Simple health check when the URL is opened in a browser.
function doGet() {
  return json_({ ok: true, service: 'koombea-survey' });
}

function getSheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TAB_NAME);
  if (!sheet) sheet = ss.insertSheet(TAB_NAME);
  return sheet;
}

// Make sure the header row exists and contains every key in `data`.
// Returns the current ordered list of header names.
function ensureHeaders_(sheet, data) {
  var lastCol = sheet.getLastColumn();
  var headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(function (h) { return h !== ''; })
    : [];

  if (headers.length === 0) {
    headers = FIELD_ORDER.filter(function (f) {
      return f === 'timestamp' || data.hasOwnProperty(f);
    });
    Object.keys(data).forEach(function (k) {
      if (headers.indexOf(k) === -1) headers.push(k);
    });
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return headers;
  }

  // Add any brand-new keys as extra columns at the end.
  var added = false;
  Object.keys(data).forEach(function (k) {
    if (headers.indexOf(k) === -1) { headers.push(k); added = true; }
  });
  if (added) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return headers;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
