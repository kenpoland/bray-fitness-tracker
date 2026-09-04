/**
 * BRAY FITNESS TRACKER — Google Sheets backend
 * =============================================
 * Paste this whole file into the Apps Script editor attached to your
 * spreadsheet (Extensions -> Apps Script), replacing whatever is there.
 *
 * *** YOU MUST REDEPLOY AFTER PASTING, NOT JUST SAVE ***
 * Saving the code alone does NOT update your live /exec URL. Go to:
 *   Deploy -> Manage deployments -> click the pencil (edit) icon on your
 *   existing deployment -> Version: "New version" -> Deploy.
 * This keeps your existing URL working:
 *   https://script.google.com/macros/s/AKfycbwgjKYnZFuObPAVq_Wwa8pkA5zxEMPX_PnHshcmkofGFx3Wg6xiCxT3nSVo1BC9aWQ3/exec
 * If you pick "Test deployments" or forget to select "New version", your
 * changes will NOT go live and the app will keep talking to the old code.
 *
 * Deployment settings required:
 *   Execute as:  Me
 *   Who has access:  Anyone
 * If access is "Anyone with Google account" instead, requests from the
 * hosted page will be redirected to a Google sign-in page and fail.
 *
 * WHY GET, NOT POST
 * Writes are sent as GET requests with the data in the URL (see
 * index.html). This sidesteps a well-known issue where browsers can't
 * reliably read the response of a cross-origin POST to an Apps Script Web
 * App — GET requests do not have this problem.
 */

const SHEETS = {
  weight: { name: 'PWA - Weight', headers: ['date', 'kg'] },
  food: { name: 'PWA - Food', headers: ['date', 'name', 'cal', 'pro', 'carb', 'fat'] },
  exercise: { name: 'PWA - Exercise', headers: ['date', 'type', 'duration', 'effort'] },
  measurements: {
    name: 'PWA - Measurements',
    headers: ['date', 'waist', 'hips', 'chest', 'weight', 'armL', 'armR', 'thighL', 'thighR'],
  },
};

// NOTE ON GET-ONLY DESIGN
// Writes are handled via doGet (action=add) rather than doPost. This is
// deliberate: Apps Script Web Apps deployed for "Anyone" access frequently
// fail to return readable cross-origin responses to POST requests from a
// browser (the CORS headers on the response after Google's internal
// redirect are inconsistent), while a plain GET request reliably works.
// Since our payloads are small (a few short fields), sending them as URL
// query parameters avoids the problem entirely. doPost is kept as a
// fallback/alias in case you ever call this from a server-side context
// that doesn't have the same browser CORS restrictions.

function doGet(e) {
  const params = (e && e.parameter) || {};
  const action = params.action || 'getAll';

  if (action === 'getAll') {
    return jsonResponse_({
      weights: readSheet_(SHEETS.weight),
      food: readSheet_(SHEETS.food),
      exercise: readSheet_(SHEETS.exercise),
      measurements: readSheet_(SHEETS.measurements),
    });
  }

  if (action === 'add') {
    try {
      const sheetKey = params.sheetKey;
      const row = JSON.parse(params.row || '{}');
      if (!SHEETS[sheetKey]) throw new Error('Unknown sheetKey: ' + sheetKey);
      addRow_(sheetKey, row);
      return jsonResponse_({ status: 'ok' });
    } catch (err) {
      return jsonResponse_({ status: 'error', message: String(err) });
    }
  }

  return jsonResponse_({ error: 'Unknown action: ' + action });
}

// Kept for completeness / non-browser callers. Browser code in index.html
// does not use this path (see note above).
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'add' && SHEETS[data.sheetKey]) {
      addRow_(data.sheetKey, data.row || {});
      return jsonResponse_({ status: 'ok' });
    }
    return jsonResponse_({ status: 'error', message: 'Unknown action or sheetKey' });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: String(err) });
  }
}

function addRow_(sheetKey, row) {
  const def = SHEETS[sheetKey];
  const sheet = getOrCreateSheet_(def.name, def.headers);
  const values = def.headers.map((h) => (row[h] === undefined ? '' : row[h]));
  sheet.appendRow(values);
}

function readSheet_(def) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(def.name);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, def.headers.length).getValues();
  const tz = Session.getScriptTimeZone();
  return values
    .map((r) => {
      const obj = {};
      def.headers.forEach((h, i) => {
        let v = r[i];
        if (v instanceof Date) v = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
        obj[h] = v;
      });
      return obj;
    })
    .filter((o) => o.date !== '' && o.date !== undefined);
}

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
