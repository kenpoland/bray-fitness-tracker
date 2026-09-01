/**
 * BRAY FITNESS TRACKER — Google Sheets backend
 * =============================================
 * Paste this whole file into the Apps Script editor attached to your
 * spreadsheet (Extensions -> Apps Script), replacing whatever is there,
 * then click Deploy -> Manage deployments -> edit (pencil) icon on your
 * existing deployment -> New version -> Deploy.
 * This keeps your existing /exec URL working:
 *   https://script.google.com/macros/s/AKfycbzXkABridi2HySg4LvbjxyiMfgqM7iIYOpHhjucMJquK64H9EXGWM33tJeoKA6hczCE/exec
 *
 * Deployment settings required (Deploy -> New deployment if setting up fresh):
 *   Execute as:  Me
 *   Who has access:  Anyone
 * (If access isn't "Anyone", the app running on GitHub Pages won't be able
 *  to reach it, since it's not you being logged in when the page loads.)
 *
 * WHAT THIS DOES
 * The app writes to four dedicated tabs it creates automatically the first
 * time you use it — "PWA - Weight", "PWA - Food", "PWA - Exercise",
 * "PWA - Measurements" — so it never touches or reformats your existing
 * Master Tracker tabs. You can reference these new tabs from your
 * dashboards with normal formulas (e.g. IMPORTRANGE-style SUMIFS) if you
 * want them to feed into the Weekly/Monthly Dashboard later.
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

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'getAll';
  if (action === 'getAll') {
    return jsonResponse_({
      weights: readSheet_(SHEETS.weight),
      food: readSheet_(SHEETS.food),
      exercise: readSheet_(SHEETS.exercise),
      measurements: readSheet_(SHEETS.measurements),
    });
  }
  return jsonResponse_({ error: 'Unknown action: ' + action });
}

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
