// PLM Route Tasks -> Google Sheet (web app bound to the target spreadsheet).
//
// One-time setup, in the spreadsheet: Extensions > Apps Script, paste this file, set TOKEN to the
// value of PLM_SHEET_TOKEN in the project's .env, then Deploy > New deployment > Web app,
// execute as "Me", access "Anyone". Put the /exec URL in .env as PLM_SHEET_WEB_APP_URL.
//
// scripts/sheet_push.py posts the rows in parts:
//   { token, gid, headers, rows, start, last, hidden, updated }
// start = 0 clears the tab and writes the header; each part is written below the previous one;
// last = true hides the columns named in `hidden` and notes the run time on the header.
var TOKEN = 'PASTE_PLM_SHEET_TOKEN_HERE';

function doPost(e) {
  var result = { ok: true };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var p = JSON.parse(e.postData.contents);
    if (!TOKEN || TOKEN.indexOf('PASTE_') === 0 || p.token !== TOKEN) throw new Error('token does not match');

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(function (s) {
      return String(s.getSheetId()) === String(p.gid);
    })[0];
    if (!sheet) throw new Error('tab with gid ' + p.gid + ' not found');

    var cols = p.headers.length;
    if (p.start === 0) {
      if (sheet.getFilter()) sheet.getFilter().remove();
      sheet.clear();
      sheet.showColumns(1, sheet.getMaxColumns());
      if (sheet.getMaxColumns() < cols) sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
      sheet.getRange(1, 1, 1, cols).setValues([p.headers])
        .setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f4e79');
      sheet.setFrozenRows(1);
    }

    var first = 2 + p.start;
    var need = first + p.rows.length - 1;
    if (sheet.getMaxRows() < need) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows());
    if (p.rows.length) sheet.getRange(first, 1, p.rows.length, cols).setNumberFormat('@').setValues(p.rows);

    if (p.last) {
      (p.hidden || []).forEach(function (name) {
        var i = p.headers.indexOf(name);
        if (i >= 0) sheet.hideColumns(i + 1);
      });
      var total = p.start + p.rows.length;
      if (sheet.getMaxRows() > total + 1) sheet.deleteRows(total + 2, sheet.getMaxRows() - total - 1);
      if (total) sheet.getRange(1, 1, total + 1, cols).createFilter();
      sheet.getRange(1, 1).setNote('Updated ' + p.updated + ' - ' + total + ' tasks');
    }
    result.written = p.rows.length;
  } catch (err) {
    result = { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
