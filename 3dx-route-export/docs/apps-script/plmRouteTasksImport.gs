// PLM Route Tasks: import Route_Tasks_latest.csv from Google Drive into the Tasks tab.
//
// scripts/route_tasks.py saves the CSV in the Google Drive for desktop folder set as
// PLM_SHEET_DRIVE_DIR in .env (G:\My Drive\PLM Route Tasks). This script runs inside the sheet
// as you, so no web app, token or sign-in is needed.
//
// One-time setup, in the Route_Tasks_2026 spreadsheet:
//   1. Extensions > Apps Script. Replace all the code with this file. Save.
//   2. Pick "setup" in the function list at the top, click Run, and allow access when asked
//      (the list shows only setup, importNow, importIfNew and onOpen; the rest are helpers).
//      That imports the CSV once and adds an hourly check.
// After that: the sheet updates within an hour of each run, or right away with
// menu PLM > Import latest (reload the sheet once to see the menu).

var CSV_NAME = 'Route_Tasks_latest.csv';
var TAB_GID = 366646702;  // the Tasks tab (gid= in the sheet link)
var HIDDEN = ['Status', 'Priority', 'Due', 'Time Taken', 'Late'];
var MERGED_COLS = 4;  // Row, Route Title, Route Created, Route Completed: one value per route

function onOpen() {
  SpreadsheetApp.getUi().createMenu('PLM').addItem('Import latest', 'importNow').addToUi();
}

function setup() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'importIfNew') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importIfNew').timeBased().everyHours(1).create();
  importNow();
}

function importNow() {
  var msg = importCsv_(true);
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'PLM import', 8); } catch (e) {}
  Logger.log(msg);
}

function importIfNew() {
  Logger.log(importCsv_(false));
}

function importCsv_(force) {
  var files = DriveApp.getFilesByName(CSV_NAME);
  var file = null;
  while (files.hasNext()) {
    var f = files.next();
    if (!f.isTrashed() && (!file || f.getLastUpdated() > file.getLastUpdated())) file = f;
  }
  if (!file) return CSV_NAME + ' not found in Google Drive';

  var props = PropertiesService.getScriptProperties();
  var stamp = String(file.getLastUpdated().getTime());
  if (!force && props.getProperty('imported') === stamp) return 'no new file';

  var values = Utilities.parseCsv(file.getBlob().getDataAsString('UTF-8'));
  if (!values.length) return CSV_NAME + ' is empty';
  var headers = values[0];
  var cols = headers.length;
  var rows = values.length;

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(function (s) {
    return s.getSheetId() === TAB_GID;
  })[0];
  if (!sheet) return 'tab with gid ' + TAB_GID + ' not found';

  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  sheet.clear();
  sheet.showColumns(1, sheet.getMaxColumns());
  if (sheet.getMaxColumns() < cols) sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
  if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  sheet.getRange(1, 1, rows, cols).setNumberFormat('@').setValues(values);
  sheet.getRange(1, 1, 1, cols).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1f4e79');
  sheet.setFrozenRows(1);
  HIDDEN.forEach(function (name) {
    var i = headers.indexOf(name);
    if (i >= 0) sheet.hideColumns(i + 1);
  });
  if (sheet.getMaxRows() > rows) sheet.deleteRows(rows + 1, sheet.getMaxRows() - rows);
  mergePerRoute_(sheet, values, rows);
  sheet.getRange(1, 1, rows, cols).createFilter();
  var note = 'Imported ' + Utilities.formatDate(file.getLastUpdated(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm')
    + ' - ' + (rows - 1) + ' tasks';
  sheet.getRange(1, 1).setNote(note);
  props.setProperty('imported', stamp);
  return note;
}

// Row, Route Title, Route Created and Route Completed shown once per route, like the Excel file:
// each is merged down the rows of one route (rows with the same Row number).
function mergePerRoute_(sheet, values, rows) {
  var start = 1;  // index in values of the route's first task (0 = header)
  for (var r = 2; r <= rows; r++) {
    if (r < rows && values[r][0] === values[start][0]) continue;
    if (r - start > 1) sheet.getRange(start + 1, 1, r - start, MERGED_COLS).mergeVertically();
    start = r;
  }
  sheet.getRange(2, 1, rows - 1, MERGED_COLS).setVerticalAlignment('middle');
}
