/**
 * remove_spacers.gs — Standalone maintenance script
 *
 * Removes all 2 px spacer rows inserted by buildTRATool().
 * Paste this into its own Apps Script project bound to the spreadsheet,
 * or add it to an existing project alongside build_tra_tool.gs.
 *
 * Functions:
 *   removeSpacerRows()             — all sheets
 *   removeSpacerRowsOnSheet()      — active sheet only
 *   listSpacerRows()               — audit without deleting (logs to console)
 */

var SPACER_HEIGHT_PX = 2;  // rows at or below this height are treated as spacers

/**
 * Remove spacer rows from every sheet in the spreadsheet.
 * Works bottom-up to avoid row-index shift errors.
 */
function removeSpacerRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var totalRemoved = 0;

  sheets.forEach(function(sh) {
    var removed = _removeSpacersFromSheet(sh);
    if (removed > 0) {
      Logger.log(sh.getName() + ': removed ' + removed + ' spacer rows');
    }
    totalRemoved += removed;
  });

  SpreadsheetApp.getUi().alert(
    'Done — removed ' + totalRemoved + ' spacer row(s) across all sheets.\n' +
    'Check the Apps Script log for a per-sheet breakdown.'
  );
}

/**
 * Remove spacer rows from the active sheet only.
 */
function removeSpacerRowsOnSheet() {
  var sh = SpreadsheetApp.getActiveSheet();
  var removed = _removeSpacersFromSheet(sh);
  SpreadsheetApp.getUi().alert(
    'Done — removed ' + removed + ' spacer row(s) from "' + sh.getName() + '".'
  );
}

/**
 * Log spacer row positions without deleting anything.
 * Useful to audit before committing to a deletion.
 */
function listSpacerRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var report = [];

  sheets.forEach(function(sh) {
    var lastRow = sh.getLastRow();
    for (var i = 1; i <= lastRow; i++) {
      if (sh.getRowHeight(i) <= SPACER_HEIGHT_PX) {
        report.push(sh.getName() + ' row ' + i + ' (' + sh.getRowHeight(i) + 'px)');
      }
    }
  });

  if (report.length === 0) {
    Logger.log('No spacer rows found.');
  } else {
    Logger.log('Spacer rows found:\n' + report.join('\n'));
  }

  SpreadsheetApp.getUi().alert(
    report.length + ' spacer row(s) found.\nSee Apps Script log for details.'
  );
}

// ── Internal helper ────────────────────────────────────────────────────────────

function _removeSpacersFromSheet(sh) {
  var lastRow = sh.getLastRow();
  var toDelete = [];

  for (var i = 1; i <= lastRow; i++) {
    if (sh.getRowHeight(i) <= SPACER_HEIGHT_PX) {
      toDelete.push(i);
    }
  }

  // Delete bottom-up to keep indices stable
  for (var j = toDelete.length - 1; j >= 0; j--) {
    sh.deleteRow(toDelete[j]);
  }

  return toDelete.length;
}
