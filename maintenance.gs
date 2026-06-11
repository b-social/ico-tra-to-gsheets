/**
 * ICO TRA Tool — Maintenance Script
 * Paste into Apps Script editor (separate from build_tra_tool.gs).
 *
 * Functions:
 *   removeSpacerRows()            — deletes all 2px spacer rows across all working sheets
 *   removeSpacerRowsActiveSheet() — same, active sheet only
 *   auditTRATool()                — diffs live sheet against expected structure, logs findings
 */

// ─── Expected structure ───────────────────────────────────────────────────────

var EXPECTED_TABS = [
  '📋 Instructions',
  'Q1 · Transfer Details',
  'Q2 · PI Risk Scores',
  'Q3 · Investigation Level',
  'Q4 · Human Rights Risk',
  'Q5 · Enforcement',
  'Q6 · Exceptions',
  '✅ Summary',
  '🏦 Kroo PI Categories',
  'Lookups',
];

var EXPECTED_NAMED_RANGES = [
  'PI_NAMES',
  'PI_SCORES',
  'MAX_SCORE',
  'BIZ_SIZE',
  'XFER_VOL',
  'DP_B',
  'KQ4_1',
  'KQ4_2',
  'EQ5_1',
  'EQ5_2',
  'EQ5_3',
  'EQ5_4',
  'DP_E',
  'DP_F_VAL',
  'DD_SCORES',
  'DD_IMPORTER_STATUS',
  'DD_ORG_TYPE',
  'DD_VULN',
  'DD_FREQUENCY',
  'DD_BIZ_SIZE',
  'DD_XFER_VOL',
  'DD_INV_LEVEL',
  'DD_KQ4_1',
  'DD_KQ4_2',
  'DD_EQ5_1',
  'DD_EQ5_2',
  'DD_EQ5_3',
  'DD_EQ5_4',
  'DD_DP_E',
  'DD_EXCEPTION_YN',
  'DD_BENEFIT_YN',
  'DD_DP_F',
];

// Key formula fragments expected in specific sheets: [sheetName, substring]
var EXPECTED_FORMULAS = [
  ['Q2 · PI Risk Scores',      'MAX('],
  ['Q3 · Investigation Level', 'MAX_SCORE'],
  ['Q3 · Investigation Level', 'BIZ_SIZE'],
  ['Q4 · Human Rights Risk',   'KQ4_1'],
  ['Q4 · Human Rights Risk',   'KQ4_2'],
  ['Q5 · Enforcement',         'EQ5_1'],
  ['Q5 · Enforcement',         'DP_E'],
  ['Q6 · Exceptions',          'DP_F_VAL'],
  ['✅ Summary',                'MAX_SCORE'],
  ['✅ Summary',                'DP_E'],
];

// ─── Audit ────────────────────────────────────────────────────────────────────

// eslint-disable-next-line no-unused-vars
function auditTRATool() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var issues = [];
  var ok = [];

  // 1. Tab existence
  var sheetNames = ss.getSheets().map(function(sh) { return sh.getName(); });
  EXPECTED_TABS.forEach(function(name) {
    if (sheetNames.indexOf(name) === -1) {
      issues.push('❌ Missing tab: "' + name + '"');
    } else {
      ok.push('✅ Tab present: "' + name + '"');
    }
  });

  // 2. Named ranges
  var namedRanges = ss.getNamedRanges().map(function(nr) { return nr.getName(); });
  EXPECTED_NAMED_RANGES.forEach(function(name) {
    if (namedRanges.indexOf(name) === -1) {
      issues.push('❌ Missing named range: ' + name);
    } else {
      ok.push('✅ Named range: ' + name);
    }
  });

  // 3. Named range targets on Lookups tab
  var lookupSheet = ss.getSheetByName('Lookups');
  if (lookupSheet) {
    ss.getNamedRanges().forEach(function(nr) {
      var range = nr.getRange();
      var shName = range.getSheet().getName();
      var ddNames = EXPECTED_NAMED_RANGES.filter(function(n) { return n.indexOf('DD_') === 0; });
      if (ddNames.indexOf(nr.getName()) !== -1 && shName !== 'Lookups') {
        issues.push('⚠️ Named range ' + nr.getName() + ' points to "' + shName + '" — expected "Lookups"');
      }
    });
  }

  // 4. Formula spot-checks — scan all cells in each sheet for formula substring
  EXPECTED_FORMULAS.forEach(function(pair) {
    var shName = pair[0];
    var fragment = pair[1];
    var sh = ss.getSheetByName(shName);
    if (!sh) return;
    var found = false;
    var data = sh.getDataRange().getFormulas();
    for (var r = 0; r < data.length && !found; r++) {
      for (var c = 0; c < data[r].length && !found; c++) {
        if (data[r][c].indexOf(fragment) !== -1) found = true;
      }
    }
    if (!found) {
      issues.push('⚠️ "' + shName + '" — no formula containing "' + fragment + '"');
    } else {
      ok.push('✅ Formula "' + fragment + '" found in "' + shName + '"');
    }
  });

  // 5. Spacer row count (informational)
  var spacerCounts = {};
  ss.getSheets().forEach(function(sh) {
    if (sh.getName() === 'Lookups') return;
    var count = 0;
    for (var r = 1; r <= sh.getLastRow(); r++) {
      if (sh.getRowHeight(r) <= 2) count++;
    }
    if (count > 0) spacerCounts[sh.getName()] = count;
  });
  if (Object.keys(spacerCounts).length > 0) {
    Object.keys(spacerCounts).forEach(function(name) {
      issues.push('ℹ️ Spacer rows remaining in "' + name + '": ' + spacerCounts[name] + ' (run removeSpacerRows to clean up)');
    });
  }

  // 6. Output
  var report = [];
  if (issues.length > 0) {
    report.push('── ISSUES (' + issues.length + ') ──');
    report = report.concat(issues);
    report.push('');
  }
  report.push('── OK (' + ok.length + '/' + (ok.length + issues.filter(function(i) { return i.indexOf('❌') === 0; }).length) + ') ──');
  report = report.concat(ok.slice(0, 10));
  if (ok.length > 10) report.push('  … and ' + (ok.length - 10) + ' more');

  Logger.log(report.join('\n'));
  SpreadsheetApp.getUi().alert(
    'TRA Audit complete.\n\n' +
    issues.length + ' issue(s) found, ' + ok.length + ' checks passed.\n\n' +
    (issues.length > 0 ? issues.join('\n') : '✅ All checks passed.')
  );
}

// ─── Spacer removal ───────────────────────────────────────────────────────────

// eslint-disable-next-line no-unused-vars
function removeSpacerRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets().filter(function(sh) {
    return sh.getName() !== 'Lookups';
  });

  var totalDeleted = 0;
  sheets.forEach(function(sh) {
    var toDelete = [];
    for (var r = sh.getLastRow(); r >= 1; r--) {
      if (sh.getRowHeight(r) <= 2) toDelete.push(r);
    }
    toDelete.forEach(function(r) { sh.deleteRow(r); });
    totalDeleted += toDelete.length;
  });

  SpreadsheetApp.getUi().alert(
    '✅ Removed ' + totalDeleted + ' spacer row(s) across ' + sheets.length + ' sheet(s).'
  );
}

// eslint-disable-next-line no-unused-vars
function removeSpacerRowsActiveSheet() {
  var sh = SpreadsheetApp.getActiveSheet();
  var toDelete = [];
  for (var r = sh.getLastRow(); r >= 1; r--) {
    if (sh.getRowHeight(r) <= 2) toDelete.push(r);
  }
  toDelete.forEach(function(r) { sh.deleteRow(r); });
  SpreadsheetApp.getUi().alert(
    '✅ Removed ' + toDelete.length + ' spacer row(s) from "' + sh.getName() + '".'
  );
}
