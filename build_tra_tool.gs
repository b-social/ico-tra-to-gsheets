/**
 * ICO Transfer Risk Assessment (TRA) Tool
 * Google Apps Script — Extensions > Apps Script > run buildTRATool()
 *
 * Constraints:
 *   • No merged cells
 *   • Risk scores 1–5 (1 Very Low → 5 Very High)
 *   • Dracula colour theme
 *   • All dropdown lists sourced from named ranges on the Lookups tab
 *   • Column widths capped at 200px
 *   • Decision Point thresholds: ≤2 proceed · 3 moderate · ≥4 high
 *   • Decision Point A driven by MAX score across PI categories
 *
 * Kroo-specific data tab:
 *   • PCI-DSS card data, KYC/CDD/EDD, AML/CTF/SARs, financial transactions,
 *     account data, auth/authz, Open Banking/PSD2, fraud/disputes,
 *     device/technical, behavioural analytics, customer support,
 *     regulatory/compliance, credit/lending
 *   • ⚠️ Transaction data inference risk: payments to trade unions, charities,
 *     political parties, religious organisations, medical providers etc. can
 *     reveal GDPR Art 9 special category data by inference. Flagged per row.
 */

// ─── Dracula palette ──────────────────────────────────────────────────────────
var BG  = '#282a36', SEL = '#44475a', FG  = '#f8f8f2', CMT = '#6272a4';
var CYN = '#8be9fd', GRN = '#50fa7b', ORG = '#ffb86c', PNK = '#ff79c6';
var PRP = '#bd93f9', RED = '#ff5555', YLW = '#f1fa8c', BLK = '#21222c';

// Score 1–5: background / foreground / label
var SBG = ['', '#50fa7b', '#8be9fd', '#f1fa8c', '#ffb86c', '#ff5555'];
var SFG = ['', '#282a36', '#282a36', '#282a36', '#282a36', '#f8f8f2'];
var SLB = ['', '1 – Very Low', '2 – Low', '3 – Moderate', '4 – High', '5 – Very High'];

// Row offsets are tracked per-sheet so we don't hard-code row numbers
// Named ranges (spreadsheet-scoped) used for cross-sheet formulas:
//   MAX_SCORE, BIZ_SIZE, XFER_VOL, DP_B,
//   KQ4_1, KQ4_2, EQ5_1, EQ5_2, EQ5_3, EQ5_4, DP_E

// ─── Entry point ──────────────────────────────────────────────────────────────
// eslint-disable-next-line no-unused-vars
function buildTRATool() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var TABS = [
    ['📋 Instructions',          PRP],
    ['Q1 · Transfer Details',    CYN],
    ['Q2 · PI Risk Scores',      ORG],
    ['Q3 · Investigation Level', YLW],
    ['Q4 · Human Rights Risk',   PNK],
    ['Q5 · Enforcement',         GRN],
    ['Q6 · Exceptions',          RED],
    ['✅ Summary',                FG],
    ['🏦 Kroo PI Categories',    GRN],
    ['Lookups',                  CMT],
  ];

  // Create or clear sheets
  var sheets = {};
  var existing = ss.getSheets();

  TABS.forEach(function(tab, i) {
    var name  = tab[0];
    var color = tab[1];
    var sh = ss.getSheetByName(name);
    if (!sh) {
      if (i === 0 && existing.length === 1 && existing[0].getName() === 'Sheet1') {
        existing[0].setName(name);
        sh = existing[0];
      } else {
        sh = ss.insertSheet(name);
      }
    }
    sh.setTabColor(color);
    sh.clearContents();
    sh.clearFormats();
    sh.clearConditionalFormatRules();
    sh.clearNotes();
    sheets[name] = sh;
  });

  // Lookups tab first — all other sheets reference it
  buildLookups(sheets['Lookups'], ss);
  appendKrooToLookups(sheets['Lookups'], ss);   // extends PI_NAMES / PI_SCORES

  // Working sheets
  buildInstructions(sheets['📋 Instructions']);
  buildQ1(sheets['Q1 · Transfer Details']);
  buildQ2(sheets['Q2 · PI Risk Scores'], ss);
  buildQ3(sheets['Q3 · Investigation Level'], ss);
  buildQ4(sheets['Q4 · Human Rights Risk'], ss);
  buildQ5(sheets['Q5 · Enforcement'], ss);
  buildQ6(sheets['Q6 · Exceptions'], ss);
  buildSummary(sheets['✅ Summary']);
  buildKrooPICategories(sheets['🏦 Kroo PI Categories']);

  ss.setActiveSheet(sheets['📋 Instructions']);

  SpreadsheetApp.getUi().alert(
    '✅ ICO TRA Tool built!\n\n' +
    'Start at 📋 Instructions, work through Q1–Q6.\n' +
    '✅ Summary tracks your outcome automatically.\n\n' +
    'All dropdown option lists live on the Lookups tab.'
  );
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function cw(sh, arr) {
  arr.forEach(function(pair) { sh.setColumnWidth(pair[0], pair[1]); });
}

function baseStyle(sh, rows, cols) {
  sh.getRange(1, 1, rows, cols)
    .setBackground(BG).setFontColor(FG).setFontSize(10).setWrap(true);
}

// H1 – title bar
function h1(sh, r, c, txt) {
  sh.getRange(r, c).setValue(txt)
    .setBackground(PRP).setFontColor(BG)
    .setFontSize(12).setFontWeight('bold').setWrap(true);
  sh.setRowHeight(r, 36);
}

// H2 – section heading
function h2(sh, r, c, txt) {
  sh.getRange(r, c).setValue(txt)
    .setBackground(SEL).setFontColor(CYN)
    .setFontSize(10).setFontWeight('bold').setWrap(true);
  sh.setRowHeight(r, 22);
}

// H3 – sub-heading
function h3(sh, r, c, txt) {
  sh.getRange(r, c).setValue(txt)
    .setBackground(BG).setFontColor(CMT)
    .setFontSize(9).setFontWeight('bold').setWrap(true);
}

// Label (italic, muted)
function lbl(sh, r, c, txt) {
  sh.getRange(r, c).setValue(txt)
    .setBackground(BG).setFontColor(CMT)
    .setFontSize(9).setFontStyle('italic').setWrap(true);
}

// Input cell (dark, bordered)
function inp(sh, r, c) {
  sh.getRange(r, c)
    .setBackground(SEL).setFontColor(FG).setFontSize(10).setWrap(true)
    .setBorder(true, true, true, true, false, false,
               CMT, SpreadsheetApp.BorderStyle.SOLID);
}

// Auto-calculated cell (very dark bg, cyan text)
function auto(sh, r, c, formula) {
  sh.getRange(r, c).setFormula(formula)
    .setBackground(BLK).setFontColor(CYN)
    .setFontSize(10).setWrap(true);
}

// Decision point result cell (prominent, auto-coloured)
function dpCell(sh, r, c, formula) {
  var cell = sh.getRange(r, c);
  cell.setFormula(formula)
      .setFontSize(11).setFontWeight('bold').setWrap(true)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, false, false,
                 PRP, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(r, 65);
  cfProceed(cell);
}

// Dropdown backed by a named range on the Lookups tab
function ddNamed(sh, r, c, namedRangeName) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      SpreadsheetApp.getActiveSpreadsheet().getRangeByName(namedRangeName), true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(r, c)
    .setDataValidation(rule)
    .setBackground(SEL).setFontColor(FG).setFontSize(10).setWrap(true);
}

// Checkbox
function ck(sh, r, c) {
  sh.getRange(r, c)
    .setDataValidation(
      SpreadsheetApp.newDataValidation().requireCheckbox().build())
    .setBackground(SEL);
}

// Set / overwrite named range — find-and-update to avoid removeNamedRange exceptions
function nr(ss, name, range) {
  var existing = ss.getNamedRanges();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].getName() === name) {
      existing[i].setRange(range);
      return;
    }
  }
  ss.setNamedRange(name, range);
}

// Tiny spacer row
function spacer(sh, r) {
  sh.getRange(r, 1).setBackground(BG);
  sh.setRowHeight(r, 6);
}

// ─── Conditional formatting ───────────────────────────────────────────────────

// Score 1–5 conditional format (works on text "1"-"5" or numbers)
function cfScore(range) {
  var sh    = range.getSheet();
  var rules = sh.getConditionalFormatRules();
  for (let s = 1; s <= 5; s++) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(String(s))
        .setBackground(SBG[s]).setFontColor(SFG[s])
        .setRanges([range]).build()
    );
  }
  sh.setConditionalFormatRules(rules);
}

// Proceed / Cannot / Conditional / pending colouring
function cfProceed(range) {
  var sh    = range.getSheet();
  var rules = sh.getConditionalFormatRules();
  var add   = function(text, bg, fg) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(text).setBackground(bg).setFontColor(fg)
        .setRanges([range]).build()
    );
  };
  add('MAY PROCEED',   GRN, BG);
  add('CANNOT PROCEED', RED, FG);
  add('CONDITIONAL',   ORG, BG);
  add('⏳',            SEL, YLW);
  sh.setConditionalFormatRules(rules);
}

// ─── LOOKUPS TAB ──────────────────────────────────────────────────────────────
// Single source of truth for all dropdown option lists and PI category scores.
// Named ranges are created here and used via ddNamed() on all other sheets.

function buildLookups(sh, ss) {
  cw(sh, [[1,220],[2,80],[3,180],[4,160],[5,160]]);
  baseStyle(sh, 200, 5);
  sh.setFrozenRows(1);

  h1(sh, 1, 1, 'Lookups — all dropdown lists and reference data (do not edit column A–B values)');
  for (let c = 2; c <= 5; c++) sh.getRange(1, c).setBackground(PRP);

  // ── Helper to write a list and create a named range ────────────────────────
  function list(startRow, col, name, values, heading) {
    sh.getRange(startRow - 1, col).setValue(heading)
      .setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);
    values.forEach(function(v, i) {
      sh.getRange(startRow + i, col).setValue(v)
        .setBackground(BG).setFontColor(FG).setFontSize(10);
    });
    nr(ss, name, sh.getRange(startRow, col, values.length, 1));
  }

  // ── Column 1: PI Category names + ICO initial scores ──────────────────────
  // Used by VLOOKUP in Q2 col C (score) and for the PI dropdown in Q2 col B.
  // Format: col A = name, col B = initial score (1–5)
  sh.getRange(2, 1).setValue('PI Category')
    .setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);
  sh.getRange(2, 2).setValue('ICO Score')
    .setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);
  sh.getRange(2, 3).setValue('Special Cat?')
    .setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);

  var piData = [
    // [name, ICO score (1-5), special category (true/false)]
    ['Name',                                                     2, false],
    ['Address & contact details',                                2, false],
    ['Age / date of birth',                                      2, false],
    ['Gender',                                                   4, true],
    ['Biometric data',                                           4, true],
    ['CCTV, photos and images (not biometric)',                  3, false],
    ['Race / ethnic origin',                                     4, true],
    ['ID documentation (passport, driving licence, NI number)',  4, false],
    ['Medical records',                                          4, true],
    ['Medication records',                                       4, true],
    ['Name and contact details of GP',                           2, false],
    ['Name and contact details of specialist medical staff',     4, true],
    ['Genetic data',                                             4, true],
    ['Current marriage / partnerships',                          3, false],
    ['Marital history',                                          3, false],
    ['Details of family / household members',                    3, false],
    ['Habits',                                                   2, false],
    ['Housing',                                                  2, false],
    ['Travel details',                                           2, false],
    ['Leisure activities',                                       2, false],
    ['Location data',                                            4, false],
    ['Membership of charitable / voluntary organisations',       2, false],
    ['Political opinions',                                       4, true],
    ['Religious or philosophical beliefs',                       4, true],
    ['Trade union membership',                                   4, true],
    ['Sex life or sexual orientation',                           4, true],
    ['Free text about an individual (email, social media, chat)',4, true],
    ['Employment and career history',                            2, false],
    ['Recruitment records',                                      2, false],
    ['Termination details',                                      4, false],
    ['Attendance records',                                       2, false],
    ['Health and safety records',                                4, true],
    ['Performance appraisals',                                   3, false],
    ['Training records',                                         2, false],
    ['Security records',                                         2, false],
    ['Financial account / credit card details',                  4, false],
    ['Income',                                                   3, false],
    ['Salary',                                                   3, false],
    ['Assets and investments',                                   3, false],
    ['Payments',                                                 3, false],
    ['Creditworthiness / credit score',                          3, false],
    ['Loans',                                                    3, false],
    ['Benefits',                                                 3, false],
    ['Grants',                                                   3, false],
    ['Insurance details',                                        3, false],
    ['Pension details',                                          3, false],
    ['Goods or services supplied',                               2, false],
    ['Marketing preferences',                                    2, false],
    ['Delivery preferences',                                     2, false],
    ['Licences issued',                                          2, false],
    ['Records of unspent criminal convictions / offences',       4, false],
    ['Records of spent criminal convictions / offences',         4, false],
    ['Records of DBS checks',                                    4, false],
    ['Criminal investigation records',                           4, false],
  ];

  var PI_DATA_START = 3;
  piData.forEach(function(row, i) {
    var r = PI_DATA_START + i;
    sh.getRange(r, 1).setValue(row[0])
      .setBackground(BG).setFontColor(FG).setFontSize(10);
    sh.getRange(r, 2).setValue(row[1])
      .setBackground(SBG[row[1]]).setFontColor(SFG[row[1]])
      .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(r, 3).setValue(row[2] ? 'Yes' : 'No')
      .setBackground(row[2] ? PNK : BG)
      .setFontColor(row[2] ? BG : CMT)
      .setFontSize(9).setHorizontalAlignment('center');
  });

  // Named ranges for VLOOKUP source (col A:B and col A for name dropdown)
  nr(ss, 'PI_SCORES',  sh.getRange(PI_DATA_START, 1, piData.length, 2));
  nr(ss, 'PI_NAMES',   sh.getRange(PI_DATA_START, 1, piData.length, 1));

  // ── Column 4: all other dropdown lists ────────────────────────────────────
  var COL = 4;
  var row = 2;

  // Each call writes a heading + values, creates named range, returns next row
  function L(name, heading, values) {
    list(row + 1, COL, name, values, heading);
    row += values.length + 2;
  }

  L('DD_SCORES', 'Score (1–5)', ['1','2','3','4','5']);
  L('DD_IMPORTER_STATUS', 'Importer status',
    ['Controller','Processor','Sub-processor','Joint controller']);
  L('DD_ORG_TYPE', 'Organisation type',
    ['Commercial — standard',
     'Commercial — multinational group',
     'Commercial — large (not multinational)',
     'Commercial — small / sole trader',
     'Public sector',
     'Not for profit',
     'Regulated — financial services',
     'Regulated — legal services',
     'Regulated — healthcare',
     'Regulated — other',
     'Other']);
  L('DD_VULN', 'Vulnerability',
    ['Adults only (not vulnerable)',
     'Children or vulnerable adults only',
     'Both adults and children / vulnerable adults']);
  L('DD_FREQUENCY', 'Transfer frequency',
    ['Once only',
     'Recurring — specify interval in notes',
     'Continuous — specify period in notes']);
  L('DD_BIZ_SIZE', 'Organisation size',
    ['SME (Tier 1 or Tier 2 data protection fee payer)',
     'Large business']);
  L('DD_XFER_VOL', 'Transfer volume',
    ['Low volume',
     'High volume (significant amount, one-off or recurring)']);
  L('DD_INV_LEVEL', 'Investigation level',
    ['Level 1 investigation',
     'Level 2 investigation',
     'Level 3 investigation — Option (i)',
     'Level 3 investigation — Option (ii) [all high-score PI = significant risk data]']);
  L('DD_KQ4_1', 'Q4 Key Question 1',
    ['No concerns — proceed to Decision Point C → tick C1',
     'Yes, we have concerns — continue to Key Question 2']);
  L('DD_KQ4_2', 'Q4 Key Question 2',
    ['N/A — no concerns raised at Key Question 1',
     'No — transfer does NOT significantly increase the risk → tick C2',
     'Yes — ALL categories of PI increase the risk → tick C3',
     'Yes — SOME categories of PI increase the risk → tick C4']);
  L('DD_EQ5_1', 'EQ5 Question 1',
    ['Yes — low / moderate only → tick D1',
     'No — includes high-score data → continue to EQ2']);
  L('DD_EQ5_2', 'EQ5 Question 2',
    ['N/A — only low/moderate data (EQ1 = Yes)',
     'No concerns → tick D2',
     'Yes / not sure → add notes and continue to EQ3']);
  L('DD_EQ5_3', 'EQ5 Question 3',
    ['N/A — only low/moderate data',
     'N/A — no concerns at EQ2',
     'Yes — high likelihood importer will accept UK Court / arbitration → tick D3',
     'No — not satisfied → continue to EQ4']);
  L('DD_EQ5_4', 'EQ5 Question 4',
    ['N/A',
     'Yes — other factors apply → note them → tick D4',
     'No → tick D5']);
  L('DD_DP_E', 'Decision Point E',
    ['E1 — No significant risk data identified. May proceed.',
     'E2 — All high-score data is both human rights AND enforceability risk data (Level 3 Option ii)',
     'E3 — All categories are human rights risk data (ticked C3)',
     'E4 — Some categories are human rights risk data (ticked C4)',
     'E5 — All high-score data is enforceability risk data (ticked D5)']);
  L('DD_EXCEPTION_YN', 'Exception applies?',
    ['Yes','No','N/A']);
  L('DD_BENEFIT_YN', 'Benefit outweighs risk?',
    ['Yes — benefits outweigh risks',
     'No — risks outweigh benefits',
     'N/A']);
  L('DD_DP_F', 'Decision Point F',
    ['F1 — One or more exceptions apply to ALL significant risk data. May proceed.',
     'F2 — Exceptions do NOT apply to all significant risk data. May NOT proceed.']);

  // Score key legend in column 5
  sh.getRange(2, 5).setValue('Score key')
    .setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);
  for (let s = 1; s <= 5; s++) {
    sh.getRange(2 + s, 5).setValue(SLB[s])
      .setBackground(SBG[s]).setFontColor(SFG[s])
      .setFontSize(9).setFontWeight('bold');
  }
}

// ─── INSTRUCTIONS ─────────────────────────────────────────────────────────────

function buildInstructions(sh) {
  cw(sh, [[1,180],[2,200],[3,180],[4,120]]);
  baseStyle(sh, 80, 4);
  sh.setFrozenRows(2);

  h1(sh, 1, 1, '🔒  ICO Transfer Risk Assessment (TRA) Tool');
  for (let c = 2; c <= 4; c++) sh.getRange(1, c).setBackground(PRP);

  sh.getRange(2, 1).setValue('UK GDPR Article 46 · Based on ICO TRA Tool (November 2022)')
    .setBackground(SEL).setFontColor(CMT).setFontSize(9).setFontStyle('italic');
  for (let c = 2; c <= 4; c++) sh.getRange(2, c).setBackground(SEL);

  var r = 3;

  h2(sh, r++, 1, 'About this tool');
  [
    'Helps you carry out and record a TRA for restricted international transfers of personal information under UK GDPR.',
    'You do NOT have to use this template, but you must record your TRA. This sheet follows the ICO TRA Tool structure exactly.',
    'Designed for a straightforward transfer: one importer, one destination country. Adapt for more complex flows.',
    'Seek professional data protection advice if needed.',
  ].forEach(function(t) {
    sh.getRange(r, 1).setValue(t).setBackground(BG).setFontColor(FG).setFontSize(10).setWrap(true);
    for (let c = 2; c <= 4; c++) sh.getRange(r, c).setBackground(BG);
    sh.setRowHeight(r++, 30);
  });

  spacer(sh, r++);
  h2(sh, r, 1, 'How to use');
  sh.getRange(r, 2).setValue('Purpose').setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);
  sh.getRange(r, 3).setValue('Key output').setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);
  sh.getRange(r, 4).setBackground(SEL);
  r++;

  var nav = [
    ['Q1 · Transfer Details',      'Document importer, people, duration, protections, PI categories',    'Complete record of transfer circumstances'],
    ['Q2 · PI Risk Scores',        'Score each PI category 1–5 (ICO initial score auto-looked-up)',       'Decision Point A → proceed or go to Q3'],
    ['Q3 · Investigation Level',   'Matrix selects proportionate investigation level',                    'Decision Point B → level 1 / 2 / 3'],
    ['Q4 · Human Rights Risk',     'Assess human rights risk in destination country',                     'Decision Point C → human rights risk data identified?'],
    ['Q5 · Enforcement',           'Assess enforceability of Article 46 transfer mechanism',              'Decision Points D & E → enforceability risk data?'],
    ['Q6 · Exceptions',            'If significant risk data exists, check if any exceptions apply',      'Decision Point F → final go / no-go'],
    ['✅ Summary',                   'All decision points and final TRA outcome',                           'Final TRA conclusion'],
    ['Lookups',                     'Dropdown lists and PI category scores — do not edit values',          'Source of truth for all data validation'],
  ];
  nav.forEach(function(row) {
    sh.getRange(r, 1).setValue(row[0]).setBackground(BG).setFontColor(CYN).setFontSize(9).setWrap(true);
    sh.getRange(r, 2).setValue(row[1]).setBackground(BG).setFontColor(FG).setFontSize(9).setWrap(true);
    sh.getRange(r, 3).setValue(row[2]).setBackground(BG).setFontColor(CMT).setFontSize(9).setWrap(true);
    sh.getRange(r, 4).setBackground(BG);
    sh.setRowHeight(r++, 26);
  });

  spacer(sh, r++);
  h2(sh, r++, 1, 'Score key (1–5)');
  for (let s = 1; s <= 5; s++) {
    sh.getRange(r, 1).setValue(SLB[s])
      .setBackground(SBG[s]).setFontColor(SFG[s]).setFontSize(10).setFontWeight('bold');
    for (let c = 2; c <= 4; c++) sh.getRange(r, c).setBackground(SBG[s]);
    r++;
  }

  spacer(sh, r++);
  h2(sh, r++, 1, 'Cell colour guide');
  var guide = [
    [SEL, FG,  'Input field — enter your data here'],
    [BLK, CYN, 'Auto-calculated — do not edit directly'],
    [CMT, FG,  'Guidance / label'],
    [PRP, BG,  'Section heading'],
  ];
  guide.forEach(function(g) {
    sh.getRange(r, 1).setValue('  ' + g[2]).setBackground(g[0]).setFontColor(g[1]).setFontSize(10);
    for (let c = 2; c <= 4; c++) sh.getRange(r, c).setBackground(g[0]);
    r++;
  });

  spacer(sh, r++);
  h2(sh, r++, 1, 'Decision Point thresholds (auto-calculated in Q2–Q6)');
  var thresholds = [
    ['Max score ≤ 2', 'Low harm risk → MAY PROCEED (Decision Point A1)'],
    ['Max score = 3', 'Moderate harm risk → go to Q3 (Decision Point A2)'],
    ['Max score ≥ 4', 'High harm risk → go to Q3 (Decision Point A3)'],
  ];
  thresholds.forEach(function(t) {
    sh.getRange(r, 1).setValue(t[0]).setBackground(BG).setFontColor(YLW).setFontSize(10).setFontWeight('bold').setWrap(true);
    sh.getRange(r, 2).setValue(t[1]).setBackground(BG).setFontColor(FG).setFontSize(10).setWrap(true);
    sh.getRange(r, 3).setBackground(BG);
    sh.getRange(r, 4).setBackground(BG);
    r++;
  });
}

// ─── Q1 · TRANSFER DETAILS ────────────────────────────────────────────────────

function buildQ1(sh) {
  cw(sh, [[1,160],[2,200],[3,160]]);
  baseStyle(sh, 130, 3);
  sh.setFrozenRows(2);

  h1(sh, 1, 1, 'Q1 · Specific circumstances of the restricted transfer');
  sh.getRange(1, 2).setBackground(PRP);
  sh.getRange(1, 3).setValue('→ Q2').setBackground(PRP).setFontColor(YLW).setFontWeight('bold').setHorizontalAlignment('right');

  sh.getRange(2, 1).setValue('Table 1 — complete all fields. You may cross-refer to your IDTA or Article 46 mechanism.')
    .setBackground(SEL).setFontColor(CMT).setFontSize(9).setFontStyle('italic');
  sh.getRange(2, 2).setBackground(SEL);
  sh.getRange(2, 3).setBackground(SEL);

  var r = 3;

  // ── Importer details ──────────────────────────────────────────────────────
  h2(sh, r, 1, 'IMPORTER DETAILS');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  r++;

  function field(labelTxt, guidanceTxt, ddName) {
    lbl(sh, r, 1, labelTxt);
    sh.getRange(r, 3).setValue(guidanceTxt)
      .setBackground(BG).setFontColor(CMT).setFontSize(8).setFontStyle('italic').setWrap(true);
    r++;
    if (ddName) {
      ddNamed(sh, r, 2, ddName);
    } else {
      inp(sh, r, 2);
    }
    sh.getRange(r, 1).setBackground(BG);
    sh.getRange(r, 3).setBackground(BG);
    sh.setRowHeight(r++, 30);
    spacer(sh, r++);
  }

  field('(1) Name of importer',          'Who is the personal information going to?');
  field('(2) Destination country',        'Country (or countries) the PI is going to');
  field('(3) Status of importer',         'Select importer type',                          'DD_IMPORTER_STATUS');
  field('(4) Importer organisation type', 'What kind of organisation? Add detail in notes', 'DD_ORG_TYPE');

  lbl(sh, r, 1, '(4) Organisation type — notes');
  sh.getRange(r, 3).setValue('e.g. name of group, size, regulator')
    .setBackground(BG).setFontColor(CMT).setFontSize(8).setFontStyle('italic');
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r++, 40);
  spacer(sh, r++);

  lbl(sh, r, 1, '(5) Importer\'s relevant activities');
  sh.getRange(r, 3).setValue('What will the importer do with the PI? Describe their activities / services.')
    .setBackground(BG).setFontColor(CMT).setFontSize(8).setFontStyle('italic').setWrap(true);
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r, 60);
  r++;
  spacer(sh, r++);

  // ── People ────────────────────────────────────────────────────────────────
  h2(sh, r, 1, 'DETAILS OF PEOPLE THE INFORMATION IS ABOUT');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  r++;

  lbl(sh, r, 1, '(6a) Vulnerability status');
  sh.getRange(r, 3).setValue('Select all that apply')
    .setBackground(BG).setFontColor(CMT).setFontSize(8);
  r++;
  ddNamed(sh, r, 2, 'DD_VULN');
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r++, 28);
  spacer(sh, r++);

  h3(sh, r, 1, '(6b) Categories of people — tick all that apply');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setValue('✔').setBackground(BG).setFontColor(CMT).setFontSize(9).setHorizontalAlignment('center');
  r++;

  var peopleCats = [
    'Staff (incl. volunteers, agents, temp / casual workers)',
    'Customers and clients (incl. their staff)',
    'Suppliers (incl. their staff)',
    'Members or supporters',
    'Shareholders',
    'Relatives, guardians and associates',
    'Complainants, correspondents and enquirers',
    'Experts and witnesses',
    'Advisers, consultants and professional experts',
    'Patients',
    'Students and pupils',
    'Offenders and suspected offenders',
    'Children and vulnerable adults',
    'Other (describe below)',
  ];
  peopleCats.forEach(function(cat) {
    sh.getRange(r, 1).setValue(cat).setBackground(BG).setFontColor(FG).setFontSize(10).setWrap(true);
    ck(sh, r, 2);
    sh.getRange(r, 3).setBackground(BG);
    r++;
  });

  lbl(sh, r, 1, '(6c) Other categories (describe):');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r++, 36);
  spacer(sh, r++);

  // ── Volume ────────────────────────────────────────────────────────────────
  h2(sh, r, 1, 'VOLUME');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  r++;

  var volFields = [
    ['(7a) PI categories per person',       'Count when completing Q2 Table 2'],
    ['(7b) Number of people per transfer',  'Estimated or actual — note which'],
    ['(7c) Total people over contract term','Estimated or actual — note which'],
  ];
  volFields.forEach(function(vf) {
    lbl(sh, r, 1, vf[0]);
    sh.getRange(r, 3).setValue(vf[1]).setBackground(BG).setFontColor(CMT).setFontSize(8);
    r++;
    inp(sh, r, 2);
    sh.getRange(r, 1).setBackground(BG);
    sh.getRange(r, 3).setBackground(BG);
    sh.setRowHeight(r++, 26);
  });
  spacer(sh, r++);

  // ── Duration ──────────────────────────────────────────────────────────────
  h2(sh, r, 1, 'DURATION');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  r++;

  lbl(sh, r, 1, '(8) Frequency of transfers');
  sh.getRange(r, 3).setValue('How often will transfers occur?')
    .setBackground(BG).setFontColor(CMT).setFontSize(8);
  r++;
  ddNamed(sh, r, 2, 'DD_FREQUENCY');
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r++, 26);

  lbl(sh, r, 1, '(8) Interval / period details');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r++, 30);
  spacer(sh, r++);

  lbl(sh, r, 1, '(9) Duration of arrangement with importer');
  sh.getRange(r, 3).setValue('How long can the importer receive / access the PI?')
    .setBackground(BG).setFontColor(CMT).setFontSize(8);
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r++, 30);
  spacer(sh, r++);

  // ── Protections ───────────────────────────────────────────────────────────
  h2(sh, r, 1, 'PROTECTIONS FOR THE TRANSFERRED PERSONAL INFORMATION');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  r++;

  var protFields = [
    ['(10) Format of personal information',    'e.g. plain text, encrypted, pseudonymised'],
    ['(11) Transfer process',                   'e.g. email, SFTP, website encryption, remote access to UK-stored data'],
    ['(12) Exporter\'s TOMs',                   'Your technical and organisational security measures before transfer'],
    ['(13) Importer\'s TOMs',                   'Importer\'s technical and organisational security measures after receipt'],
  ];
  protFields.forEach(function(pf) {
    lbl(sh, r, 1, pf[0]);
    sh.getRange(r, 3).setValue(pf[1]).setBackground(BG).setFontColor(CMT).setFontSize(8).setFontStyle('italic').setWrap(true);
    r++;
    inp(sh, r, 2);
    sh.getRange(r, 1).setBackground(BG);
    sh.getRange(r, 3).setBackground(BG);
    sh.setRowHeight(r, 52);
    r++;
    spacer(sh, r++);
  });

  lbl(sh, r, 1, '(14) Categories of PI being transferred');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setValue('→ List these in column B of Q2 · PI Risk Scores')
    .setBackground(BG).setFontColor(CYN).setFontSize(9).setWrap(true);
}

// ─── Q2 · PI RISK SCORES ──────────────────────────────────────────────────────

function buildQ2(sh, ss) {
  // Cols: A=# B=Category C=ICO Score D=Special? E=Suggested F=Aggravating G=Mitigating H=Final Score I=Notes
  cw(sh, [[1,28],[2,200],[3,75],[4,65],[5,75],[6,160],[7,160],[8,80],[9,160]]);
  baseStyle(sh, 80, 9);
  sh.setFrozenRows(3);

  h1(sh, 1, 1, 'Q2 · What is the risk level in the personal information you are transferring?');
  for (let c = 2; c <= 9; c++) sh.getRange(1, c).setBackground(PRP);
  sh.getRange(1, 9).setValue('← Q1  → Q3').setBackground(PRP).setFontColor(YLW)
    .setFontWeight('bold').setHorizontalAlignment('right');

  // Score key row
  sh.getRange(2, 1).setValue('Score key:').setBackground(SEL).setFontColor(CMT).setFontSize(9);
  for (let s = 1; s <= 5; s++) {
    sh.getRange(2, 1 + s).setValue(SLB[s])
      .setBackground(SBG[s]).setFontColor(SFG[s]).setFontSize(8).setFontWeight('bold').setHorizontalAlignment('center');
  }
  sh.getRange(2, 7).setValue('Adjust ±1 from ICO initial score based on factors')
    .setBackground(SEL).setFontColor(CMT).setFontSize(8).setWrap(true);
  sh.getRange(2, 8).setBackground(SEL);
  sh.getRange(2, 9).setBackground(SEL);
  sh.setRowHeight(2, 22);

  // Table column headers
  var hdrs = ['#','PI Category','ICO Initial\nScore (auto)','Special\nCat?','Suggested\nScore',
              'Aggravating Factors','Mitigating Factors','Your Final\nScore','Notes'];
  hdrs.forEach(function(h, i) {
    sh.getRange(3, 1 + i).setValue(h)
      .setBackground(SEL).setFontColor(CYN).setFontWeight('bold')
      .setFontSize(9).setWrap(true).setHorizontalAlignment('center');
  });
  sh.setRowHeight(3, 40);

  var DATA_START = 4;
  var NUM_ROWS   = 15;

  for (var i = 0; i < NUM_ROWS; i++) {
    var r = DATA_START + i;

    // A: row number
    sh.getRange(r, 1).setValue(i + 1)
      .setBackground(SEL).setFontColor(CMT).setFontSize(9).setHorizontalAlignment('center');

    // B: category name — dropdown from PI_NAMES
    ddNamed(sh, r, 2, 'PI_NAMES');

    // C: ICO initial score — VLOOKUP into PI_SCORES (col A=name, col B=score)
    auto(sh, r, 3,
      '=IF(B' + r + '="","",IFERROR(VLOOKUP(B' + r + ',PI_SCORES,2,0),"?"))');
    sh.getRange(r, 3).setHorizontalAlignment('center');
    cfScore(sh.getRange(r, 3));

    // D: special category?
    auto(sh, r, 4,
      '=IF(B' + r + '="","",IFERROR(IF(VLOOKUP(B' + r + ',PI_SCORES,2,0),""),"")&' +
      'IF(B' + r + '="","",IFERROR(IF(INDEX(\'Lookups\'!C:C,MATCH(B' + r + ',PI_NAMES,0))="Yes","⚠️ Yes","No"),"No")))');
    // Simpler formula: look up col C in Lookups (Special Cat?)
    sh.getRange(r, 4).setFormula(
      '=IF(B' + r + '="","",IFERROR(IF(INDEX(\'Lookups\'!$C:$C,MATCH(B' + r + ',PI_NAMES,0))="Yes","⚠️ Yes","No"),"No"))');
    sh.getRange(r, 4).setBackground(BLK).setFontColor(CYN).setFontSize(10).setHorizontalAlignment('center');

    // Special cat conditional format
    var rules4 = sh.getConditionalFormatRules();
    rules4.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('⚠️ Yes').setBackground(PNK).setFontColor(BG).setBold(true)
      .setRanges([sh.getRange(r, 4)]).build());
    sh.setConditionalFormatRules(rules4);

    // E: suggested score (ICO score +1 if special category, capped at 5)
    auto(sh, r, 5,
      '=IF(B' + r + '="","",IF(ISNUMBER(C' + r + '),' +
      'IF(D' + r + '="⚠️ Yes",MIN(5,C' + r + '+1),C' + r + '),""))');
    sh.getRange(r, 5).setHorizontalAlignment('center');
    cfScore(sh.getRange(r, 5));

    // F: aggravating factors (free text)
    inp(sh, r, 6);

    // G: mitigating factors (free text)
    inp(sh, r, 7);

    // H: final score — dropdown 1–5 from Lookups
    ddNamed(sh, r, 8, 'DD_SCORES');
    sh.getRange(r, 8).setHorizontalAlignment('center');
    cfScore(sh.getRange(r, 8));

    // I: notes
    inp(sh, r, 9);

    sh.setRowHeight(r, 36);
  }

  var AFTER = DATA_START + NUM_ROWS;
  spacer(sh, AFTER);

  // Additional notes
  h3(sh, AFTER + 1, 1, 'Additional notes:');
  for (let c = 2; c <= 9; c++) sh.getRange(AFTER + 1, c).setBackground(BG);
  inp(sh, AFTER + 2, 2);
  for (let c = 3; c <= 9; c++) sh.getRange(AFTER + 2, c).setBackground(BG);
  sh.setRowHeight(AFTER + 2, 50);

  spacer(sh, AFTER + 3);

  // Max score calculation
  var MS_ROW = AFTER + 4;
  h3(sh, MS_ROW, 1, 'Max score across all categories:');
  for (let c = 2; c <= 9; c++) sh.getRange(MS_ROW, c).setBackground(BG);
  sh.getRange(MS_ROW, 2)
    .setFormula('=IFERROR(MAX(IFERROR(VALUE(H' + DATA_START + ':H' + (DATA_START + NUM_ROWS - 1) + '),0)),"")')
    .setBackground(BLK).setFontColor(CYN).setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center');
  cfScore(sh.getRange(MS_ROW, 2));
  nr(ss, 'MAX_SCORE', sh.getRange(MS_ROW, 2));

  // Decision Point A
  var DPA_ROW = MS_ROW + 2;
  h2(sh, DPA_ROW, 1, '⬛  DECISION POINT A');
  for (let c = 2; c <= 9; c++) sh.getRange(DPA_ROW, c).setBackground(SEL);

  var dpAf =
    '=IF(MAX_SCORE="","⏳ Complete Table 2 above — enter PI categories (col B) and final scores (col H)",' +
    'IF(MAX_SCORE<=2,"✅ A1 — All data is low harm risk (max score: "&MAX_SCORE&"). MAY PROCEED with the restricted transfer. Record this as your TRA final decision.",' +
    'IF(MAX_SCORE=3,"⚠️ A2 — Data includes moderate harm risk (max score: "&MAX_SCORE&"). No high harm risk. → Go to Q3 · Investigation Level.",' +
    '"🔴 A3 — Data includes high harm risk (max score: "&MAX_SCORE&"). → Go to Q3 · Investigation Level.")))';

  dpCell(sh, DPA_ROW + 1, 1, dpAf);
  for (let c = 2; c <= 9; c++) {
    sh.getRange(DPA_ROW + 1, c).setBackground(BG);
    sh.setRowHeight(DPA_ROW + 1, 65);
  }
}

// ─── Q3 · INVESTIGATION LEVEL ─────────────────────────────────────────────────

function buildQ3(sh, ss) {
  cw(sh, [[1,180],[2,200],[3,160],[4,120]]);
  baseStyle(sh, 80, 4);
  sh.setFrozenRows(2);

  h1(sh, 1, 1, 'Q3 · What is a reasonable and proportionate level of investigation?');
  sh.getRange(1, 2).setBackground(PRP);
  sh.getRange(1, 3).setBackground(PRP);
  sh.getRange(1, 4).setValue('← Q2  → Q4').setBackground(PRP).setFontColor(YLW).setFontWeight('bold').setHorizontalAlignment('right');

  sh.getRange(2, 1).setValue('Tables 3 & 4 — three factors determine the required investigation level.')
    .setBackground(SEL).setFontColor(CMT).setFontSize(9).setFontStyle('italic');
  sh.getRange(2, 2).setBackground(SEL);
  sh.getRange(2, 3).setBackground(SEL);
  sh.getRange(2, 4).setBackground(SEL);

  var r = 3;

  // Factors explanation
  h2(sh, r, 1, 'The three factors');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  sh.getRange(r, 4).setBackground(SEL);
  r++;

  var factors = [
    ['Factor 1', 'Risk level in PI (max score from Q2, Decision Point A)'],
    ['Factor 2', 'Size of your organisation (SME = Tier 1 or 2 data protection fee payer)'],
    ['Factor 3', 'Total volume of PI being transferred'],
  ];
  factors.forEach(function(f) {
    sh.getRange(r, 1).setValue(f[0]).setBackground(BG).setFontColor(YLW).setFontSize(10).setFontWeight('bold');
    sh.getRange(r, 2).setValue(f[1]).setBackground(BG).setFontColor(FG).setFontSize(10).setWrap(true);
    sh.getRange(r, 3).setBackground(BG);
    sh.getRange(r, 4).setBackground(BG);
    r++;
  });
  spacer(sh, r++);

  // Your inputs
  h2(sh, r, 1, 'Your inputs (Factors 2 & 3)');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  sh.getRange(r, 4).setBackground(SEL);
  r++;

  lbl(sh, r, 1, 'Max score from Q2 (auto):');
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  r++;
  auto(sh, r, 2, '=IFERROR(MAX_SCORE,"⏳ Complete Q2 first")');
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  cfScore(sh.getRange(r, 2));
  sh.setRowHeight(r++, 26);
  spacer(sh, r++);

  lbl(sh, r, 1, 'Organisation size:');
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  r++;
  ddNamed(sh, r, 2, 'DD_BIZ_SIZE');
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  sh.setRowHeight(r, 26);
  nr(ss, 'BIZ_SIZE', sh.getRange(r, 2));
  r++;
  spacer(sh, r++);

  lbl(sh, r, 1, 'Transfer volume:');
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  r++;
  ddNamed(sh, r, 2, 'DD_XFER_VOL');
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  sh.setRowHeight(r, 26);
  nr(ss, 'XFER_VOL', sh.getRange(r, 2));
  r++;
  spacer(sh, r++);

  // Recommended level (auto)
  h2(sh, r, 1, 'Recommended investigation level (auto)');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  sh.getRange(r, 4).setBackground(SEL);
  r++;

  var recF =
    '=IF(OR(BIZ_SIZE="",XFER_VOL="",MAX_SCORE=""),"⏳ Complete inputs above",' +
    'IF(MAX_SCORE<=2,"No investigation needed — see Decision Point A",' +
    'IF(MAX_SCORE=3,' +
      'IF(ISNUMBER(SEARCH("SME",BIZ_SIZE)),"Level 1","Level 2"),' +
    // max score >= 4
    'IF(ISNUMBER(SEARCH("SME",BIZ_SIZE)),' +
      'IF(ISNUMBER(SEARCH("Low",XFER_VOL)),"Level 2","Level 3"),' +
      '"Level 3"))))';

  auto(sh, r, 2, recF);
  sh.getRange(r, 2).setFontSize(11).setFontWeight('bold');
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  sh.setRowHeight(r++, 30);
  spacer(sh, r++);

  // Investigation matrix table (read-only reference)
  h2(sh, r, 1, 'Table 3: Investigation matrix (reference)');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  sh.getRange(r, 4).setBackground(SEL);
  r++;

  var matHdrs = ['', 'All PI score ≤ 2', 'PI includes score 3 (no ≥4)', 'PI includes score ≥ 4'];
  matHdrs.forEach(function(h, i) {
    sh.getRange(r, 1 + i).setValue(h)
      .setBackground(i === 0 ? BG : SEL).setFontColor(i === 0 ? BG : CYN)
      .setFontWeight('bold').setFontSize(9).setWrap(true).setHorizontalAlignment('center');
  });
  sh.setRowHeight(r++, 30);

  var matRows = [
    ['SME',           'No investigation needed', 'Level 1', 'Low vol → Level 2\nHigh vol → Level 3'],
    ['Large business','No investigation needed', 'Level 2', 'Level 3'],
  ];
  matRows.forEach(function(mrow) {
    mrow.forEach(function(cell, i) {
      sh.getRange(r, 1 + i).setValue(cell)
        .setBackground(i === 0 ? SEL : BG)
        .setFontColor(i === 0 ? FG : CMT)
        .setFontSize(9).setWrap(true).setHorizontalAlignment(i === 0 ? 'center' : 'left');
    });
    sh.setRowHeight(r++, 34);
  });
  spacer(sh, r++);

  // Levels of investigation (Table 4 — condensed reference)
  h2(sh, r, 1, 'Table 4: Levels of investigation (reference)');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  sh.getRange(r, 4).setBackground(SEL);
  r++;

  var levels = [
    ['Level 1', 'Own knowledge · FCDO Human Rights & Democracy Report · DIT Exporting country guides · ≥1 human rights report (e.g. Amnesty International)'],
    ['Level 2', 'Everything in Level 1 + additional human rights reports (charities, US State Dept) + newspaper reports. Consider report biases.'],
    ['Level 3\n(i)',  'Everything in Levels 1 & 2 + detailed human rights analysis. Professional advice likely needed.'],
    ['Level 3\n(ii)', 'Mark ALL high-score PI (≥4) as significant risk data. Run Level 2 for the rest. Skip Q5. High-score data can only transfer if exception applies (Q6).'],
  ];
  levels.forEach(function(lv) {
    sh.getRange(r, 1).setValue(lv[0])
      .setBackground(SEL).setFontColor(PRP).setFontWeight('bold').setFontSize(9).setWrap(true).setHorizontalAlignment('center');
    sh.getRange(r, 2).setValue(lv[1])
      .setBackground(BG).setFontColor(FG).setFontSize(9).setWrap(true);
    sh.getRange(r, 3).setBackground(BG);
    sh.getRange(r, 4).setBackground(BG);
    sh.setRowHeight(r++, 52);
  });
  spacer(sh, r++);

  // Decision Point B — user selects their level
  h2(sh, r, 1, '⬛  DECISION POINT B');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  sh.getRange(r, 4).setBackground(SEL);
  r++;

  lbl(sh, r, 1, 'Select the investigation level you will carry out:');
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  r++;
  ddNamed(sh, r, 2, 'DD_INV_LEVEL');
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  sh.setRowHeight(r, 30);
  nr(ss, 'DP_B', sh.getRange(r, 2));
  r++;
  spacer(sh, r++);

  lbl(sh, r, 1, 'Reasons this level is reasonable and proportionate:');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  sh.setRowHeight(r++, 52);
  spacer(sh, r++);

  lbl(sh, r, 1, 'Resources used in investigation:');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.getRange(r, 4).setBackground(BG);
  sh.setRowHeight(r, 65);
}

// ─── Q4 · HUMAN RIGHTS RISK ───────────────────────────────────────────────────

function buildQ4(sh, ss) {
  cw(sh, [[1,180],[2,200],[3,160]]);
  baseStyle(sh, 100, 3);
  sh.setFrozenRows(2);

  h1(sh, 1, 1, 'Q4 · Is the transfer significantly increasing the risk of a human rights breach?');
  sh.getRange(1, 2).setBackground(PRP);
  sh.getRange(1, 3).setValue('← Q3  → Q5').setBackground(PRP).setFontColor(YLW).setFontWeight('bold').setHorizontalAlignment('right');

  sh.getRange(2, 1).setValue('Table 6: Record of investigation and human rights risk assessment.')
    .setBackground(SEL).setFontColor(CMT).setFontSize(9).setFontStyle('italic');
  sh.getRange(2, 2).setBackground(SEL);
  sh.getRange(2, 3).setBackground(SEL);

  var r = 3;

  // Level 3 Option (ii) warning
  sh.getRange(r, 1).setValue('⚠️ If you selected Level 3 Option (ii) at Q3: go directly to Decision Point E(2) on Q5. Then complete Q4 for low / moderate data only.')
    .setBackground(SEL).setFontColor(YLW).setFontSize(9).setWrap(true);
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  sh.setRowHeight(r++, 36);
  spacer(sh, r++);

  // Human rights reference table (Table 5)
  h2(sh, r, 1, 'Table 5: Key human rights (simplified ECHR) — reference only');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  r++;

  var rights = [
    ['Art 2',  'Right to life (note: is the death penalty available?)'],
    ['Art 3',  'Prohibition of torture'],
    ['Art 4',  'Prohibition of slavery and forced labour'],
    ['Art 5',  'Right to liberty and security'],
    ['Art 6',  'Right to a fair trial (presumption of innocence, right to legal representation)'],
    ['Art 7',  'No punishment without law'],
    ['Art 8',  'Right to respect for private / family life, home and correspondence'],
    ['Art 9',  'Freedom of thought, conscience and religion'],
    ['Art 10', 'Freedom of expression (incl. press freedom)'],
    ['Art 11', 'Freedom of assembly and association (incl. trade unions)'],
    ['Art 12', 'Right to marry'],
    ['Art 13', 'Right to an effective remedy'],
    ['Art 14', 'No discrimination'],
  ];
  rights.forEach(function(rt) {
    sh.getRange(r, 1).setValue(rt[0])
      .setBackground(SEL).setFontColor(PRP).setFontWeight('bold').setFontSize(9);
    sh.getRange(r, 2).setValue(rt[1])
      .setBackground(BG).setFontColor(CMT).setFontSize(9).setWrap(true);
    sh.getRange(r, 3).setBackground(BG);
    r++;
  });
  spacer(sh, r++);

  // Investigation level reference
  lbl(sh, r, 1, 'Investigation level used (from Q3, auto):');
  sh.getRange(r, 3).setBackground(BG);
  r++;
  auto(sh, r, 2, '=IFERROR(DP_B,"⏳ Complete Q3 first")');
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r++, 26);
  spacer(sh, r++);

  // Record of investigation
  h2(sh, r, 1, 'Record of investigation');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  r++;

  lbl(sh, r, 1, 'Resources used:');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r++, 65);
  spacer(sh, r++);

  // Key Question 1
  h3(sh, r, 1, 'Key Question 1: Any concerns about human rights in the destination country?');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setValue('From your investigation — see Table 5 above')
    .setBackground(BG).setFontColor(CMT).setFontSize(8).setFontStyle('italic');
  r++;
  ddNamed(sh, r, 2, 'DD_KQ4_1');
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r, 26);
  nr(ss, 'KQ4_1', sh.getRange(r, 2));
  r++;

  lbl(sh, r, 1, '  Which Articles are relevant? (list and describe concerns)');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r++, 52);
  spacer(sh, r++);

  // Key Question 2
  h3(sh, r, 1, 'Key Question 2: By making this transfer, are you making the risk SIGNIFICANTLY WORSE for the people?');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setValue('Consider: more likely a breach will happen, OR more severe if it did. Risk must be clear, meaningful and linked to this transfer.')
    .setBackground(BG).setFontColor(CMT).setFontSize(8).setFontStyle('italic').setWrap(true);
  r++;
  ddNamed(sh, r, 2, 'DD_KQ4_2');
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r, 26);
  nr(ss, 'KQ4_2', sh.getRange(r, 2));
  r++;

  lbl(sh, r, 1, '  Which PI categories cause the increase in risk (if applicable):');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r++, 40);
  spacer(sh, r++);

  // Decision Point C (auto)
  h2(sh, r, 1, '⬛  DECISION POINT C');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  r++;

  var dpCf =
    '=IF(OR(KQ4_1="",KQ4_2=""),"⏳ Complete Key Questions above",' +
    'IF(ISNUMBER(SEARCH("No concerns",KQ4_1)),"✅ C1 — No human rights concerns in destination country. → Go to Q5 · Enforcement.",' +
    'IF(ISNUMBER(SEARCH("does NOT",KQ4_2)),"✅ C2 — Human rights concerns exist, but transfer does NOT significantly increase the risk. → Go to Q5 · Enforcement.",' +
    'IF(ISNUMBER(SEARCH("ALL categories",KQ4_2)),"🔴 C3 — Transfer significantly increases human rights risk for ALL PI categories. → Go to Decision Point E(3). Then go to Q5 (or Q6 if Level 3 Option ii).",' +
    'IF(ISNUMBER(SEARCH("SOME categories",KQ4_2)),"🔴 C4 — Transfer significantly increases human rights risk for SOME PI categories. → Go to Decision Point E(4). List those categories below. Then go to Q5 (or Q6 if Level 3 Option ii).",' +
    '"⏳ Review answers above")))))';

  dpCell(sh, r, 1, dpCf);
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  spacer(sh, r++);

  lbl(sh, r, 1, 'List of human rights risk data (if C3 or C4 — required):');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r, 40);
}

// ─── Q5 · ENFORCEMENT ─────────────────────────────────────────────────────────

function buildQ5(sh, ss) {
  // Cols: A=question, B=answer(dropdown), C=notes
  cw(sh, [[1,200],[2,200],[3,160]]);
  baseStyle(sh, 100, 3);
  sh.setFrozenRows(2);

  h1(sh, 1, 1, 'Q5 · Can the Article 46 transfer mechanism be enforced against the importer?');
  sh.getRange(1, 2).setBackground(PRP);
  sh.getRange(1, 3).setValue('← Q4  → Q6').setBackground(PRP).setFontColor(YLW).setFontWeight('bold').setHorizontalAlignment('right');

  sh.getRange(2, 1).setValue('Table 7: Enforcement questionnaire.')
    .setBackground(SEL).setFontColor(CMT).setFontSize(9).setFontStyle('italic');
  sh.getRange(2, 2).setBackground(SEL);
  sh.getRange(2, 3).setBackground(SEL);

  var r = 3;

  sh.getRange(r, 1).setValue('⚠️ If you selected Level 3 Option (ii) at Q3, you do NOT need to answer Q5. Skip to Q6 · Exceptions.')
    .setBackground(SEL).setFontColor(YLW).setFontSize(9).setWrap(true);
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  sh.setRowHeight(r++, 30);
  spacer(sh, r++);

  // Column headers
  sh.getRange(r, 1).setValue('Enforcement question').setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);
  sh.getRange(r, 2).setValue('Your answer').setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);
  sh.getRange(r, 3).setValue('Notes').setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);
  r++;

  function eq(labelTxt, guidanceTxt, ddName, namedRange) {
    lbl(sh, r, 1, labelTxt);
    sh.getRange(r, 2).setBackground(BG);
    sh.getRange(r, 3).setValue(guidanceTxt).setBackground(BG).setFontColor(CMT).setFontSize(8).setFontStyle('italic').setWrap(true);
    r++;
    ddNamed(sh, r, 2, ddName);
    sh.getRange(r, 1).setBackground(BG);
    inp(sh, r, 3);
    sh.setRowHeight(r, 28);
    if (namedRange) nr(ss, namedRange, sh.getRange(r, 2));
    r++;
    spacer(sh, r++);
  }

  eq('EQ1: Are you transferring ONLY PI that is low or moderate risk (max score ≤ 3)?',
     'If yes, the likelihood of needing to enforce in the destination country is low.',
     'DD_EQ5_1', 'EQ5_1');

  eq('EQ2: In your investigation, have you found issues about rule of law, court independence, or delays in the destination country?',
     'Consider independence of judiciary, access to justice, case backlog.',
     'DD_EQ5_2', 'EQ5_2');

  // EQ3 with factor checkboxes
  h3(sh, r, 1, 'EQ3: Is there a HIGH likelihood the importer will accept a UK Court decision or UK arbitration award?');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setValue('Consider factors (a)–(d) below — tick those that are satisfied.')
    .setBackground(BG).setFontColor(CMT).setFontSize(8).setFontStyle('italic').setWrap(true);
  r++;

  var eq3Factors = [
    '(a) Importer has insurance covering UK Court / arbitration claims (check annually)',
    '(b) Importer has evidence of always accepting past UK Court / arbitration decisions',
    '(c) Importer must comply with professional / regulatory rules — you could complain to oversight body',
    '(d) Strong commercial reasons for importer to accept UK Court / arbitration decision',
  ];
  eq3Factors.forEach(function(f) {
    sh.getRange(r, 1).setValue(f).setBackground(BG).setFontColor(FG).setFontSize(9).setWrap(true);
    ck(sh, r, 2);
    sh.getRange(r, 3).setBackground(BG);
    sh.setRowHeight(r++, 26);
  });
  lbl(sh, r, 1, '  Overall answer to EQ3:');
  sh.getRange(r, 3).setBackground(BG);
  r++;
  ddNamed(sh, r, 2, 'DD_EQ5_3');
  sh.getRange(r, 1).setBackground(BG);
  inp(sh, r, 3);
  sh.setRowHeight(r, 28);
  nr(ss, 'EQ5_3', sh.getRange(r, 2));
  r++;
  spacer(sh, r++);

  eq('EQ4: Are there other factors making it very unlikely a claim would need to be brought in the destination country?',
     'e.g. strong contractual protections, regulatory regime, other safeguards.',
     'DD_EQ5_4', 'EQ5_4');

  // Decision Point D (auto)
  h2(sh, r, 1, '⬛  DECISION POINT D');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  r++;

  var dpDf =
    '=IF(EQ5_1="","⏳ Complete enforcement questions above",' +
    'IF(ISNUMBER(SEARCH("low / moderate",EQ5_1)),"✅ D1 — Only low / moderate risk data. Low likelihood of needing to enforce in destination country.",' +
    'IF(ISNUMBER(SEARCH("No concerns",EQ5_2)),"✅ D2 — No concerns about rule of law / court independence in destination country.",' +
    'IF(ISNUMBER(SEARCH("high likelihood",EQ5_3)),"✅ D3 — High likelihood importer will accept UK Court / arbitration award.",' +
    'IF(ISNUMBER(SEARCH("Yes — other",EQ5_4)),"✅ D4 — Very unlikely a claim would need to be brought in destination country (other factors).",' +
    'IF(ISNUMBER(SEARCH("No →",EQ5_4)),"🔴 D5 — Concerns about enforceability. You and the people may not be able to enforce the mechanism in the UK or destination country. → Go to Decision Point E(5).",' +
    '"⏳ Review answers above"))))))';

  dpCell(sh, r, 1, dpDf);
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  spacer(sh, r++);

  // Decision Point E — user selects
  h2(sh, r, 1, '⬛  DECISION POINT E — Have you identified any "significant risk data"?');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  r++;

  sh.getRange(r, 1).setValue('Significant risk data = human rights risk data (Q4) OR enforceability risk data (D5). Must pass Q6 exceptions test to be transferred.')
    .setBackground(BG).setFontColor(CMT).setFontSize(9).setFontStyle('italic').setWrap(true);
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r++, 30);

  lbl(sh, r, 1, 'Select Decision Point E:');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  ddNamed(sh, r, 2, 'DD_DP_E');
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r, 28);
  nr(ss, 'DP_E', sh.getRange(r, 2));
  r++;
  spacer(sh, r++);

  var dpEresultF =
    '=IF(DP_E="","⏳ Select Decision Point E above",' +
    'IF(ISNUMBER(SEARCH("E1",DP_E)),"✅ E1 — No significant risk data. MAY PROCEED with the transfer.",' +
    '"🔴 Significant risk data identified: "&DP_E&" → Complete Q6 · Exceptions."))';

  dpCell(sh, r, 1, dpEresultF);
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  spacer(sh, r++);

  lbl(sh, r, 1, 'List all significant risk data categories (recommended):');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r, 40);
}

// ─── Q6 · EXCEPTIONS ──────────────────────────────────────────────────────────

function buildQ6(sh, ss) {
  // Cols: A=exception name, B=description, C=applies?, D=which data, E=benefit outweighs?, F=reasons
  cw(sh, [[1,130],[2,200],[3,80],[4,140],[5,100],[6,180]]);
  baseStyle(sh, 60, 6);
  sh.setFrozenRows(3);

  h1(sh, 1, 1, 'Q6 · Do any of the exceptions apply to the significant risk data?');
  for (let c = 2; c <= 6; c++) sh.getRange(1, c).setBackground(PRP);
  sh.getRange(1, 6).setValue('← Q5').setBackground(PRP).setFontColor(YLW).setFontWeight('bold').setHorizontalAlignment('right');

  sh.getRange(2, 1).setValue('Table 8: Exceptions checklist.')
    .setBackground(SEL).setFontColor(CMT).setFontSize(9).setFontStyle('italic');
  for (let c = 2; c <= 6; c++) sh.getRange(2, c).setBackground(SEL);

  sh.getRange(3, 1).setValue('⚠️ Only complete if significant risk data identified at Decision Point E. Article 46 mechanism provides SOME (not all) appropriate safeguards.')
    .setBackground(SEL).setFontColor(YLW).setFontSize(9).setWrap(true);
  for (let c = 2; c <= 6; c++) sh.getRange(3, c).setBackground(SEL);
  sh.setRowHeight(3, 30);

  var r = 4;

  // Significant risk data reference
  lbl(sh, r, 1, 'Significant risk data (from Q5, Decision Point E, auto):');
  for (let c = 2; c <= 6; c++) sh.getRange(r, c).setBackground(BG);
  r++;
  auto(sh, r, 2, '=IFERROR(DP_E,"⏳ Complete Q5 first")');
  for (let c = 3; c <= 6; c++) sh.getRange(r, c).setBackground(BG);
  sh.getRange(r, 1).setBackground(BG);
  sh.setRowHeight(r++, 26);
  spacer(sh, r++);

  // Table headers
  var thdrs = ['Exception', 'Scenario', 'Applies?\n(Y/N/N-A)', 'Which significant\nrisk data?', 'Benefit outweighs\nall identified risks?', 'Your reasons'];
  thdrs.forEach(function(h, i) {
    sh.getRange(r, 1 + i).setValue(h)
      .setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9).setWrap(true).setHorizontalAlignment('center');
  });
  sh.setRowHeight(r++, 38);

  var exceptions = [
    ['Exception 1\nContractual necessity',
     'Transfer is necessary for performance of a contract between you and the data subject, or to implement pre-contractual measures they requested.'],
    ['Exception 2\nThird-party contract',
     'Transfer is necessary to enter into / perform a contract between you and a third party, where that contract is in the interests of the data subject.'],
    ['Exception 3\nPublic interest',
     'Transfer is necessary for important reasons of public interest.'],
    ['Exception 4\nLegal claims',
     'Transfer is necessary to establish, make or defend a legal claim.'],
    ['Exception 5\nVital interests',
     'Transfer is necessary to protect vital interests of a person who is physically or legally incapable of giving consent to the transfer.'],
    ['Exception 6\nPublic register',
     'Transfer is from a public register meeting relevant legal requirements about access to that register.'],
    ['Exception 7\nCompelling legitimate interests',
     'Transfer necessary for your compelling legitimate interests. LAST RESORT — read ICO guidance carefully. Additional requirements apply.'],
  ];

  exceptions.forEach(function(ex) {
    sh.getRange(r, 1).setValue(ex[0])
      .setBackground(SEL).setFontColor(PRP).setFontWeight('bold').setFontSize(9).setWrap(true).setVerticalAlignment('top');
    sh.getRange(r, 2).setValue(ex[1])
      .setBackground(BG).setFontColor(FG).setFontSize(9).setWrap(true).setVerticalAlignment('top');

    ddNamed(sh, r, 3, 'DD_EXCEPTION_YN');
    sh.getRange(r, 3).setHorizontalAlignment('center');

    inp(sh, r, 4);

    ddNamed(sh, r, 5, 'DD_BENEFIT_YN');

    inp(sh, r, 6);

    sh.setRowHeight(r++, 70);
  });

  spacer(sh, r++);

  // Decision Point F — user selects
  h2(sh, r, 1, '⬛  DECISION POINT F');
  for (let c = 2; c <= 6; c++) sh.getRange(r, c).setBackground(SEL);
  r++;

  lbl(sh, r, 1, 'Select your Decision Point F conclusion:');
  for (let c = 2; c <= 6; c++) sh.getRange(r, c).setBackground(BG);
  r++;
  ddNamed(sh, r, 2, 'DD_DP_F');
  sh.getRange(r, 1).setBackground(BG);
  for (let c = 3; c <= 6; c++) sh.getRange(r, c).setBackground(BG);
  sh.setRowHeight(r, 28);
  r++;
  spacer(sh, r++);

  nr(ss, 'DP_F_VAL', sh.getRange(r - 2, 2));

  var dpFf =
    '=IF(DP_F_VAL="","⏳ Select Decision Point F above",' +
    'IF(ISNUMBER(SEARCH("F1",DP_F_VAL)),' +
    '"✅ MAY PROCEED — One or more exceptions apply to all significant risk data.",' +
    '"❌ CANNOT PROCEED — Exceptions do not apply to all significant risk data. May NOT proceed. Consider removing significant risk data from scope and repeating the TRA."))';

  dpCell(sh, r, 1, dpFf);
  for (let c = 2; c <= 6; c++) sh.getRange(r, c).setBackground(BG);
  r++;
  spacer(sh, r++);

  lbl(sh, r, 1, 'Additional notes:');
  for (let c = 2; c <= 6; c++) sh.getRange(r, c).setBackground(BG);
  r++;
  inp(sh, r, 2);
  for (let c = 3; c <= 6; c++) sh.getRange(r, c).setBackground(BG);
  sh.setRowHeight(r, 52);
}

// ─── SUMMARY ──────────────────────────────────────────────────────────────────

function buildSummary(sh) {
  cw(sh, [[1,160],[2,200],[3,160]]);
  baseStyle(sh, 80, 3);
  sh.setFrozenRows(2);

  h1(sh, 1, 1, '✅ TRA Summary — Decision Points & Final Outcome');
  sh.getRange(1, 2).setBackground(PRP);
  sh.getRange(1, 3).setBackground(PRP);

  sh.getRange(2, 1).setValue('Auto-calculated from your answers across Q1–Q6. Keep this tab as your TRA record.')
    .setBackground(SEL).setFontColor(CMT).setFontSize(9).setFontStyle('italic');
  sh.getRange(2, 2).setBackground(SEL);
  sh.getRange(2, 3).setBackground(SEL);

  var r = 3;

  // Transfer identification
  h2(sh, r, 1, 'Transfer identification');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  r++;

  var idFields = [
    ['Date of assessment', '=TEXT(TODAY(),"dd mmm yyyy")'],
    ['Assessed by',         ''],
    ['Role / organisation', ''],
  ];
  idFields.forEach(function(f) {
    lbl(sh, r, 1, f[0]);
    sh.getRange(r, 3).setBackground(BG);
    r++;
    if (f[1].charAt(0) === '=') {
      auto(sh, r, 2, f[1]);
    } else {
      inp(sh, r, 2);
    }
    sh.getRange(r, 1).setBackground(BG);
    sh.getRange(r, 3).setBackground(BG);
    sh.setRowHeight(r++, 24);
  });
  spacer(sh, r++);

  // Decision points table
  h2(sh, r, 1, 'Decision Points');
  sh.getRange(r, 2).setValue('Question').setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);
  sh.getRange(r, 3).setValue('Result (auto)').setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);
  r++;

  var dps = [
    ['A', 'Q2 — PI Risk Scores',       '=IFERROR(IF(MAX_SCORE<=2,"✅ A1 — Max score "&MAX_SCORE&". MAY PROCEED.",IF(MAX_SCORE=3,"⚠️ A2 — Max score "&MAX_SCORE&". → Go to Q3.",IF(MAX_SCORE>=4,"🔴 A3 — Max score "&MAX_SCORE&". → Go to Q3.","⏳"))),"⏳ Not yet completed")'],
    ['B', 'Q3 — Investigation Level',  '=IFERROR(DP_B,"⏳ Not yet completed")'],
    ['C', 'Q4 — Human Rights Risk',    '=IFERROR(IF(ISNUMBER(SEARCH("C1",KQ4_1)),"✅ C1 — No concerns",IF(ISNUMBER(SEARCH("C2",KQ4_2)),"✅ C2 — No significant increase",IF(ISNUMBER(SEARCH("C3",KQ4_2)),"🔴 C3 — All categories HR risk",IF(ISNUMBER(SEARCH("C4",KQ4_2)),"🔴 C4 — Some categories HR risk","⏳")))),"⏳ Not yet completed")'],
    ['D', 'Q5 — Enforcement',          '=IFERROR(IF(ISNUMBER(SEARCH("D1",EQ5_1)),"✅ D1",IF(ISNUMBER(SEARCH("D2",EQ5_2)),"✅ D2",IF(ISNUMBER(SEARCH("D3",EQ5_3)),"✅ D3",IF(ISNUMBER(SEARCH("D4",EQ5_4)),"✅ D4",IF(ISNUMBER(SEARCH("D5",EQ5_4)),"🔴 D5 — Enforceability concerns","⏳"))))),"⏳ Not yet completed")'],
    ['E', 'Q5 — Significant Risk Data', '=IFERROR(DP_E,"⏳ Not yet completed")'],
    ['F', 'Q6 — Exceptions',            '=IFERROR(DP_F_VAL,"⏳ Not yet completed")'],
  ];

  dps.forEach(function(dp) {
    sh.getRange(r, 1).setValue(dp[0])
      .setBackground(SEL).setFontColor(PRP).setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
    sh.getRange(r, 2).setValue(dp[1])
      .setBackground(BG).setFontColor(CMT).setFontSize(10);
    var resultCell = sh.getRange(r, 3);
    resultCell.setFormula(dp[2])
      .setBackground(BG).setFontColor(FG).setFontSize(10).setWrap(true);
    cfProceed(resultCell);
    sh.setRowHeight(r++, 40);
  });

  spacer(sh, r++);

  // Final outcome
  h2(sh, r, 1, '🏁  FINAL TRA OUTCOME');
  sh.getRange(r, 2).setBackground(SEL);
  sh.getRange(r, 3).setBackground(SEL);
  r++;

  var finalF =
    '=IF(MAX_SCORE="","⏳ TRA not yet started — complete Q1 then Q2.",' +
    'IF(MAX_SCORE<=2,"✅ MAY PROCEED — All PI is low harm risk (max score "&MAX_SCORE&"). Decision Point A1. No further assessment required.",' +
    'IF(DP_E="","⏳ TRA in progress — complete all questions Q1–Q5.",' +
    'IF(ISNUMBER(SEARCH("E1",DP_E)),"✅ MAY PROCEED — No significant risk data (Decision Point E1). Transfer may proceed with Article 46 mechanism.",' +
    'IF(ISNUMBER(SEARCH("F1",DP_F_VAL)),"✅ MAY PROCEED (CONDITIONAL) — Exception(s) apply to all significant risk data (Decision Point F1). Transfer may proceed with Article 46 mechanism, relying on identified exceptions for significant risk data.",' +
    'IF(ISNUMBER(SEARCH("F2",DP_F_VAL)),"❌ CANNOT PROCEED — Exceptions do not apply to all significant risk data (Decision Point F2). Transfer may NOT proceed. Consider removing significant risk data from scope or seek professional advice.",' +
    '"⏳ TRA in progress — complete all questions Q1–Q6."))))))';

  var finalCell = sh.getRange(r, 1);
  finalCell.setFormula(finalF)
    .setFontSize(12).setFontWeight('bold').setWrap(true).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, false, false, PRP, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.setRowHeight(r, 90);
  cfProceed(finalCell);
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  spacer(sh, r++);

  // Notes
  lbl(sh, r, 1, 'General notes and caveats:');
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  r++;
  inp(sh, r, 2);
  sh.getRange(r, 1).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r++, 65);
  spacer(sh, r++);

  sh.getRange(r, 1)
    .setValue('⚠️ This tool follows the ICO TRA Tool (November 2022). It is not legal advice. The ICO TRA Tool is one method of carrying out a TRA — other methods exist. Seek professional data protection advice to review your assessment.')
    .setBackground(BG).setFontColor(CMT).setFontSize(8).setWrap(true);
  sh.getRange(r, 2).setBackground(BG);
  sh.getRange(r, 3).setBackground(BG);
  sh.setRowHeight(r, 40);
}

// ─── KROO PI DATA ─────────────────────────────────────────────────────────────
// Returns array of:
//   [name, score(1-5), isSpecialCat, category, regulatoryContext, notes, inferenceRisk]
//
// inferenceRisk: true where the data element can INFER special category data
// even though the data itself is not Art 9 on its face.
// e.g. payments to a trade union employer, a named mosque, a cancer charity,
// an addiction clinic, or a political party can reveal trade union membership,
// religion, health, or political opinions by inference.
//
// ICO TRA Tool note: "Special category data also includes types of information
// that can be used to INFER any of the information on that list."

function getKrooPIData() {
  return [
    // ── PCI-DSS Card Data ────────────────────────────────────────────────────
    // [name, score, specialCat, category, regulatoryContext, notes, inferenceRisk]
    ['PAN (Primary Account Number / full card number)',       5, false, 'PCI-DSS Card Data',       'PCI-DSS SAQ/QSA scope; GDPR Art 32',        'Must never be stored unencrypted. Core PCI-DSS scope element.',                                    false],
    ['Card expiry date',                                      4, false, 'PCI-DSS Card Data',       'PCI-DSS',                                   'Sensitive in combination with PAN and CVV.',                                                       false],
    ['CVV / CV2 / CVC (card verification value)',             5, false, 'PCI-DSS Card Data',       'PCI-DSS',                                   'Must NEVER be stored post-authorisation. PCI-DSS prohibition.',                                    false],
    ['Card PIN / PIN block',                                  5, false, 'PCI-DSS Card Data',       'PCI-DSS',                                   'Must never be stored in any form.',                                                                false],
    ['Magnetic stripe / track data',                         5, false, 'PCI-DSS Card Data',       'PCI-DSS',                                   'Must never be stored post-authorisation.',                                                         false],
    ['Card present authentication data (chip, contactless)', 5, false, 'PCI-DSS Card Data',       'PCI-DSS',                                   'Includes cryptograms and dynamic data used during chip/contactless transactions.',                  false],

    // ── KYC / CDD / EDD ──────────────────────────────────────────────────────
    ['Government-issued photo ID (passport, driving licence)',4, false, 'KYC / CDD / EDD',         'MLR 2017; FCA SYSC 6.3',                    'Core CDD document. Contains biographic data and photo.',                                           false],
    ['Proof of address document',                             3, false, 'KYC / CDD / EDD',         'MLR 2017',                                  'Utility bill, bank statement or similar. Moderate risk in isolation.',                             false],
    ['Selfie / liveness check photo',                        4, true,  'KYC / CDD / EDD',         'GDPR Art 9 — biometric; MLR 2017',          'Processed to verify identity. Biometric data under Art 9.',                                        false],
    ['Video KYC recording',                                   4, true,  'KYC / CDD / EDD',         'GDPR Art 9 — biometric; MLR 2017',          'Full video call or automated liveness check recording.',                                           false],
    ['Source of funds declaration',                           3, false, 'KYC / CDD / EDD',         'MLR 2017',                                  'Standard CDD. May reveal employment, business interests.',                                         false],
    ['Source of wealth declaration',                          4, false, 'KYC / CDD / EDD',         'MLR 2017 — EDD',                            'Required for enhanced due diligence. Reveals financial profile.',                                  false],
    ['Enhanced due diligence (EDD) pack',                    5, false, 'KYC / CDD / EDD',         'MLR 2017',                                  'Comprehensive EDD documentation. Highly sensitive aggregate.',                                     false],
    ['KYC risk rating',                                       4, false, 'KYC / CDD / EDD',         'FCA SYSC 6.3; MLR 2017',                    'Low / medium / high / unacceptable. Drives relationship decisions.',                               false],
    ['Politically Exposed Person (PEP) status',               5, true,  'KYC / CDD / EDD',         'GDPR Art 9 — political opinions; MLR 2017 r.35', 'Directly reveals political exposure. Art 9 special category.',                              false],
    ['Sanctions screening result',                            5, false, 'KYC / CDD / EDD',         'OFSI; Financial Sanctions Act; MLR 2017',   'OFSI, OFAC, UN, EU consolidated lists.',                                                           false],
    ['Adverse media screening results',                       4, false, 'KYC / CDD / EDD',         'MLR 2017',                                  'Negative news, reputational risk flags. May reference criminal matters.',                          false],
    ['Customer due diligence (CDD) records',                  4, false, 'KYC / CDD / EDD',         'MLR 2017',                                  'Core CDD documentation package.',                                                                  false],

    // ── AML / CTF / SARs ─────────────────────────────────────────────────────
    ['Suspicious Activity Report (SAR)',                      5, false, 'AML / CTF / SARs',        'POCA 2002 s.338; Terrorism Act 2000',       'Highly sensitive. Disclosure is a criminal offence (tipping off, POCA s.333A).',                   false],
    ['SAR internal referral / MLRO referral',                5, false, 'AML / CTF / SARs',        'POCA 2002',                                 'Internal referral to MLRO before external SAR decision. Tipping off risk.',                        false],
    ['Transaction monitoring alert',                          5, false, 'AML / CTF / SARs',        'FCA SYSC 6.3; MLR 2017',                    'Output of TM system. Indicates suspected suspicious activity.',                                    false],
    ['AML risk rating',                                       5, false, 'AML / CTF / SARs',        'FCA SYSC 6.3; MLR 2017',                    'Customer-level AML risk score. Drives monitoring intensity.',                                      false],
    ['CTF risk assessment',                                   5, false, 'AML / CTF / SARs',        'Terrorism Act 2000; CTA 2008',              'Counter-terrorism financing risk flags.',                                                          false],
    ['NCA disclosure records',                                5, false, 'AML / CTF / SARs',        'POCA 2002 s.338; NCA / UKFIU',             'Records of SARs submitted to the NCA. Tipping off prohibition applies.',                           false],
    ['MLRO / DMLRO decisions and records',                   5, false, 'AML / CTF / SARs',        'POCA 2002',                                 'Money Laundering Reporting Officer and Deputy MLRO decision records.',                             false],

    // ── Financial Transactions ────────────────────────────────────────────────
    // Several transaction types carry inference risk: payments to named trade unions,
    // religious organisations, political parties, medical providers, addiction
    // services, or named charities (e.g. cancer, HIV, mental health) can reveal
    // special category data by inference under GDPR Art 9 / ICO guidance.
    ['Payment transaction records',                           4, false, 'Financial Transactions',  'PSR 2017; FCA COBS',                        '⚠️ Inference risk: merchant name / payee may reveal trade union membership, religion, health, political opinions or other Art 9 data. Treat as high risk where destination country has weak rule of law.', true],
    ['Bank transfer details (sort code, account number, payee name)', 4, false, 'Financial Transactions', 'PSR 2017',               '⚠️ Inference risk: payee names can identify trade unions (e.g. "UNITE"), charities (e.g. cancer / HIV / addiction), religious organisations, or political parties.',                              true],
    ['Direct debit records',                                  3, false, 'Financial Transactions',  'Bacs rules; PSR 2017',                      '⚠️ Inference risk: recurring DDs to trade unions, faith organisations, political parties or medical providers reveal Art 9 data. Mandate holder name is particularly sensitive.',             true],
    ['Standing order records',                                3, false, 'Financial Transactions',  'PSR 2017',                                  '⚠️ Inference risk: standing orders to named religious, political or health organisations can infer special category data.',                                                               true],
    ['International wire / SWIFT transfer details',           4, false, 'Financial Transactions',  'MLR 2017; PSR 2017',                        '⚠️ Inference risk: remittance payees / corridors may reveal ethnic origin, religion or political affiliation. SWIFT/SEPA/cross-border data.',                                            true],
    ['Merchant transaction data (MCC, merchant name, location)', 3, false, 'Financial Transactions', 'PSD2 / PSR',                             '⚠️ Inference risk: MCCs and merchant names can reveal health conditions (pharmacy, specialist clinic), religion (mosque, church donations), sexual orientation (LGBTQ+ venues), political affiliation.',  true],
    ['Refund / chargeback records',                           3, false, 'Financial Transactions',  'PSR 2017; card scheme rules',               'Dispute resolution records. Lower inference risk in isolation.',                                    false],
    ['Full account statement data',                           4, false, 'Financial Transactions',  'FCA COBS',                                  '⚠️ Inference risk: aggregate statement is highest inference risk — holistic view of all payees, MCCs and amounts reveals comprehensive Art 9 profile. Treat as high risk.',                     true],
    ['Crypto transaction records',                            5, false, 'Financial Transactions',  'MLR 2017 (cryptoasset firms); FCA reg',     'If Kroo offers crypto services. Blockchain-traceable. Wallet addresses.',                          false],

    // ── Account Data ──────────────────────────────────────────────────────────
    ['Account number and sort code',                          4, false, 'Account Data',            'PSR 2017',                                  'Core account identifier. Sufficient to receive payments; risk if misused.',                        false],
    ['IBAN / BIC',                                            4, false, 'Account Data',            'PSR 2017',                                  'International account identifiers.',                                                               false],
    ['Virtual account numbers',                               4, false, 'Account Data',            'PSR 2017',                                  'Virtual IBANs / sort codes issued by Kroo.',                                                      false],
    ['Account opening date and channel',                      2, false, 'Account Data',            'FCA COBS',                                  'Low sensitivity in isolation.',                                                                    false],
    ['Account status (active, dormant, closed, blocked)',    3, false, 'Account Data',            'FCA COBS',                                  'Blocked / frozen status may reveal AML or fraud concerns.',                                       false],
    ['Account closure reason',                                4, false, 'Account Data',            'FCA COBS',                                  'May reveal fraud, AML action, financial difficulty or customer dispute.',                          false],

    // ── Authentication / Authorisation ────────────────────────────────────────
    ['Password hash',                                         5, false, 'Auth / Authz',            'GDPR Art 32; NCSC guidance',                'bcrypt / argon2 hash. Still sensitive — brute force / rainbow table risk.',                       false],
    ['App PIN',                                               5, false, 'Auth / Authz',            'GDPR Art 32; PSD2/PSR SCA',                 'Must never be stored in plaintext.',                                                               false],
    ['One-time passcode (OTP) records',                       4, false, 'Auth / Authz',            'PSD2 / PSR SCA',                            'OTP delivery logs. Active codes are equivalent to credentials.',                                   false],
    ['MFA method and status',                                 4, false, 'Auth / Authz',            'PSD2 / PSR SCA',                            'Reveals second factor type (SMS, TOTP, biometric). Aids social engineering.',                     false],
    ['Biometric authentication data (fingerprint, Face ID)', 5, true,  'Auth / Authz',            'GDPR Art 9 — biometric; PSD2/PSR SCA',      'Template / hash of biometric. Art 9 special category.',                                          false],
    ['Device fingerprint / device ID',                        4, false, 'Auth / Authz',            'GDPR; PECR',                                'Unique device identifier. Persistent cross-session tracking capability.',                          false],
    ['Session token / refresh token',                         5, false, 'Auth / Authz',            'GDPR Art 32',                               'Active session credentials. Equivalent to username + password if intercepted.',                   false],
    ['API key / OAuth access token',                          5, false, 'Auth / Authz',            'GDPR Art 32; PSD2/PSR',                     'Third-party API access credentials. Scope of access may be broad.',                               false],
    ['Security questions and answers',                        4, false, 'Auth / Authz',            'GDPR Art 32',                               'Knowledge-based authentication. Often personally revealing (mother\'s maiden name etc.).',         false],

    // ── Open Banking / PSD2 ───────────────────────────────────────────────────
    ['Open Banking consent records',                          3, false, 'Open Banking / PSD2',     'PSD2 / PSR; Open Banking Ltd',              'Records of TPP consent grants and revocations.',                                                   false],
    ['TPP access logs',                                       3, false, 'Open Banking / PSD2',     'PSD2 / PSR',                                'Third-party provider access history.',                                                             false],
    ['PISP payment initiation data',                          4, false, 'Open Banking / PSD2',     'PSD2 / PSR',                                '⚠️ Inference risk: payment initiation payee details carry same inference risks as transaction records.', true],
    ['AISP account access history',                           3, false, 'Open Banking / PSD2',     'PSD2 / PSR',                                'Account information service provider access logs.',                                                false],
    ['Variable recurring payment (VRP) mandates',            3, false, 'Open Banking / PSD2',     'Open Banking Ltd; PSR',                     '⚠️ Inference risk: VRP payee may reveal trade union, charity, religious or political affiliation.', true],

    // ── Fraud / Disputes ──────────────────────────────────────────────────────
    ['Fraud flag / fraud score',                              5, false, 'Fraud / Disputes',        'FCA SYSC; PSR 2017',                        'Internal fraud risk rating. Highly sensitive — used in account decisions.',                       false],
    ['Fraud investigation records',                           5, false, 'Fraud / Disputes',        'FCA SYSC',                                  'Detailed fraud case records including evidence and decisions.',                                    false],
    ['Chargeback records',                                    4, false, 'Fraud / Disputes',        'PSR 2017; card scheme rules (Visa/MC)',      'Dispute resolution records.',                                                                      false],
    ['Dispute history',                                       4, false, 'Fraud / Disputes',        'PSR 2017',                                  'Payment dispute records and outcomes.',                                                            false],
    ['Velocity check results',                                4, false, 'Fraud / Disputes',        'FCA SYSC',                                  'Transaction velocity monitoring outputs.',                                                         false],
    ['Device / IP fraud signals',                             4, false, 'Fraud / Disputes',        'FCA SYSC',                                  'Device reputation and IP-based fraud detection signals.',                                         false],
    ['Account takeover investigation records',                5, false, 'Fraud / Disputes',        'FCA SYSC',                                  'ATO investigation details including attacker tactics.',                                            false],
    ['Mule account flag',                                     5, false, 'Fraud / Disputes',        'PSR 2017; POCA 2002',                       'Money mule indicator. Links to criminal activity.',                                               false],

    // ── Device / Technical Data ───────────────────────────────────────────────
    ['Device ID / IMEI',                                      3, false, 'Device / Technical',      'GDPR; PECR',                                'Unique device identifier. Persistent.',                                                            false],
    ['IP address (at transaction time)',                      3, false, 'Device / Technical',      'GDPR',                                      'Timestamped IP. Enables geolocation.',                                                             false],
    ['Geolocation data (from app)',                           4, false, 'Device / Technical',      'GDPR',                                      '⚠️ Inference risk: precise location visits to places of worship, clinics, political meetings, or trade union offices can infer Art 9 data.', true],
    ['Push notification token',                               2, false, 'Device / Technical',      'GDPR; PECR',                                'Low sensitivity in isolation.',                                                                    false],
    ['Operating system / app version',                        2, false, 'Device / Technical',      'GDPR',                                      'Low sensitivity. Used for fraud and support.',                                                     false],
    ['Browser / user agent fingerprint',                      3, false, 'Device / Technical',      'GDPR; PECR',                                'Enables cross-session tracking.',                                                                  false],
    ['Network / carrier information',                         2, false, 'Device / Technical',      'GDPR',                                      'Mobile network / carrier.',                                                                        false],
    ['Jailbreak / root detection flag',                       3, false, 'Device / Technical',      'GDPR; FCA SYSC',                            'Device security flag used in fraud assessment.',                                                   false],

    // ── Behavioural / Analytics ───────────────────────────────────────────────
    ['Spending category analytics',                           3, false, 'Behavioural / Analytics', 'GDPR Art 22',                               '⚠️ Inference risk: aggregated spending categories (health, religion, political) can infer Art 9 data. Profiling under Art 22.', true],
    ['Merchant preference data',                              2, false, 'Behavioural / Analytics', 'GDPR',                                      '⚠️ Inference risk: preferred merchants may reveal religion, health or political affiliations.', true],
    ['Budgeting / savings goal data',                         2, false, 'Behavioural / Analytics', 'GDPR',                                      'Low sensitivity unless goals reveal health or relationship context.',                               false],
    ['App usage patterns',                                    2, false, 'Behavioural / Analytics', 'GDPR',                                      'Feature usage, session frequency.',                                                                false],
    ['Predicted financial behaviour (ML model output)',       4, false, 'Behavioural / Analytics', 'GDPR Art 22; CCA',                          'Automated decision-making output. Art 22 rights apply. May embed inference.',                     false],
    ['Credit risk model inputs / outputs',                    4, false, 'Behavioural / Analytics', 'GDPR Art 22; FCA CONC; CCA',                'Creditworthiness scoring. Explainability obligations under Art 22.',                              false],

    // ── Customer Support ──────────────────────────────────────────────────────
    ['Support ticket / chat logs',                            3, false, 'Customer Support',        'GDPR; FCA DISP',                            '⚠️ Inference risk: support conversations may spontaneously reveal health, relationship or financial vulnerability — treat as potentially Art 9 if unstructured.', true],
    ['Complaint records',                                     3, false, 'Customer Support',        'FCA DISP',                                  'FCA-regulated complaints. Retention obligations apply.',                                           false],
    ['Vulnerability flag (financial difficulty, mental health, etc.)', 5, true, 'Customer Support', 'GDPR Art 9 — health; FCA Consumer Duty PS22/9', 'FCA Consumer Duty requires identification and appropriate treatment of vulnerable customers. Health / mental health data under Art 9.', false],
    ['Call recordings',                                       4, false, 'Customer Support',        'FCA COBS 11.8; MiFID II',                   'Regulatory call recording requirements. May capture sensitive disclosures.',                      false],
    ['Customer satisfaction / NPS scores',                   2, false, 'Customer Support',        'GDPR',                                      'Low sensitivity.',                                                                                 false],
    ['Escalation and case management records',               3, false, 'Customer Support',        'FCA DISP',                                  'May reflect severity of customer harm.',                                                           false],
    ['Bereavement / power of attorney records',               5, true,  'Customer Support',        'GDPR Art 9 — health; FCA Consumer Duty',    'Reveals death of account holder or loss of mental capacity. Art 9 health data.',                  false],

    // ── Regulatory / Compliance ───────────────────────────────────────────────
    ['CASS client money records',                             4, false, 'Regulatory / Compliance', 'FCA CASS',                                  'Client asset safeguarding. FCA inspection scope.',                                                false],
    ['FCA regulatory reporting data',                         4, false, 'Regulatory / Compliance', 'FCA SUP',                                   'Regulatory returns to FCA.',                                                                       false],
    ['Consumer Duty outcome monitoring data',                3, false, 'Regulatory / Compliance', 'FCA Consumer Duty PS22/9',                  'Outcome monitoring for FCA Consumer Duty compliance.',                                            false],
    ['Regulatory investigation / enforcement records',        5, false, 'Regulatory / Compliance', 'FCA; PRA',                                  'Highly sensitive. Subject to legal privilege considerations.',                                    false],
    ['HMRC / CRS / FATCA reporting data',                    4, false, 'Regulatory / Compliance', 'HMRC; IGA; CRS (OECD)',                     'International tax reporting. May be shared with overseas tax authorities.',                       false],
    ['Credit reference agency data',                          4, false, 'Regulatory / Compliance', 'CCA 1974; GDPR',                            'Experian, Equifax, TransUnion data. Third-party sourced.',                                        false],
    ['Affordability assessment records',                      4, false, 'Regulatory / Compliance', 'FCA CONC; Consumer Duty',                   'Lending affordability assessment. May include income, expenditure, debt.',                        false],
    ['Overdraft / credit decisioning records',               4, false, 'Regulatory / Compliance', 'FCA CONC; CCA',                             'Credit decision audit trail. Art 22 automated decision-making implications.',                     false],

    // ── Credit / Lending ──────────────────────────────────────────────────────
    ['Credit score (internal)',                               4, false, 'Credit / Lending',        'GDPR Art 22; CCA; FCA CONC',                'Proprietary credit score. Art 22 rights apply.',                                                  false],
    ['Credit reference agency report',                        4, false, 'Credit / Lending',        'CCA 1974; GDPR',                            'Full CRA report from Experian / Equifax / TransUnion.',                                          false],
    ['Loan / overdraft application',                          4, false, 'Credit / Lending',        'FCA CONC; CCA',                             'Application data including income, employment, expenditure.',                                     false],
    ['Repayment history',                                     4, false, 'Credit / Lending',        'CCA 1974',                                  'Repayment track record.',                                                                          false],
    ['Default / delinquency records',                         5, false, 'Credit / Lending',        'CCA 1974; FCA CONC',                        'Highly sensitive. Affects credit profile and future borrowing.',                                  false],
    ['Early repayment / settlement records',                  3, false, 'Credit / Lending',        'CCA 1974',                                  'Moderate sensitivity.',                                                                            false],
    ['BNPL / embedded finance records',                       4, false, 'Credit / Lending',        'FCA CONC; CCA (if regulated)',              'Buy Now Pay Later or embedded lending. FCA regulation applies if credit regulated.',             false],
  ];
}

// ─── APPEND KROO DATA TO LOOKUPS ──────────────────────────────────────────────
// Appends Kroo-specific PI entries to cols A–C of the Lookups tab,
// then updates PI_NAMES and PI_SCORES named ranges to include them.

function appendKrooToLookups(sh, ss) {
  var piNamesRange = ss.getRangeByName('PI_NAMES');
  var piStartRow   = piNamesRange.getRow();
  var piOrigCount  = piNamesRange.getNumRows();
  var nextRow      = piStartRow + piOrigCount;

  var krooData = getKrooPIData();

  // Section divider
  sh.getRange(nextRow, 1).setValue('── Kroo-specific data ──')
    .setBackground(SEL).setFontColor(CYN).setFontWeight('bold').setFontSize(9);
  sh.getRange(nextRow, 2).setBackground(SEL);
  sh.getRange(nextRow, 3).setBackground(SEL);
  nextRow++;

  krooData.forEach(function(row, i) {
    var r = nextRow + i;
    sh.getRange(r, 1).setValue(row[0])
      .setBackground(BG).setFontColor(FG).setFontSize(10);
    sh.getRange(r, 2).setValue(row[1])
      .setBackground(SBG[row[1]]).setFontColor(SFG[row[1]])
      .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center');
    sh.getRange(r, 3).setValue(row[2] ? 'Yes' : 'No')
      .setBackground(row[2] ? PNK : BG).setFontColor(row[2] ? BG : CMT)
      .setFontSize(9).setHorizontalAlignment('center');
  });

  // Extend named ranges to cover original + divider row + kroo data
  var newTotal = piOrigCount + 1 + krooData.length;
  nr(ss, 'PI_SCORES', sh.getRange(piStartRow, 1, newTotal, 2));
  nr(ss, 'PI_NAMES',  sh.getRange(piStartRow, 1, newTotal, 1));
}

// ─── BUILD KROO PI CATEGORIES TAB ─────────────────────────────────────────────

function buildKrooPICategories(sh) {
  // Cols: A=Category  B=Data element  C=Score  D=Special cat?
  //       E=Inference risk?  F=Regulatory context  G=Notes
  cw(sh, [[1,130],[2,195],[3,55],[4,60],[5,65],[6,155],[7,195]]);
  baseStyle(sh, 160, 7);
  sh.setFrozenRows(3);

  h1(sh, 1, 1, '🏦 Kroo — PI Data Categories, Risk Scores & Regulatory Context');
  for (let c = 2; c <= 7; c++) sh.getRange(1, c).setBackground(PRP);

  sh.getRange(2, 1).setValue(
    '⚠️ INFERENCE RISK: Transaction and location data may reveal GDPR Art 9 special category data by inference. ' +
    'e.g. payments to a named trade union, mosque/church/synagogue, cancer charity, addiction clinic, ' +
    'HIV organisation, political party, or LGBTQ+ venue can reveal trade union membership, religion, ' +
    'health, or political opinions. ICO guidance: special category data includes information from which Art 9 data CAN BE INFERRED.'
  ).setBackground(SEL).setFontColor(YLW).setFontSize(9).setWrap(true);
  for (let c = 2; c <= 7; c++) sh.getRange(2, c).setBackground(SEL);
  sh.setRowHeight(2, 52);

  // Column headers
  var hdrs = ['Category', 'Data Element', 'Score\n(1–5)', 'Special\nCat?', 'Inference\nRisk?', 'Regulatory Context', 'Notes / Examples'];
  hdrs.forEach(function(h, i) {
    sh.getRange(3, 1 + i).setValue(h)
      .setBackground(SEL).setFontColor(CYN).setFontWeight('bold')
      .setFontSize(9).setWrap(true).setHorizontalAlignment('center');
  });
  sh.setRowHeight(3, 36);

  var data    = getKrooPIData();
  var r       = 4;
  var lastCat = '';

  data.forEach(function(row) {
    var name        = row[0];
    var score       = row[1];
    var specialCat  = row[2];
    var category    = row[3];
    var regCtx      = row[4];
    var notes       = row[5];
    var inferRisk   = row[6];

    // Category section divider
    if (category !== lastCat) {
      h2(sh, r, 1, category);
      for (let c = 2; c <= 7; c++) sh.getRange(r, c).setBackground(SEL);
      r++;
      lastCat = category;
    }

    // A: category label (muted, repeated for readability)
    sh.getRange(r, 1).setValue(category)
      .setBackground(BG).setFontColor(CMT).setFontSize(8);

    // B: data element name
    sh.getRange(r, 2).setValue(name)
      .setBackground(BG).setFontColor(FG).setFontSize(9).setWrap(true);

    // C: risk score
    sh.getRange(r, 3).setValue(score)
      .setBackground(SBG[score]).setFontColor(SFG[score])
      .setFontSize(10).setFontWeight('bold').setHorizontalAlignment('center');

    // D: special category
    sh.getRange(r, 4).setValue(specialCat ? '⚠️ Yes' : 'No')
      .setBackground(specialCat ? PNK : BG)
      .setFontColor(specialCat ? BG : CMT)
      .setFontSize(9).setHorizontalAlignment('center');

    // E: inference risk
    sh.getRange(r, 5).setValue(inferRisk ? '⚠️ Yes' : 'No')
      .setBackground(inferRisk ? ORG : BG)
      .setFontColor(inferRisk ? BG : CMT)
      .setFontSize(9).setHorizontalAlignment('center');

    // F: regulatory context
    sh.getRange(r, 6).setValue(regCtx)
      .setBackground(BG).setFontColor(CMT).setFontSize(8).setWrap(true);

    // G: notes
    sh.getRange(r, 7).setValue(notes)
      .setBackground(inferRisk ? '#3d3320' : BG)
      .setFontColor(inferRisk ? YLW : FG)
      .setFontSize(8).setWrap(true);

    sh.setRowHeight(r, inferRisk ? 50 : 34);
    r++;
  });

  spacer(sh, r++);

  // Score key
  h2(sh, r, 1, 'Score key');
  for (let c = 2; c <= 7; c++) sh.getRange(r, c).setBackground(SEL);
  r++;
  for (let s = 1; s <= 5; s++) {
    sh.getRange(r, 1).setValue(SLB[s])
      .setBackground(SBG[s]).setFontColor(SFG[s]).setFontSize(9).setFontWeight('bold');
    for (let c = 2; c <= 7; c++) sh.getRange(r, c).setBackground(SBG[s]);
    r++;
  }

  spacer(sh, r++);

  // Legend
  h2(sh, r, 1, 'Column legend');
  for (let c = 2; c <= 7; c++) sh.getRange(r, c).setBackground(SEL);
  r++;
  var legend = [
    [PNK,  BG,  'Special Cat? = ⚠️ Yes', 'GDPR Art 9 special category data — requires explicit legal basis and extra protections.'],
    [ORG,  BG,  'Inference Risk? = ⚠️ Yes', 'Data is NOT itself Art 9 but can INFER Art 9 data (trade union, religion, health, political opinions etc.) from payee names, merchant names, or location. Treat score as +1 when destination country has weak rule of law or human rights concerns.'],
    ['#3d3320', YLW, 'Yellow notes row', 'Notes row is highlighted to flag inference risk detail.'],
  ];
  legend.forEach(function(lg) {
    sh.getRange(r, 1).setValue(lg[2]).setBackground(lg[0]).setFontColor(lg[1]).setFontSize(9).setFontWeight('bold').setWrap(true);
    sh.getRange(r, 2).setValue(lg[3]).setBackground(BG).setFontColor(CMT).setFontSize(8).setWrap(true);
    for (let c = 3; c <= 7; c++) sh.getRange(r, c).setBackground(BG);
    sh.setRowHeight(r++, 40);
  });
}
