# ICO Transfer Risk Assessment (TRA) Tool — Google Sheets

Google Apps Script that builds a complete, interactive TRA workbook inside Google Sheets.

Based on the **ICO TRA Tool (November 2022)** for UK GDPR Article 46 restricted transfers.
Includes a **Kroo-specific PI data tab** covering PCI-DSS, KYC/AML/CTF, Open Banking, fraud, and more.

→ For how to **use the sheet once built**, see [INSTRUCTIONS.md](./INSTRUCTIONS.md).

---

## What it builds

| Tab | Purpose |
|-----|---------|
| Overview | Template description and maintainer info |
| Q0 · About | Project details (assessor, dept, project, country) |
| 📋 Instructions | Score key and colour guide |
| Q1 · Transfer Details | Importer, people, duration, protections, PI categories |
| Q2 · PI Risk Scores | Score each PI category 1–5; Decision Point A |
| Q3 · Investigation Level | Investigation matrix; Decision Point B |
| Q4 · Human Rights Risk | Human rights assessment; Decision Point C |
| Q5 · Enforcement | Enforcement questionnaire; Decision Points D & E |
| Q6 · Exceptions | Exceptions checklist; Decision Point F |
| ✅ Summary | All decision points + final TRA outcome (auto-calculated) |
| 🏦 Kroo PI Categories | Kroo-specific PI data, scores, regulatory context, inference risk flags |
| Lookups | All dropdown option lists and PI category scores (source of truth) |

---

## Setup

### Prerequisites

- A Google account
- A Google Sheet (new or existing)

### Steps

1. Open (or create) a Google Sheet at [sheets.google.com](https://sheets.google.com)
2. Go to **Extensions → Apps Script**
3. Delete any placeholder code in the editor
4. Copy the entire contents of [`build_tra_tool.gs`](./build_tra_tool.gs) and paste it in
5. Click **Save** (`Cmd/Ctrl + S`)
6. In the function dropdown, select **`buildTRATool`**
7. Click **▶ Run**
8. When prompted, click **Review permissions → Allow**
9. Switch back to your spreadsheet — all tabs will have been created

> The script takes ~30–60 seconds. A confirmation dialog appears when complete.

---

## Development

### Repository

[github.com/b-social/ico-tra-to-gsheets](https://github.com/b-social/ico-tra-to-gsheets)

Use [GitHub Flow](https://docs.github.com/en/get-started/using-github/github-flow): work on a branch, open a PR, merge to `main`.

```bash
git checkout -b your-branch
# make changes
git commit --no-gpg-sign -m "feat: description"
gh pr create
```

### Customising PI categories

Edit the `piData` array in `buildLookups()` or `getKrooPIData()` for Kroo-specific data:

```javascript
// Standard format
['Category name', score /* 1-5 */, isSpecialCategory /* true/false */]

// Kroo format
['Name', score, isSpecialCat, 'Category group', 'Regulatory context', 'Notes', inferenceRisk /* true/false */]
```

### Adding dropdown options

Edit the `L(...)` calls in `buildLookups()`. Each named range maps to a dropdown on the working sheets. All dropdowns must reference named ranges on the Lookups tab — never use hardcoded `requireValueInList` arrays.

### Re-running / rebuilding

`buildTRATool()` clears and rebuilds all tabs. **All entered data is lost.** Keep a copy before re-running.

### Sheet comparison (compare_sheets.py)

Downloads live and reference sheets as XLSX, diffs cell values, merges and column widths per tab.

```bash
source .venv/bin/activate
python compare_sheets.py --ref-id SHEET_ID --no-values
```

Requires `gcloud auth login --enable-gdrive-access`.

### Maintenance scripts

- `maintenance.gs` — `removeSpacerRows()`, `removeSpacerRowsActiveSheet()`, `auditTRATool()`
- `remove_spacers.gs` — standalone spacer row removal

---

## Design decisions

- **No merged cells** — every value is in a single cell; easier to reference and less fragile
- **Dracula colour theme** — high-contrast dark theme throughout
- **All dropdowns reference named ranges on the Lookups tab** — change options in one place
- **Decision Points auto-calculate** — formulas reference named ranges cross-sheet
- **Max score drives Decision Point A** — one high-risk category triggers a full investigation
- **Scores 1–5** rather than Low/Moderate/High — finer granularity, consistent with the +1 special category formula

---

## Files

```
ico-tra/
├── .github/
│   └── copilot-instructions.md     # Copilot context (auto-loaded)
├── build_tra_tool.gs               # Main build script — paste into Apps Script
├── maintenance.gs                  # Spacer removal and audit helpers
├── remove_spacers.gs               # Standalone spacer removal script
├── compare_sheets.py               # XLSX diff tool (live vs reference sheet)
├── diff_sheet.py                   # Sheets API structural diff
├── transfer-risk-assessments-tool-20221117.doc  # Source ICO document
├── INSTRUCTIONS.md                 # End-user guide for completing a TRA
├── LICENSE                         # MIT
└── README.md                       # This file
```

---

## License

MIT © Kroo Bank Ltd. — see [LICENSE](./LICENSE).

This tool is **not legal advice**. The ICO TRA Tool is one method of carrying out a TRA — other methods exist (including the EDPB approach). Seek professional data protection advice to review your assessment.
