# Copilot Instructions — ico-tra

## What this repo is

Google Apps Script (`build_tra_tool.gs`) that builds an interactive ICO Transfer Risk Assessment (TRA) workbook in Google Sheets. Based on the ICO TRA Tool (November 2022) for UK GDPR Article 46 restricted transfers. Includes a Kroo-specific PI data tab.

## Key constraints (always apply)

- **No merged cells** — ever. Every value is in a single cell.
- **Risk scores 1–5** (1 Very Low → 5 Very High). Never Low/Moderate/High strings for scores.
- **Dracula colour theme** throughout. Palette constants at top of `.gs` file: `BG`, `SEL`, `FG`, `CMT`, `CYN`, `GRN`, `ORG`, `PNK`, `PRP`, `RED`, `YLW`, `BLK`.
- **Column widths capped at 200px**.
- **All dropdowns reference named ranges on the Lookups tab** — never hardcoded `requireValueInList` arrays on working sheets. Use `ddNamed(sh, row, col, 'NAMED_RANGE')`.
- **Decision Point A threshold**: max score ≤2 → proceed; =3 → moderate/Q3; ≥4 → high/Q3.
- **Decision Point A is driven by MAX score** across all PI categories in Q2.

## Architecture

### Named ranges (spreadsheet-scoped, cross-sheet formulas)

| Name | Set in | Used in |
|------|--------|---------|
| `PI_NAMES` | Lookups | Q2 col B dropdown |
| `PI_SCORES` | Lookups | Q2 col C VLOOKUP |
| `MAX_SCORE` | Q2 | Q3, Summary |
| `BIZ_SIZE` | Q3 | Q3 formula |
| `XFER_VOL` | Q3 | Q3 formula |
| `DP_B` | Q3 | Q4, Summary |
| `KQ4_1` | Q4 | Q4 DP-C formula, Summary |
| `KQ4_2` | Q4 | Q4 DP-C formula, Summary |
| `EQ5_1–4` | Q5 | Q5 DP-D formula, Summary |
| `DP_E` | Q5 | Q6, Summary |
| `DP_F_VAL` | Q6 | Summary |

### Lookup tab structure

- **Col A**: PI category name (source for `PI_NAMES`)
- **Col B**: ICO initial score 1–5 (source for `PI_SCORES`)
- **Col C**: Special category? (Yes/No) — used by Q2 col D formula
- **Col 4+**: All dropdown option lists — one per named range, written by `list()` helper

Kroo-specific PI data is appended to the Lookups tab by `appendKrooToLookups()`, called immediately after `buildLookups()`. It extends `PI_NAMES` and `PI_SCORES` named ranges.

### Helper functions

| Function | Purpose |
|----------|---------|
| `h1/h2/h3(sh, r, c, txt)` | Heading styles (purple/teal/muted) |
| `lbl(sh, r, c, txt)` | Italic muted label |
| `inp(sh, r, c)` | Dark bordered input cell |
| `auto(sh, r, c, formula)` | Auto-calculated cell (very dark bg, cyan text) |
| `dpCell(sh, r, c, formula)` | Decision point result cell (prominent, purple border) |
| `ddNamed(sh, r, c, namedRangeName)` | Dropdown backed by Lookups named range |
| `ck(sh, r, c)` | Checkbox |
| `nr(ss, name, range)` | Set/overwrite named range |
| `cfScore(range)` | Conditional format 1–5 score colours |
| `cfProceed(range)` | Conditional format MAY PROCEED/CANNOT/CONDITIONAL/⏳ |
| `spacer(sh, r)` | 6px spacer row |
| `cw(sh, [[col,width],...])` | Set column widths |

## Kroo-specific data

- Tab: `🏦 Kroo PI Categories`
- Data function: `getKrooPIData()` — returns array of `[name, score, isSpecialCat, category, regulatoryContext, notes, inferenceRisk]`
- Appended to Lookups by: `appendKrooToLookups(sh, ss)`
- Tab built by: `buildKrooPICategories(sh)`

### Inference risk flag

Many financial transaction data types carry **inference risk** — they are not themselves GDPR Art 9 special category data but can reveal Art 9 data by inference. Examples:

- Payments to named trade unions → trade union membership
- Payments to named religious organisations → religion/belief
- Payments to cancer charities, HIV organisations, addiction clinics → health
- Payments to political parties → political opinions
- Location data near places of worship / clinics / political meetings → religion/health/political opinions
- Merchant data at LGBTQ+ venues → sexual orientation

The `inferenceRisk` boolean (7th element) controls orange highlighting in the Kroo tab. When flagged, the notes column is highlighted yellow-on-dark-background.

ICO basis: *"Special category data also includes types of information that can be used to infer any of the information on that list."*

## Re-running the script

Running `buildTRATool()` clears and rebuilds all tabs. All entered data is lost. Warn user before re-running.

## File structure

```
ico-tra/
├── .github/
│   └── copilot-instructions.md   ← this file (auto-loaded by Copilot)
├── build_tra_tool.gs             ← Google Apps Script, paste into Apps Script editor
├── transfer-risk-assessments-tool-20221117.doc  ← source ICO document
├── COPILOT.md                    ← human-readable session record
└── README.md                     ← setup and usage instructions
```
