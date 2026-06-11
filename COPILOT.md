# COPILOT.md — Session Record

Planning and decision log for the ICO TRA Tool / Kroo PI data project.

---

## Session: 2026-06-11

### What was asked

Convert the ICO TRA Tool (November 2022, 41-page Word doc) into an interactive spreadsheet with:
- Each section on its own tab
- Dropdowns, formulae, cross-references
- Automated scoring and handling
- User-friendly and steering

### Format decision

**Google Sheets via Google Apps Script** — chosen over Excel because it's easier to distribute, collaborate on, and run without local installs. Single `.gs` file, paste into Extensions → Apps Script, run `buildTRATool()`.

---

## Design decisions

### No merged cells
Requested explicitly. Every value in a single cell. Avoids fragility, makes cross-sheet referencing reliable.

### Risk scores 1–5
Requested explicitly. `1 – Very Low` through `5 – Very High`. Not Low/Moderate/High strings. Pre-filled ICO scores from reference doc mapped to numeric scale:
- ICO "Low" → 2
- ICO "Moderate" → 3
- ICO "High" → 4

Score 1 and 5 available for user to adjust down/up based on mitigating/aggravating factors.

### Decision Point A thresholds
Agreed in session:
- Max score ≤ 2 → MAY PROCEED (low harm risk)
- Max score = 3 → moderate, go to Q3
- Max score ≥ 4 → high, go to Q3

Driven by **MAX** score across all PI categories (not average, not count).

### Column width cap: 200px
Requested. No "super wide" columns.

### Dropdowns from named ranges on Lookups tab
Requested — "dropdowns use ranges from a tab". All option lists live on the `Lookups` tab as named ranges. Working sheets use `ddNamed()` which calls `requireValueInRange()`. Single source of truth — change an option once on Lookups, all dropdowns update.

### Single Lookups tab
Chosen over separate ref tabs per section. Cleaner. All dropdown lists and PI category scores in one place.

---

## Colour theme

Dracula throughout. Constants at top of `.gs`:

```javascript
var BG  = '#282a36', SEL = '#44475a', FG  = '#f8f8f2', CMT = '#6272a4';
var CYN = '#8be9fd', GRN = '#50fa7b', ORG = '#ffb86c', PNK = '#ff79c6';
var PRP = '#bd93f9', RED = '#ff5555', YLW = '#f1fa8c', BLK = '#21222c';
```

Score colours: green (1) → cyan (2) → yellow (3) → orange (4) → red (5).

---

## Sheet structure

| Tab | What it does |
|-----|-------------|
| 📋 Instructions | Intro, score key, colour guide, navigation |
| Q1 · Transfer Details | Importer details, people categories (checkboxes), volume, duration, TOMs |
| Q2 · PI Risk Scores | 15 PI category rows; ICO score VLOOKUP; special cat flag; suggested score; aggravating/mitigating notes; final score dropdown; Decision Point A |
| Q3 · Investigation Level | Business size + volume inputs; investigation matrix (read-only ref); auto-recommended level; Decision Point B (user selects) |
| Q4 · Human Rights Risk | ECHR rights reference; KQ1/KQ2 dropdowns; Decision Point C (auto) |
| Q5 · Enforcement | EQ1–EQ4 questionnaire; factor checkboxes for EQ3; Decision Point D (auto); Decision Point E (user selects) |
| Q6 · Exceptions | 7 exceptions; applies/data/benefit dropdowns; Decision Point F (user selects) |
| ✅ Summary | All decision points (auto); final outcome (auto); notes; disclaimer |
| 🏦 Kroo PI Categories | Kroo-specific reference tab (see below) |
| Lookups | PI names + scores; special cat flag; all dropdown lists as named ranges |

---

## Kroo-specific PI tab

Requested: add a tab for bank customer data relevant to a UK Neobank (Kroo). Decision: call it "Kroo" not "NeoBank" (user correction).

### Categories included (all agreed, "add everything")

- PCI-DSS Card Data (PAN, CVV, PIN, track data, card present auth)
- KYC / CDD / EDD (photo ID, liveness checks, PEP status, sanctions, adverse media, EDD packs)
- AML / CTF / SARs (SARs, MLRO referrals, TM alerts, AML risk ratings, NCA disclosures)
- Financial Transactions (payments, DDs, SOs, SWIFT, merchant data, statements, crypto)
- Account Data (sort codes, IBANs, virtual accounts, account status/closure)
- Auth / Authz (password hashes, PINs, OTPs, MFA, biometric auth, session tokens, API keys)
- Open Banking / PSD2 (consent records, TPP logs, PISP/AISP data, VRP mandates)
- Fraud / Disputes (fraud scores, chargeback records, mule flags, ATO records)
- Device / Technical (device IDs, IP addresses, geolocation, jailbreak detection)
- Behavioural / Analytics (spending analytics, ML outputs, credit risk models)
- Customer Support (chat logs, complaints, vulnerability flags, call recordings, bereavement/POA)
- Regulatory / Compliance (CASS, FCA reporting, CRS/FATCA, CRA data, Consumer Duty)
- Credit / Lending (credit scores, CRA reports, loan applications, default records, BNPL)

### Pre-scored: yes (editable)

Agreed: pre-fill ICO-style 1–5 scores, user can edit.

### Lives in both places: Kroo tab + Lookups

Agreed: reference/browsing tab (`🏦 Kroo PI Categories`) AND appended to Lookups so entries appear in Q2 dropdowns.

### Inference risk flag

User explicitly raised: *"note some of the transaction details may hit inference — trades union payments, charities etc."*

Implemented as `inferenceRisk` boolean (7th element in data array). When `true`:
- Orange flag in `Inference Risk?` column of Kroo tab
- Notes cell highlighted yellow-on-dark
- Notes text includes explicit examples of what can be inferred

**Legal basis (ICO TRA Tool):** *"Special category data also includes types of information that can be used to infer any of the information on that list."*

Concrete examples documented:
- Payments to named **trade unions** → trade union membership (Art 9)
- Payments to **mosques, churches, synagogues** → religion/belief (Art 9)
- Payments to **cancer charities, HIV organisations, addiction clinics** → health (Art 9)
- Payments to **political parties** → political opinions (Art 9)
- **Merchant data** at LGBTQ+ venues → sexual orientation (Art 9)
- **Geolocation** near places of worship / political meetings → religion / political opinions (Art 9)
- **Standing orders / direct debits** to trade union employers → trade union membership (Art 9)
- **Full account statement** → highest aggregate inference risk, holistic Art 9 profile

---

## Files

| File | Purpose |
|------|---------|
| `build_tra_tool.gs` | Google Apps Script — the whole tool |
| `transfer-risk-assessments-tool-20221117.doc` | Source ICO document (antiword used to extract content) |
| `README.md` | Setup and usage instructions |
| `COPILOT.md` | This file — session record |
| `.github/copilot-instructions.md` | Copilot auto-loaded instructions for future sessions |

---

## Outstanding / future work

- [ ] Q2 currently supports 15 PI category rows — could be made dynamic (insert rows button via script)
- [ ] No print / PDF export formatting — could add a `printView()` function that reformats for A4
- [ ] Summary tab could output a formatted narrative TRA statement ready to sign off
- [ ] Inference risk scoring: consider adding +1 to suggested score automatically when `inferenceRisk=true` and destination country has known human rights concerns (would need a country risk lookup)
- [ ] Country risk lookup tab — pre-scored destination countries by human rights / rule of law rating (FCDO / Freedom House data)
- [ ] Version / audit trail — track when the TRA was last updated and by whom
