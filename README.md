# ICO Transfer Risk Assessment (TRA) Tool — Google Sheets

Google Apps Script that builds a complete, interactive TRA workbook inside Google Sheets.

Based on the **ICO TRA Tool (November 2022)** for UK GDPR Article 46 restricted transfers.
Includes a **Kroo-specific PI data tab** covering PCI-DSS, KYC/AML/CTF, Open Banking, fraud, and more.

---

## What it builds

| Tab | Purpose |
|-----|---------|
| 📋 Instructions | How to use the tool, score key, colour guide |
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
5. Click **Save** (or `Cmd/Ctrl + S`)
6. In the function dropdown at the top, select **`buildTRATool`**
7. Click **▶ Run**
8. When prompted, click **Review permissions → Allow**
9. Switch back to your spreadsheet — all tabs will have been created

> The script takes ~30–60 seconds to run. A confirmation dialog appears when complete.

---

## Usage

Work through the tabs **in order, Q1 → Q6**. The ✅ Summary tab auto-calculates throughout.

### Q1 — Transfer Details

Document the full circumstances of the transfer:

- Importer name, destination country, status (controller / processor etc.)
- Organisation type (dropdown)
- Categories of people (checkboxes)
- Volume, frequency and duration
- Technical and organisational measures (TOMs) — exporter and importer
- Format and transfer process

### Q2 — PI Risk Scores

For each category of personal information being transferred:

1. Select the category from the **dropdown in column B** (Kroo-specific categories are included)
2. **ICO Initial Score** auto-populates from the Lookups tab (VLOOKUP)
3. **Special Category?** auto-flags Art 9 data
4. **Suggested Score** = ICO score, +1 if special category (capped at 5)
5. Note any aggravating or mitigating factors
6. Set your **Final Score** (1–5) in column H

**Decision Point A** calculates automatically from the max score:

| Max score | Outcome |
|-----------|---------|
| ≤ 2 | ✅ Low harm risk — MAY PROCEED |
| 3 | ⚠️ Moderate — go to Q3 |
| ≥ 4 | 🔴 High — go to Q3 |

### Q3 — Investigation Level

Select your organisation size and transfer volume. The **recommended investigation level** auto-calculates using the matrix from the ICO tool.

**Confirm** your chosen level at Decision Point B (dropdown). Add your reasons and resources used.

| Level | When required |
|-------|--------------|
| 1 | SME, moderate risk data |
| 2 | Large org moderate risk; or SME high risk low volume |
| 3(i) | Full detailed investigation — professional advice likely needed |
| 3(ii) | All high-score (≥4) PI treated as significant risk data; Level 2 for the rest |

### Q4 — Human Rights Risk

Based on your investigation:

1. Answer **Key Question 1**: concerns about human rights in the destination country?
2. If yes, answer **Key Question 2**: does the transfer significantly increase the risk?

**Decision Point C** auto-calculates (C1–C4).

### Q5 — Enforcement

Work through the enforcement questionnaire (EQ1–EQ4).

**Decision Point D** auto-calculates from your answers.

**Decision Point E** — select whether you have identified any **significant risk data** (human rights risk data and/or enforceability risk data). If E1 (none identified), the transfer may proceed.

> If you selected Level 3 Option (ii) at Q3, skip Q5 entirely.

### Q6 — Exceptions

Only required if significant risk data was identified at Decision Point E.

For each of the 7 exceptions, indicate whether it applies, which data it covers, and whether the benefit outweighs the risk. **Decision Point F** is a manual selection (F1 = may proceed / F2 = cannot proceed).

---

## Scoring

Risk scores follow a 1–5 scale:

| Score | Label | Colour | Typical meaning |
|-------|-------|--------|----------------|
| 1 | Very Low | 🟢 Green | Unlikely to cause any harm |
| 2 | Low | 🔵 Cyan | Inconsequential harm at most |
| 3 | Moderate | 🟡 Yellow | Minor harm; some remediation needed |
| 4 | High | 🟠 Orange | Significant harm; urgent action needed |
| 5 | Very High | 🔴 Red | Severe harm; criminal convictions, biometric, AML data etc. |

ICO initial scores are pre-loaded and auto-looked-up. You adjust them based on aggravating / mitigating factors.

---

## Kroo PI Categories tab

The 🏦 Kroo PI Categories tab covers data specific to Kroo as a UK-regulated Neobank:

| Category | Examples |
|----------|---------|
| PCI-DSS Card Data | PAN, CVV, PIN, track data |
| KYC / CDD / EDD | Photo ID, liveness checks, PEP status, sanctions screening |
| AML / CTF / SARs | SARs, TM alerts, AML risk ratings, NCA disclosures |
| Financial Transactions | Payments, DDs, standing orders, SWIFT transfers, statements |
| Account Data | Sort codes, IBANs, account status |
| Auth / Authz | Password hashes, OTPs, biometric auth, session tokens |
| Open Banking / PSD2 | Consent records, TPP access logs, PISP/AISP data |
| Fraud / Disputes | Fraud scores, chargeback records, mule flags |
| Device / Technical | Device IDs, IP addresses, geolocation |
| Behavioural / Analytics | Spending analytics, ML model outputs |
| Customer Support | Chat logs, complaints, vulnerability flags, call recordings |
| Regulatory / Compliance | CASS, FCA reporting, CRS/FATCA, CRA data |
| Credit / Lending | Credit scores, loan applications, default records |

### ⚠️ Inference risk

Several transaction and location data types are flagged with **Inference Risk** (orange column).

Even though these are not themselves Art 9 special category data, they can **reveal** Art 9 data by inference — for example:

- Payments to a named **trade union** → reveals trade union membership
- Payments to a **mosque, church, synagogue, or temple** → reveals religion
- Payments to a **cancer charity, HIV organisation, or addiction clinic** → reveals health condition
- Payments to a **political party or campaign** → reveals political opinions
- Location visits to a **place of worship or political meeting** → reveals religion / political opinions
- Merchant data at **LGBTQ+ venues** → reveals sexual orientation

The ICO TRA Tool states: *"special category data also includes types of information that can be used to **infer** any of the information on that list."*

Where inference risk is flagged, consider increasing the score by 1 if the destination country has weak rule of law, human rights concerns, or limited data protection enforcement.

All Kroo-specific categories are added to the Lookups tab automatically and appear in the Q2 dropdown.

---

## Customisation

### Adding or editing PI categories

Edit the `piData` array in `buildLookups()` or the `getKrooPIData()` array for Kroo-specific data. Each entry is:

```javascript
['Category name', score /* 1-5 */, isSpecialCategory /* true/false */]
```

For Kroo data, the full format is:

```javascript
['Name', score, isSpecialCat, 'Category group', 'Regulatory context', 'Notes', inferenceRisk /* true/false */]
```

### Adding dropdown options

Edit the `L(...)` calls in `buildLookups()`. Each named range maps to a dropdown used on the working sheets.

### Re-running / rebuilding

You can re-run `buildTRATool()` at any time. **All tabs are cleared and rebuilt** — any data you have entered will be lost. Keep a copy before re-running.

---

## Design decisions

- **No merged cells** — every value is in a single cell; easier to reference and less fragile
- **Dracula colour theme** — high-contrast dark theme throughout
- **All dropdowns reference named ranges on the Lookups tab** — change options in one place
- **Decision Points auto-calculate** — formulas reference named ranges cross-sheet; no manual transcription
- **Max score drives Decision Point A** — one high-risk category is enough to trigger a full investigation
- **Scores 1–5** rather than Low/Moderate/High — finer granularity, consistent with the suggested-score formula (+1 for special category data)

---

## Source document

`transfer-risk-assessments-tool-20221117.doc` — ICO TRA Tool, November 2022.

This tool follows that structure exactly. It is **not legal advice**. The ICO TRA Tool is one method of carrying out a TRA — other methods exist (including the EDPB approach). Seek professional data protection advice to review your assessment.

---

## Files

```
ico-tra/
├── build_tra_tool.gs                    # Google Apps Script — run this
├── transfer-risk-assessments-tool-20221117.doc   # Source ICO document
└── README.md                            # This file
```
