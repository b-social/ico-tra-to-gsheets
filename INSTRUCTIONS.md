# Using the ICO TRA Tool

Work through tabs in order: **Overview → Q0 → 📋 Instructions → Q1 → Q2 → Q3 → Q4 → Q5 → Q6**.
The ✅ Summary tab auto-calculates throughout.

> Read the **📋 Instructions** tab first — it contains the score key, colour guide and decision point reference.

---

## Q0 — About this Assessment

Fill in project details before starting:

- Name(s) of assessor
- Department / Team
- Project Name / Proposed outsourcer
- Country of outsourcer
- Link to onboarding request (OneTrust / Jira)
- Criticality of proposed supplier

---

## Q1 — Transfer Details

Document the full circumstances of the transfer:

- Importer name, destination country, status (controller / processor etc.)
- Organisation type (dropdown)
- Categories of people (checkboxes)
- Volume, frequency and duration
- Technical and organisational measures (TOMs) — exporter and importer
- Format and transfer process

---

## Q2 — PI Risk Scores

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

---

## Q3 — Investigation Level

Select your organisation size and transfer volume. The **recommended investigation level** auto-calculates using the matrix from the ICO tool.

**Confirm** your chosen level at Decision Point B (dropdown). Add your reasons and resources used.

| Level | When required |
|-------|--------------|
| 1 | SME, moderate risk data |
| 2 | Large org moderate risk; or SME high risk low volume |
| 3(i) | Full detailed investigation — professional advice likely needed |
| 3(ii) | All high-score (≥4) PI treated as significant risk data; Level 2 for the rest |

---

## Q4 — Human Rights Risk

Based on your investigation:

1. Answer **Key Question 1**: concerns about human rights in the destination country?
2. If yes, answer **Key Question 2**: does the transfer significantly increase the risk?

**Decision Point C** auto-calculates (C1–C4).

---

## Q5 — Enforcement

Work through the enforcement questionnaire (EQ1–EQ4).

**Decision Point D** auto-calculates from your answers.

**Decision Point E** — select whether you have identified any **significant risk data** (human rights risk data and/or enforceability risk data). If E1 (none identified), the transfer may proceed.

> If you selected Level 3 Option (ii) at Q3, skip Q5 entirely.

---

## Q6 — Exceptions

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

ICO initial scores are pre-loaded and auto-looked-up. Adjust them based on aggravating / mitigating factors.

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

## Source document

`transfer-risk-assessments-tool-20221117.doc` — ICO TRA Tool, November 2022.

This tool follows that structure exactly. It is **not legal advice**. The ICO TRA Tool is one method of carrying out a TRA — other methods exist (including the EDPB approach). Seek professional data protection advice to review your assessment.
