"""
Diff the live Google Sheet against the structure expected by build_tra_tool.gs.

Usage:
    python diff_sheet.py

Requires:
    gcloud auth login --enable-gdrive-access  (already done)
"""

import json
import subprocess
import sys

import googleapiclient.discovery

SPREADSHEET_ID = "1hrlIC5yuqntgBaLRwN57m_uJHin7AY1i0AYrZgkV0Vs"

EXPECTED_TABS = [
    "📋 Instructions",
    "Q1 · Transfer Details",
    "Q2 · PI Risk Scores",
    "Q3 · Investigation Level",
    "Q4 · Human Rights Risk",
    "Q5 · Enforcement",
    "Q6 · Exceptions",
    "✅ Summary",
    "🏦 Kroo PI Categories",
    "Lookups",
]

EXPECTED_NAMED_RANGES = [
    "PI_NAMES", "PI_SCORES", "MAX_SCORE",
    "BIZ_SIZE", "XFER_VOL", "DP_B",
    "KQ4_1", "KQ4_2",
    "EQ5_1", "EQ5_2", "EQ5_3", "EQ5_4",
    "DP_E", "DP_F_VAL",
    "DD_SCORES", "DD_IMPORTER_STATUS", "DD_ORG_TYPE", "DD_VULN",
    "DD_FREQUENCY", "DD_BIZ_SIZE", "DD_XFER_VOL", "DD_INV_LEVEL",
    "DD_KQ4_1", "DD_KQ4_2",
    "DD_EQ5_1", "DD_EQ5_2", "DD_EQ5_3", "DD_EQ5_4",
    "DD_DP_E", "DD_EXCEPTION_YN", "DD_BENEFIT_YN", "DD_DP_F",
]

# Expected column widths (pixels) per sheet — from build_tra_tool.gs cw() calls
EXPECTED_COL_WIDTHS = {
    "Q1 · Transfer Details":    {1: 160, 2: 200, 3: 160},
    "Q2 · PI Risk Scores":      {1: 28, 2: 200, 3: 75, 4: 65, 5: 75, 6: 160, 7: 160, 8: 80, 9: 160},
    "Q3 · Investigation Level": {1: 160, 2: 200, 3: 160, 4: 160},
    "Q4 · Human Rights Risk":   {1: 180, 2: 200, 3: 160},
    "Q5 · Enforcement":         {1: 200, 2: 200, 3: 160},
    "Q6 · Exceptions":          {1: 130, 2: 200, 3: 80, 4: 140, 5: 100, 6: 180},
    "✅ Summary":                {1: 160, 2: 200, 3: 160},
}


def get_access_token():
    result = subprocess.run(
        ["gcloud", "auth", "print-access-token"],
        capture_output=True, text=True, check=True
    )
    return result.stdout.strip()


def build_service(token):
    from google.oauth2.credentials import Credentials
    creds = Credentials(token=token)
    return googleapiclient.discovery.build("sheets", "v4", credentials=creds)


def col_letter(n):
    """1-based column index to letter(s)."""
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def main():
    print("🔑 Getting access token...")
    token = get_access_token()
    service = build_service(token)
    sheets_api = service.spreadsheets()

    print(f"📥 Fetching spreadsheet metadata...")
    meta = sheets_api.get(
        spreadsheetId=SPREADSHEET_ID,
        includeGridData=False,
        fields="sheets(properties,merges),namedRanges"
    ).execute()

    issues = []
    info = []

    # ── 1. Tabs ───────────────────────────────────────────────────────────────
    live_tabs = {s["properties"]["title"]: s for s in meta.get("sheets", [])}
    print(f"\n📋 Tabs: {len(live_tabs)} found, {len(EXPECTED_TABS)} expected")

    missing_tabs = [t for t in EXPECTED_TABS if t not in live_tabs]
    extra_tabs   = [t for t in live_tabs if t not in EXPECTED_TABS]

    for t in missing_tabs:
        issues.append(f"❌ Missing tab: \"{t}\"")
    for t in extra_tabs:
        info.append(f"ℹ️  Extra tab (manual addition?): \"{t}\"")

    # ── 2. Named ranges ───────────────────────────────────────────────────────
    live_nr = {nr["name"]: nr for nr in meta.get("namedRanges", [])}
    print(f"🏷  Named ranges: {len(live_nr)} found, {len(EXPECTED_NAMED_RANGES)} expected")

    for name in EXPECTED_NAMED_RANGES:
        if name not in live_nr:
            issues.append(f"❌ Missing named range: {name}")

    extra_nr = [n for n in live_nr if n not in EXPECTED_NAMED_RANGES]
    for n in extra_nr:
        info.append(f"ℹ️  Extra named range: {n}")

    # ── 3. Merged cells (new — not in original build script) ─────────────────
    print("🔀 Checking merged cells...")
    for sheet_data in meta.get("sheets", []):
        title = sheet_data["properties"]["title"]
        merges = sheet_data.get("merges", [])
        if merges:
            for m in merges:
                r1 = m["startRowIndex"] + 1
                c1 = col_letter(m["startColumnIndex"] + 1)
                c2 = col_letter(m["endColumnIndex"])
                info.append(f"🔀 Merged cells in \"{title}\": row {r1}, cols {c1}–{c2}")

    # ── 4. Column widths ──────────────────────────────────────────────────────
    print("📐 Fetching column widths...")
    for sheet_name, expected_widths in EXPECTED_COL_WIDTHS.items():
        if sheet_name not in live_tabs:
            continue
        sheet_id = live_tabs[sheet_name]["properties"]["sheetId"]
        sheet_detail = sheets_api.get(
            spreadsheetId=SPREADSHEET_ID,
            ranges=[f"'{sheet_name}'!A1:Z1"],
            includeGridData=False,
            fields="sheets(properties,columnGroups,data)"
        ).execute()
        col_meta = sheet_detail["sheets"][0]["properties"].get("gridProperties", {})
        # Column widths are in columnMetadata inside gridData — need separate call
        grid = sheets_api.get(
            spreadsheetId=SPREADSHEET_ID,
            ranges=[f"'{sheet_name}'!1:1"],
            includeGridData=True,
        ).execute()
        col_data = grid["sheets"][0].get("data", [{}])[0].get("columnMetadata", [])
        for col_idx, expected_px in expected_widths.items():
            if col_idx - 1 < len(col_data):
                actual_px = col_data[col_idx - 1].get("pixelSize", "?")
                if actual_px != expected_px:
                    info.append(
                        f"📐 \"{sheet_name}\" col {col_letter(col_idx)}: "
                        f"expected {expected_px}px, got {actual_px}px"
                    )

    # ── 5. Row heights — detect spacer rows still present ────────────────────
    print("📏 Checking for residual spacer rows...")
    for sheet_name in EXPECTED_TABS:
        if sheet_name == "Lookups" or sheet_name not in live_tabs:
            continue
        grid = sheets_api.get(
            spreadsheetId=SPREADSHEET_ID,
            ranges=[f"'{sheet_name}'!A:A"],
            includeGridData=True,
        ).execute()
        row_data = grid["sheets"][0].get("data", [{}])[0].get("rowMetadata", [])
        spacers = [i + 1 for i, r in enumerate(row_data) if r.get("pixelSize", 99) <= 2]
        if spacers:
            issues.append(
                f"⚠️  Spacer rows still present in \"{sheet_name}\": rows {spacers} "
                f"— run removeSpacerRows() in maintenance.gs"
            )

    # ── Report ────────────────────────────────────────────────────────────────
    print("\n" + "═" * 60)
    print(f"ISSUES ({len(issues)})")
    print("═" * 60)
    for i in issues:
        print(i)

    print("\n" + "─" * 60)
    print(f"INFO / CHANGES FROM BUILD SCRIPT ({len(info)})")
    print("─" * 60)
    for i in info:
        print(i)

    print("\n" + "═" * 60)
    summary = f"{'✅ All clear' if not issues else '⚠️  Issues found'}: {len(issues)} issue(s), {len(info)} informational"
    print(summary)

    # Write JSON report
    report = {"issues": issues, "info": info}
    with open("diff_report.json", "w") as f:
        json.dump(report, f, indent=2)
    print("📄 Full report written to diff_report.json")


if __name__ == "__main__":
    main()
