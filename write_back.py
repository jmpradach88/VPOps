"""
write_back.py — Write all analysis results back to the source Google Sheet,
and create a Google Doc for the executive memo.

Updates four tabs by reading each tab's existing headers and filling matching
columns — no clearing, no reformatting, just data in the right cells.

  1. Vendor Analysis     — fills Department, Description, Suggestion columns
  2. Top 3 Opportunities — fills opportunity title, description, savings columns
  3. Methodology         — fills label/value rows
  4. Recommendations     — fills a link to the Google Doc memo

Also creates:
  • A Google Doc containing the full 1-page executive memo
  • Sets the doc to "anyone with link can view"

Requires Google credentials. Two auth methods are supported:

  Method A — Service Account (recommended for automation):
    1. Go to console.cloud.google.com → IAM → Service Accounts → Create
    2. Download the JSON key file
    3. Share the Google Sheet with the service account email (Editor access)
    4. Set: export GOOGLE_CREDENTIALS_FILE=/path/to/credentials.json

  Method B — OAuth (personal use, one-time browser login):
    1. Go to console.cloud.google.com → APIs → Credentials → OAuth 2.0 Client ID
    2. Download as credentials.json
    3. Set: export GOOGLE_CREDENTIALS_FILE=/path/to/credentials.json
    4. First run opens browser for authorization; token saved to token.pickle
    5. Add your Gmail as a test user on the OAuth consent screen first
"""
from __future__ import annotations

import json
import os
import pickle
import re
import sys
from datetime import date

import gspread
from gspread.utils import rowcol_to_a1
from google.auth.transport.requests import Request
from google.oauth2.service_account import Credentials as ServiceCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build as build_service

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/documents",
    "https://www.googleapis.com/auth/drive.file",
]
TOKEN_CACHE = "token.pickle"

# Tab name patterns — matched case-insensitively against the actual sheet names
TAB_PATTERNS = {
    "vendors":       ["vendor", "spend", "ledger"],
    "opportunities": ["opportunit", "top 3", "top3"],
    "methodology":   ["method"],
    "memo":          ["recommend", "memo", "executive"],
}


# ── Auth ──────────────────────────────────────────────────────────────────────

def _get_credentials():
    creds_path = os.environ.get("GOOGLE_CREDENTIALS_FILE", "")
    if not creds_path:
        print(
            "\nERROR: GOOGLE_CREDENTIALS_FILE is not set.\n"
            "Set it to the path of your Google credentials JSON file.\n"
        )
        sys.exit(1)
    if not os.path.exists(creds_path):
        print(f"\nERROR: Credentials file not found: {creds_path}\n")
        sys.exit(1)

    with open(creds_path) as f:
        creds_data = json.load(f)

    if creds_data.get("type") == "service_account":
        return ServiceCredentials.from_service_account_file(creds_path, scopes=SCOPES)

    # OAuth — try cached token, but only if it has all required scopes
    if os.path.exists(TOKEN_CACHE):
        with open(TOKEN_CACHE, "rb") as f:
            creds = pickle.load(f)
        token_scopes = set(creds.scopes or [])
        required_scopes = set(SCOPES)
        if creds and creds.valid and required_scopes.issubset(token_scopes):
            return creds
        if creds and creds.expired and creds.refresh_token and required_scopes.issubset(token_scopes):
            creds.refresh(Request())
            with open(TOKEN_CACHE, "wb") as f:
                pickle.dump(creds, f)
            return creds
        # Scopes changed or token invalid — delete cache and re-auth
        os.remove(TOKEN_CACHE)
        print("  Re-authorizing (new permissions required for Docs + Drive)...")

    flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
    creds = flow.run_local_server(port=0)
    with open(TOKEN_CACHE, "wb") as f:
        pickle.dump(creds, f)
    return creds


def _find_worksheet(spreadsheet: gspread.Spreadsheet, patterns: list[str]):
    """Returns the first worksheet whose title matches any of the patterns."""
    for ws in spreadsheet.worksheets():
        title_lower = ws.title.lower()
        if any(p in title_lower for p in patterns):
            return ws
    return None


def _find_col(header: list[str], keywords: list[str]) -> int | None:
    """Returns 0-based column index of the first header matching any keyword."""
    for i, h in enumerate(header):
        if any(k in h.lower() for k in keywords):
            return i
    return None


def _ascii_key(name: str) -> str:
    """Keep only ASCII alphanumeric chars, lowercase.

    Used as a fallback match key so that vendor names with encoding artifacts
    (e.g. mojibake like 'SveuÃ¤Â\\x8dIli') match their proper UTF-8 equivalents
    ('Sveučili') returned by the Sheets API.
    """
    return re.sub(r'[^a-z0-9]', '', name.lower())


def _ensure_grid(ws, rows_needed: int, cols_needed: int) -> None:
    """Expand the worksheet grid if it is too small to hold the data."""
    if ws.row_count < rows_needed or ws.col_count < cols_needed:
        ws.resize(
            rows=max(ws.row_count, rows_needed),
            cols=max(ws.col_count, cols_needed),
        )


# ── Google Doc memo creation ──────────────────────────────────────────────────

def _build_memo_text(insights: dict) -> str:
    """Assembles the full memo as plain text for insertion into a Google Doc."""
    memo = insights.get("executive_memo", {})
    opps = insights.get("opportunities", [])
    total_low  = sum(o.get("savings_low_usd",  0) for o in opps[:3])
    total_high = sum(o.get("savings_high_usd", 0) for o in opps[:3])

    lines = [
        "MEMORANDUM",
        "",
        f"TO:      {memo.get('to', 'Chief Executive Officer, Chief Financial Officer')}",
        f"FROM:    {memo.get('from', 'VP of Operations')}",
        f"DATE:    {memo.get('date', insights.get('analysis_date', date.today().isoformat()))}",
        f"RE:      {memo.get('subject', 'Vendor Spend Analysis — Strategic Cost Reduction')}",
        "",
        "─" * 60,
        "",
        "EXECUTIVE SUMMARY",
        "",
        memo.get("executive_summary", ""),
        "",
        "─" * 60,
        "",
        "STRATEGIC OPPORTUNITIES",
    ]

    for opp in opps[:3]:
        savings_low  = opp.get("savings_low_usd",  0)
        savings_high = opp.get("savings_high_usd", 0)
        lines += [
            "",
            f"{opp.get('rank', '')}. {opp.get('title', '')}",
            f"   Est. Annual Savings: ${savings_low:,.0f} – ${savings_high:,.0f}",
            "",
            f"   {opp.get('description', '')}",
        ]
        if opp.get("implementation_steps"):
            lines += ["", f"   Implementation: {opp['implementation_steps']}"]
        if opp.get("risks"):
            lines += [f"   Risks: {opp['risks']}"]
        if opp.get("timeline"):
            lines += [f"   Timeline: {opp['timeline']}"]

    lines += [
        "",
        "─" * 60,
        "",
        "DEPARTMENT HIGHLIGHTS",
        "",
        memo.get("department_highlights", ""),
        "",
        "─" * 60,
        "",
        "RECOMMENDED IMMEDIATE ACTIONS (30-Day Sprint)",
        "",
        memo.get("immediate_actions", ""),
        "",
        "─" * 60,
        "",
        f"TOTAL ESTIMATED ANNUAL SAVINGS:  ${total_low:,.0f} – ${total_high:,.0f}",
        "",
        memo.get("total_savings_statement", ""),
    ]

    return "\n".join(lines)


def _create_memo_doc(creds, insights: dict) -> str:
    """
    Creates a Google Doc containing the executive memo.
    Sets it to 'anyone with link can view'.
    Returns the doc URL.
    """
    analysis_date = insights.get("analysis_date", date.today().isoformat())
    doc_title = f"Executive Memo — Vendor Spend Analysis ({analysis_date})"
    memo_text = _build_memo_text(insights)

    docs_service = build_service("docs", "v1", credentials=creds)
    doc = docs_service.documents().create(body={"title": doc_title}).execute()
    doc_id = doc["documentId"]

    docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": [{"insertText": {"location": {"index": 1}, "text": memo_text}}]},
    ).execute()

    drive_service = build_service("drive", "v3", credentials=creds)
    drive_service.permissions().create(
        fileId=doc_id,
        body={"type": "anyone", "role": "reader"},
    ).execute()

    return f"https://docs.google.com/document/d/{doc_id}/edit"


# ── Tab writers ───────────────────────────────────────────────────────────────

def _write_vendors_tab(ws, vendors: list[dict], classifications: list[dict]) -> None:
    """Fills Department, Description, and Suggestion columns on the vendor tab."""
    all_values = ws.get_all_values()
    if not all_values:
        print("  WARNING: Vendor tab appears empty — skipping.")
        return

    header = all_values[0]

    dept_col = _find_col(header, ["department"])
    desc_col = _find_col(header, ["description", "what the vendor", "vendor does"])
    rec_col  = _find_col(header, ["suggestion", "consolidate", "terminate", "optimize"])

    if None in (dept_col, desc_col, rec_col):
        print(f"  WARNING: Could not find all target columns in vendor tab. Headers: {header}")
        return

    class_map      = {c["vendor_name"].lower().strip(): c for c in classifications}
    ascii_class_map = {_ascii_key(c["vendor_name"]): c for c in classifications}
    updates = []
    matched, skipped = 0, 0

    for row_num, row in enumerate(all_values[1:], start=2):
        raw_name = row[0] if row else ""
        c = class_map.get(raw_name.lower().strip()) \
            or ascii_class_map.get(_ascii_key(raw_name))
        if not c:
            skipped += 1
            continue
        updates += [
            {"range": rowcol_to_a1(row_num, dept_col + 1), "values": [[c.get("department", "")]]},
            {"range": rowcol_to_a1(row_num, desc_col + 1), "values": [[c.get("description", "")]]},
            {"range": rowcol_to_a1(row_num, rec_col  + 1), "values": [[c.get("recommendation", "")]]},
        ]
        matched += 1

    if updates:
        ws.batch_update(updates, value_input_option="RAW")
    print(f"  Vendor tab: {matched} rows updated, {skipped} skipped (encoding mismatch).")


def _write_opportunities_tab(ws, insights: dict) -> None:
    """Fills opportunity data into the existing columns of the Top 3 Opportunities tab."""
    opps = insights.get("opportunities", [])
    if not opps:
        print("  WARNING: No opportunities in insights — skipping tab.")
        return

    all_values = ws.get_all_values()
    if not all_values:
        print("  WARNING: Opportunities tab appears empty — skipping.")
        return

    header = all_values[0]
    rank_col   = _find_col(header, ["#", "rank", "no.", "number"])
    title_col  = _find_col(header, ["opportunity", "title", "summary"])
    desc_col   = _find_col(header, ["description", "explanation", "brief", "detail"])
    vendor_col = _find_col(header, ["vendor", "affected"])
    spend_col  = _find_col(header, ["current spend", "spend", "cost"])
    low_col    = _find_col(header, ["savings low", "low", "min"])
    high_col   = _find_col(header, ["savings high", "high", "max"])
    timeline_col = _find_col(header, ["timeline"])
    risks_col  = _find_col(header, ["risk"])

    updates = []
    for i, opp in enumerate(opps[:3]):
        row_num = i + 2  # data starts at row 2

        def add(col, value):
            if col is not None:
                updates.append({"range": rowcol_to_a1(row_num, col + 1), "values": [[value]]})

        add(rank_col,    opp.get("rank", i + 1))
        add(title_col,   opp.get("title", ""))
        add(desc_col,    opp.get("description", "")
                         + ("\n\nActions: " + opp["implementation_steps"] if opp.get("implementation_steps") else ""))
        add(vendor_col,  ", ".join(opp.get("affected_vendors", [])))
        add(spend_col,   opp.get("current_spend_usd", ""))
        add(low_col,     opp.get("savings_low_usd", ""))
        add(high_col,    opp.get("savings_high_usd", ""))
        add(timeline_col, opp.get("timeline", ""))
        add(risks_col,   opp.get("risks", ""))

    # Write combined savings row if there's a row for it
    total_low  = sum(o.get("savings_low_usd",  0) for o in opps[:3])
    total_high = sum(o.get("savings_high_usd", 0) for o in opps[:3])
    summary_row = len(opps[:3]) + 2
    if title_col is not None:
        updates.append({"range": rowcol_to_a1(summary_row, title_col + 1), "values": [["COMBINED SAVINGS"]]})
    if low_col is not None:
        updates.append({"range": rowcol_to_a1(summary_row, low_col + 1), "values": [[total_low]]})
    if high_col is not None:
        updates.append({"range": rowcol_to_a1(summary_row, high_col + 1), "values": [[total_high]]})

    if updates:
        rows_needed = len(opps[:3]) + 2          # header + data rows + summary row
        cols_needed = max(
            (c for c in [rank_col, title_col, desc_col, vendor_col,
                         spend_col, low_col, high_col, timeline_col, risks_col]
             if c is not None),
            default=0,
        ) + 1
        _ensure_grid(ws, rows_needed, cols_needed)
        ws.batch_update(updates, value_input_option="RAW")

    # Wrap text in column C and auto-resize those rows so content is fully visible
    if desc_col is not None:
        data_rows = len(opps[:3])
        col_letter = chr(ord('A') + desc_col)
        ws.format(
            f"{col_letter}2:{col_letter}{data_rows + 1}",
            {"wrapStrategy": "WRAP"},
        )
        ws.spreadsheet.batch_update({
            "requests": [{
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": ws.id,
                        "dimension": "ROWS",
                        "startIndex": 1,            # row 2 = index 1
                        "endIndex": data_rows + 1,  # up to and including last data row
                    }
                }
            }]
        })

    print(f"  Opportunities tab: {len(opps[:3])} opportunities written.")


def _build_methodology_text(insights: dict, qa_report: dict | None) -> str:
    """Builds the full methodology content as a single plain-text block for cell A2."""
    qa_ok      = qa_report.get("ok",          "N/A") if qa_report else "N/A"
    qa_warn    = qa_report.get("warn",         "N/A") if qa_report else "N/A"
    qa_err     = qa_report.get("error",        "N/A") if qa_report else "N/A"
    qa_reclass = qa_report.get("reclassified", "N/A") if qa_report else "N/A"

    dept_lines = "\n".join(
        f"  {d['department']}: {d['vendor_count']} vendors, ${d['total_spend']:,.0f}"
        for d in sorted(
            insights.get("department_summary", []),
            key=lambda x: x.get("total_spend", 0), reverse=True,
        )
    )

    return (
        f"OVERVIEW\n"
        f"Analysis of {insights.get('total_vendors', '')} vendors totalling "
        f"${insights.get('total_spend_usd', 0):,.0f} TTM spend. "
        f"Analysis date: {insights.get('analysis_date', '')}.\n"
        f"\n"
        f"PIPELINE\n"
        f"1 — Fetch: Downloaded AP ledger via Google Sheets CSV export. "
        f"Auto-detects name and spend columns.\n"
        f"\n"
        f"2 — Research: Claude training-knowledge pass for all vendors (batches of 50). "
        f"LOW-confidence vendors above $20K spend received a DuckDuckGo web lookup. "
        f"Results cached to vendors_researched.json.\n"
        f"\n"
        f"3 — Classify: Claude API classification with prompt caching "
        f"(cache_control: ephemeral) on system prompt. Saves ~85% of input tokens "
        f"after batch 1. Batch size 50. Crash recovery after each batch.\n"
        f"\n"
        f"4 — QA Review: Second Claude pass reviews every classification as a senior "
        f"procurement auditor. Criteria: department fit, description quality, "
        f"recommendation consistency, factual accuracy. Error-flagged vendors "
        f"automatically re-classified with QA feedback as context.\n"
        f"\n"
        f"5 — Synthesize: Claude analyzes the full classified dataset and generates "
        f"Top 3 opportunities and executive memo from actual data. No hardcoded outputs.\n"
        f"\n"
        f"QA STATISTICS\n"
        f"  Passed (ok):      {qa_ok}\n"
        f"  Warnings:         {qa_warn}\n"
        f"  Errors flagged:   {qa_err}\n"
        f"  Re-classified:    {qa_reclass}\n"
        f"\n"
        f"RECOMMENDATION FRAMEWORK\n"
        f"  Terminate   — No recurring business value; one-off purchases; spend <$500 "
        f"with no recurring purpose; or duplicate entry.\n"
        f"  Consolidate — Overlapping service with another vendor on the list — "
        f"both flagged, duplicate named in note.\n"
        f"  Optimize    — Strategic vendor with high spend relative to market — "
        f"renegotiate, right-size, or seek volume discounts.\n"
        f"\n"
        f"SPEND BY DEPARTMENT\n"
        f"{dept_lines}\n"
        f"\n"
        f"TOOLS\n"
        f"  Python 3.9, anthropic SDK (claude-sonnet-4-6), openpyxl, gspread, "
        f"google-api-python-client, DuckDuckGo Instant Answer API\n"
        f"\n"
        f"LIMITATIONS\n"
        f"  1. No contract terms or seat data available — savings estimates based on benchmarks.\n"
        f"  2. Spend figures are AP payments; may differ from contracted amounts.\n"
        f"  3. Some vendor names contain encoding artifacts — a small number unclassified.\n"
        f"  4. Analysis does not cover vendor performance or SLA compliance."
    )


def _write_methodology_tab(ws, insights: dict, qa_report: dict | None) -> None:
    """Writes all methodology content into cell A2 as a single wrapped text block.

    Row 1 (the header) is left untouched to preserve existing formatting.
    A2 is set to wrap text and the row is auto-resized to fit the content.
    """
    _ensure_grid(ws, 2, 1)

    text = _build_methodology_text(insights, qa_report)
    ws.update("A2", [[text]], value_input_option="RAW")

    # Enable text wrap on A2
    ws.format("A2", {"wrapStrategy": "WRAP"})

    # Auto-resize row 2 to fit the content
    ws.spreadsheet.batch_update({
        "requests": [{
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": ws.id,
                    "dimension": "ROWS",
                    "startIndex": 1,   # 0-indexed; row 2 = index 1
                    "endIndex": 2,
                }
            }
        }]
    })
    print(f"  Methodology tab: written.")


def _write_memo_tab(ws, insights: dict, doc_url: str) -> None:
    """Writes the Google Doc link and a brief summary into the Recommendations tab."""
    all_values = ws.get_all_values()
    header = all_values[0] if all_values else []

    memo = insights.get("executive_memo", {})
    opps = insights.get("opportunities", [])
    total_low  = sum(o.get("savings_low_usd",  0) for o in opps[:3])
    total_high = sum(o.get("savings_high_usd", 0) for o in opps[:3])

    # Try to detect a label + value column structure
    label_col = _find_col(header, ["section", "label", "field", "item"])
    value_col = _find_col(header, ["detail", "value", "content", "link", "url", "notes"])
    if label_col is None:
        label_col = 0
    if value_col is None:
        value_col = 1

    analysis_date = insights.get("analysis_date", date.today().isoformat())

    content_rows = [
        ("Executive Memo (Google Doc)", doc_url),
        ("Date",    analysis_date),
        ("To",      memo.get("to", "Chief Executive Officer, Chief Financial Officer")),
        ("From",    memo.get("from", "VP of Operations")),
        ("Subject", memo.get("subject", "Vendor Spend Analysis — Strategic Cost Reduction")),
        ("Executive Summary", memo.get("executive_summary", "")),
        ("Total Est. Annual Savings", f"${total_low:,.0f} – ${total_high:,.0f}"),
    ]

    existing_labels = {
        row[label_col].strip().lower(): row_num + 1
        for row_num, row in enumerate(all_values[1:], start=1)
        if len(row) > label_col and row[label_col].strip()
    } if all_values else {}

    updates = []
    next_empty_row = (len(all_values) + 1) if all_values else 2

    for label, value in content_rows:
        matched_row = existing_labels.get(label.lower())
        if matched_row:
            row_num = matched_row
        else:
            row_num = next_empty_row
            next_empty_row += 1
            updates.append({"range": rowcol_to_a1(row_num, label_col + 1), "values": [[label]]})
        updates.append({"range": rowcol_to_a1(row_num, value_col + 1), "values": [[value]]})

    if updates:
        _ensure_grid(ws, next_empty_row - 1, value_col + 1)
        ws.batch_update(updates, value_input_option="RAW")
    print(f"  Recommendations tab: memo link written → {doc_url}")


# ── Public API ────────────────────────────────────────────────────────────────

def write_back(
    sheets_id: str,
    vendors: list[dict],
    classifications: list[dict],
    insights: dict | None = None,
    qa_report: dict | None = None,
) -> None:
    """
    Writes all analysis results back to the Google Sheet and creates a Google Doc memo.

    Reads each tab's existing column headers and fills matching cells — no clearing,
    no reformatting. Falls back gracefully if a tab or column is not found.
    """
    print(f"  Authenticating with Google Sheets...")
    creds = _get_credentials()
    client = gspread.authorize(creds)

    print(f"  Opening sheet: {sheets_id}")
    spreadsheet = client.open_by_key(sheets_id)
    sheet_titles = [ws.title for ws in spreadsheet.worksheets()]
    print(f"  Tabs found: {sheet_titles}")

    # Vendor tab
    ws = _find_worksheet(spreadsheet, TAB_PATTERNS["vendors"])
    if ws:
        _write_vendors_tab(ws, vendors, classifications)
    else:
        print("  WARNING: Could not find vendor tab — skipping.")

    if insights:
        # Opportunities tab
        ws = _find_worksheet(spreadsheet, TAB_PATTERNS["opportunities"])
        if ws:
            _write_opportunities_tab(ws, insights)
        else:
            print("  WARNING: Could not find Top 3 Opportunities tab — skipping.")

        # Methodology tab
        ws = _find_worksheet(spreadsheet, TAB_PATTERNS["methodology"])
        if ws:
            _write_methodology_tab(ws, insights, qa_report)
        else:
            print("  WARNING: Could not find Methodology tab — skipping.")

        # Create Google Doc memo, then link in Recommendations tab
        ws = _find_worksheet(spreadsheet, TAB_PATTERNS["memo"])
        if ws:
            print("  Creating executive memo Google Doc...")
            doc_url = _create_memo_doc(creds, insights)
            _write_memo_tab(ws, insights, doc_url)
        else:
            print("  WARNING: Could not find Recommendations tab — skipping.")
    else:
        print("  NOTE: No insights data — only vendor tab updated.")
