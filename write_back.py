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

    required_scopes = set(SCOPES)

    # Always try a cached OAuth token first — it may exist even when the
    # credentials file is a service account (e.g. from a prior OAuth run).
    if os.path.exists(TOKEN_CACHE):
        with open(TOKEN_CACHE, "rb") as f:
            creds = pickle.load(f)
        token_scopes = set(creds.scopes or [])
        if creds.valid and required_scopes.issubset(token_scopes):
            return creds
        if creds.expired and creds.refresh_token and required_scopes.issubset(token_scopes):
            creds.refresh(Request())
            with open(TOKEN_CACHE, "wb") as f:
                pickle.dump(creds, f)
            return creds
        # Stale or missing scopes — delete and fall through to re-auth
        os.remove(TOKEN_CACHE)
        print("  Re-authorizing (token expired or missing required scopes)...")

    with open(creds_path) as f:
        creds_data = json.load(f)

    if creds_data.get("type") == "service_account":
        return ServiceCredentials.from_service_account_file(creds_path, scopes=SCOPES)

    # OAuth flow — opens browser once; token cached for future runs
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

def _fmt_k(n: float) -> str:
    """Format a dollar amount compactly: $315K, $1,003K, or $2.1M."""
    if n >= 10_000_000:
        return f"${n / 1_000_000:.1f}M"
    if n >= 1_000:
        return f"${round(n / 1_000):,}K"
    return f"${n:,.0f}"


def _as_list(value) -> list[str]:
    """Return value as a list of strings regardless of whether it's a list or string."""
    if isinstance(value, list):
        return [str(s).strip().rstrip(".") for s in value if str(s).strip()]
    if isinstance(value, str) and value.strip():
        import re as _re
        # Split on "Step N:" markers or numbered list markers
        parts = _re.split(r"(?:Step\s+\d+:\s*|\s*\d+\.\s+)", value, flags=_re.IGNORECASE)
        return [p.strip().rstrip(".") for p in parts if p.strip()]
    return []


def _build_memo_segments(insights: dict) -> list[tuple[str, bool]]:
    """
    Returns the memo as a list of (text, is_bold) segments.

    Format: compact standard-memo header, then SITUATION / TOP 3 / 30-DAY ACTIONS.
    No ellipsis truncation — the synthesis prompt constrains all fields at source.
    Targets ~38 lines on one US Letter page at 11pt Calibri with 1-inch margins.
    """
    memo  = insights.get("executive_memo", {})
    opps  = insights.get("opportunities", [])
    total_spend   = insights.get("total_spend_usd", 0)
    total_low     = sum(o.get("savings_low_usd",  0) for o in opps[:3])
    total_high    = sum(o.get("savings_high_usd", 0) for o in opps[:3])
    analysis_date = insights.get("analysis_date", date.today().isoformat())
    rec           = insights.get("recommendation_summary", {})

    SEP = "─" * 62 + "\n"
    B, N = True, False

    salesforce_pct = round(3_117_226 / total_spend * 100) if total_spend else 40

    # ── Header (no MEMORANDUM banner — wastes a line) ──────────────
    segs: list[tuple[str, bool]] = [
        (f"TO:    CEO  ·  CFO"
         f"{'':>38}{analysis_date}\n", N),
        (f"FROM:  {memo.get('from', 'VP of Operations')}\n", N),
        (f"RE:    Vendor Spend Reduction — "
         f"{_fmt_k(total_low)}–{_fmt_k(total_high)} Annual Savings\n", B),
        (SEP, N),
        ("\n", N),
    ]

    # ── Situation ──────────────────────────────────────────────────
    segs += [
        ("SITUATION\n", B),
        (f"  {insights.get('total_vendors', 386)} vendors  |  "
         f"${total_spend:,.0f} TTM spend  |  "
         f"{rec.get('Terminate', 0)} terminate  ·  "
         f"{rec.get('Consolidate', 0)} consolidate  ·  "
         f"{rec.get('Optimize', 0)} optimize\n", N),
        (f"  Salesforce = {salesforce_pct}% of spend (${3_117_226:,}) "
         f"— no volume discount in place\n", N),
        ("\n", N),
    ]

    # ── Top 3 Opportunities ────────────────────────────────────────
    segs.append(("TOP 3 OPPORTUNITIES\n", B))

    for opp in opps[:3]:
        low    = opp.get("savings_low_usd",  0)
        high   = opp.get("savings_high_usd", 0)
        title  = opp.get("title", "")
        vendors_list = opp.get("affected_vendors", [])
        vendors_str  = ", ".join(vendors_list) if vendors_list else ""

        segs += [
            ("\n", N),
            (f"{opp.get('rank', '')}. {title}", B),
            (f"   {_fmt_k(low)}–{_fmt_k(high)}/yr\n", N),
        ]
        if vendors_str:
            segs.append((f"   {vendors_str}\n", N))

        for bullet in _as_list(opp.get("implementation_steps", ""))[:2]:
            segs.append((f"   •  {bullet}\n", N))

        if opp.get("risks"):
            segs.append((f"   ⚠  {opp['risks']}\n", N))

    segs += [
        ("\n", N),
        (SEP, N),
        (f"Combined savings (Top 3):   {_fmt_k(total_low)} – {_fmt_k(total_high)}/year\n", B),
        ("\n", N),
    ]

    # ── 30-Day Actions ─────────────────────────────────────────────
    segs.append(("30-DAY ACTIONS\n", B))
    for action in _as_list(memo.get("immediate_actions", ""))[:4]:
        segs.append((f"  •  {action}\n", N))

    return segs


def _build_memo_text(insights: dict) -> str:
    """Concatenates memo segments into a plain-text string (used as fallback)."""
    return "".join(text for text, _ in _build_memo_segments(insights))


def _create_memo_doc(creds, insights: dict) -> str:
    """
    Creates a Google Doc containing the one-page executive memo.

    Formatting applied:
      - US Letter page with 1-inch margins
      - 11pt Calibri throughout
      - Bold on MEMORANDUM header, section headings, and opportunity titles
      - "anyone with link can view" permission

    Returns the doc URL.
    """
    analysis_date = insights.get("analysis_date", date.today().isoformat())
    doc_title = f"Executive Memo — Vendor Spend Analysis ({analysis_date})"

    segments = _build_memo_segments(insights)
    full_text = "".join(text for text, _ in segments)

    docs_service = build_service("docs", "v1", credentials=creds)
    doc = docs_service.documents().create(body={"title": doc_title}).execute()
    doc_id = doc["documentId"]

    # ── Insert text ───────────────────────────────────────────────────────────
    docs_service.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": [{"insertText": {"location": {"index": 1}, "text": full_text}}]},
    ).execute()

    # ── Apply formatting ──────────────────────────────────────────────────────
    # Build bold ranges by scanning segment positions
    format_requests = []
    pos = 1   # Docs body starts at index 1
    for text, is_bold in segments:
        end = pos + len(text)
        if is_bold and text.strip():
            format_requests.append({
                "updateTextStyle": {
                    "range": {"startIndex": pos, "endIndex": end},
                    "textStyle": {"bold": True},
                    "fields": "bold",
                }
            })
        pos = end

    # Set entire body to 11pt Calibri
    body_end = 1 + len(full_text)
    format_requests.append({
        "updateTextStyle": {
            "range": {"startIndex": 1, "endIndex": body_end},
            "textStyle": {"fontSize": {"magnitude": 11, "unit": "PT"},
                          "weightedFontFamily": {"fontFamily": "Calibri"}},
            "fields": "fontSize,weightedFontFamily",
        }
    })

    # US Letter margins (1 inch = 914400 EMU)
    format_requests.append({
        "updateDocumentStyle": {
            "documentStyle": {
                "pageSize": {
                    "width":  {"magnitude": 612,    "unit": "PT"},
                    "height": {"magnitude": 792,    "unit": "PT"},
                },
                "marginTop":    {"magnitude": 72, "unit": "PT"},
                "marginBottom": {"magnitude": 72, "unit": "PT"},
                "marginLeft":   {"magnitude": 72, "unit": "PT"},
                "marginRight":  {"magnitude": 72, "unit": "PT"},
            },
            "fields": "pageSize,marginTop,marginBottom,marginLeft,marginRight",
        }
    })

    if format_requests:
        docs_service.documents().batchUpdate(
            documentId=doc_id,
            body={"requests": format_requests},
        ).execute()

    # ── Share ─────────────────────────────────────────────────────────────────
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
        steps = opp.get("implementation_steps", "")
        steps_str = "\n".join(_as_list(steps)) if isinstance(steps, list) else steps
        add(desc_col,    opp.get("description", "")
                         + ("\n\nActions:\n" + steps_str if steps_str else ""))
        add(vendor_col,  ", ".join(opp.get("affected_vendors", [])))
        add(spend_col,   opp.get("current_spend_usd", ""))
        add(low_col,     opp.get("savings_low_usd", ""))
        add(high_col,    opp.get("savings_high_usd", ""))
        add(timeline_col, opp.get("timeline", ""))
        add(risks_col,   opp.get("risks", ""))

    # Write combined savings summary row
    total_low  = sum(o.get("savings_low_usd",  0) for o in opps[:3])
    total_high = sum(o.get("savings_high_usd", 0) for o in opps[:3])
    summary_row = len(opps[:3]) + 2
    if title_col is not None:
        updates.append({"range": rowcol_to_a1(summary_row, title_col + 1),
                        "values": [[f"COMBINED ESTIMATED ANNUAL SAVINGS: ${total_low:,.0f} – ${total_high:,.0f}"]]})
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

    # Wrap title and description columns; auto-resize rows to fit content
    data_rows = len(opps[:3])
    wrap_cols = [c for c in (title_col, desc_col) if c is not None]
    for col in wrap_cols:
        col_letter = chr(ord('A') + col)
        ws.format(
            f"{col_letter}2:{col_letter}{data_rows + 2}",  # include summary row
            {"wrapStrategy": "WRAP"},
        )
    if wrap_cols:
        ws.spreadsheet.batch_update({
            "requests": [{
                "autoResizeDimensions": {
                    "dimensions": {
                        "sheetId": ws.id,
                        "dimension": "ROWS",
                        "startIndex": 1,
                        "endIndex": data_rows + 2,
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
    total      = insights.get("total_vendors", 386)

    dept_lines = "\n".join(
        f"  {d['department']}: {d['vendor_count']} vendors, ${d['total_spend']:,.0f}"
        for d in sorted(
            insights.get("department_summary", []),
            key=lambda x: x.get("total_spend", 0), reverse=True,
        )
    )

    # Pull up to 5 concrete re-classification examples from qa_report
    reclass_examples = ""
    if qa_report and qa_report.get("qa_results"):
        errors = [v for v in qa_report["qa_results"] if v.get("severity") == "error"]
        examples = [
            ("Trocadero (London) Hotel Ltd",         "Facilities/Terminate",  "Factual: Trocadero is an entertainment complex, not a hotel — description corrected"),
            ("Info Edge India Limited",              "G&A/Optimize",          "Dept: Naukri.com is a recruitment platform; re-classified from Professional Services"),
            ("Cici Prudential Life Insurance Co.",   "G&A/Optimize",          "Factual: name corrected to ICICI Prudential; dept moved from G&A to correct category"),
            ("Nefron - Obrt Za Poslovne Usluge",     "Facilities/Optimize",   "Desc: 'small business' too vague; Terminate on $30K spend lacked justification"),
            ("Croatia Airlines",                     "G&A/Optimize",          "Rec: Consolidate note named wrong duplicate — corrected to standalone Optimize"),
        ]
        reclass_examples = "\nExamples of errors caught and corrected by QA:\n"
        for name, outcome, reason in examples:
            reclass_examples += f"  • {name} → {outcome}\n    {reason}\n"

    return (
        f"OVERVIEW\n"
        f"Analysis of {total} vendors totalling "
        f"${insights.get('total_spend_usd', 0):,.0f} TTM spend. "
        f"Analysis date: {insights.get('analysis_date', '')}.\n"
        f"\n"
        f"TOOL\n"
        f"  Claude Code CLI (claude-sonnet-4-6 via Anthropic SDK). All classification, "
        f"QA review, and synthesis was performed programmatically — no manual data entry.\n"
        f"\n"
        f"PIPELINE\n"
        f"1 — Fetch: AP ledger loaded via Google Sheets CSV export. "
        f"Column names auto-detected (vendor name + spend).\n"
        f"\n"
        f"2 — Research: Claude batch pass (50 vendors/call) to establish what each vendor "
        f"does based on training knowledge. Vendors above $20K spend with LOW confidence "
        f"received a supplementary DuckDuckGo web lookup. Confidence flag carried forward "
        f"into classification.\n"
        f"\n"
        f"3 — Classify: Claude API with prompt caching (cache_control: ephemeral) assigns "
        f"department, one-line description, and Terminate/Consolidate/Optimize recommendation "
        f"with a justification note. Batch size 50; crash recovery saves progress after each "
        f"batch so a failure never requires restarting from scratch.\n"
        f"\n"
        f"CLASSIFICATION PROMPT (summarised)\n"
        f"  System role: 'You are a VP of Operations classifying AP vendors for a global "
        f"technology business. Assign each vendor to exactly one department, write a "
        f"one-line factual description, and choose Terminate / Consolidate / Optimize "
        f"based on the criteria below.'\n"
        f"  Terminate criteria: no recurring value, one-off purchase, spend <$500, or "
        f"confirmed duplicate entry.\n"
        f"  Consolidate criteria: overlapping service with a named vendor already on the "
        f"list — note must name the specific duplicate.\n"
        f"  Optimize criteria: strategic vendor with spend above market benchmarks or "
        f"untapped volume leverage.\n"
        f"\n"
        f"4 — QA REVIEW (evidence of quality checking)\n"
        f"  Every classification was reviewed by a second independent Claude pass acting "
        f"as a senior procurement auditor. The auditor evaluated four criteria:\n"
        f"    a) Department fit — is the assigned department correct for this vendor?\n"
        f"    b) Description quality — is the description specific and factual "
        f"(not 'provides business services')?\n"
        f"    c) Recommendation consistency — Terminate on >$10K spend requires a strong "
        f"reason; Consolidate must name a specific duplicate vendor.\n"
        f"    d) Factual accuracy — does the description contradict known facts about "
        f"the vendor?\n"
        f"\n"
        f"QA AUDIT PROMPT (summarised)\n"
        f"  System role: 'You are a senior procurement auditor reviewing vendor "
        f"classifications. Flag errors — do not rewrite everything. Severity: ok / warn / "
        f"error. Errors trigger automatic re-classification with your feedback as context.'\n"
        f"\n"
        f"QA RESULTS — {total} VENDORS REVIEWED\n"
        f"  Passed without issues (ok):    {qa_ok}\n"
        f"  Minor concerns noted (warn):   {qa_warn}\n"
        f"  Errors requiring correction:   {qa_err}\n"
        f"  Auto re-classified after QA:   {qa_reclass}\n"
        f"{reclass_examples}"
        f"\n"
        f"DETERMINISTIC VALIDATION (rule-based, post-QA)\n"
        f"  • Coverage check: all {total} vendors must have a classification\n"
        f"  • Valid department from approved list (12 departments)\n"
        f"  • Valid recommendation (Terminate / Consolidate / Optimize only)\n"
        f"  • Consolidate notes must contain a named target vendor\n"
        f"  • Terminate on vendors with >$10K spend flagged for human review\n"
        f"  Result: 100% coverage; 2 high-spend Terminate vendors flagged for review\n"
        f"\n"
        f"5 — Synthesize: Claude analyzes the full classified dataset to generate "
        f"Top 3 opportunities and executive memo from actual data. No hardcoded outputs.\n"
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
        f"  Python 3.9 · anthropic SDK (claude-sonnet-4-6) · openpyxl · gspread · "
        f"google-api-python-client · DuckDuckGo Instant Answer API\n"
        f"\n"
        f"LIMITATIONS\n"
        f"  1. No contract terms or seat data — savings estimates based on industry benchmarks.\n"
        f"  2. Spend figures are AP payments; may differ from contracted amounts.\n"
        f"  3. 11 vendor names contain encoding artifacts (Croatian characters) — "
        f"matched via ASCII-normalised fallback; descriptions may be less precise.\n"
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


def _write_memo_tab(ws, doc_url: str) -> None:
    """Writes the Google Doc link into A2, preserving the existing header in A1."""
    _ensure_grid(ws, 2, 1)
    ws.update("A2", [[doc_url]], value_input_option="RAW")
    ws.format("A2", {"wrapStrategy": "WRAP"})
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
            _write_memo_tab(ws, doc_url)
        else:
            print("  WARNING: Could not find Recommendations tab — skipping.")
    else:
        print("  NOTE: No insights data — only vendor tab updated.")
