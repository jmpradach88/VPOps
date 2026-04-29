"""
write_back.py — Write classification results back to the source Google Sheet.

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
    4. First run will open a browser for authorization; token saved to token.json

The sheet must have these columns (auto-detected by header name):
  - Vendor Name
  - Department
  - Last 12 months Cost (USD)  [read-only]
  - 1-line Description on what the Vendor does
  - Suggestions (Consolidate / Terminate / Optimize costs)
"""
from __future__ import annotations

import json
import os
import sys

import gspread
from google.oauth2.service_account import Credentials as ServiceCredentials
from google.oauth2.credentials import Credentials as OAuthCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pickle

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TOKEN_CACHE = "token.pickle"

# Columns we write to — matched by header text, not position
WRITABLE_COLUMNS = {
    "department":        ["department"],
    "description":       ["1-line description", "description on what", "vendor does"],
    "recommendation":    ["suggestions", "consolidate", "terminate", "optimize"],
}


def _get_credentials() -> ServiceCredentials | OAuthCredentials:
    """
    Returns Google credentials from the GOOGLE_CREDENTIALS_FILE env var.
    Tries service account first; falls back to OAuth flow.
    """
    creds_path = os.environ.get("GOOGLE_CREDENTIALS_FILE", "")
    if not creds_path:
        print(
            "\nERROR: GOOGLE_CREDENTIALS_FILE is not set.\n"
            "Set it to the path of your Google credentials JSON file.\n"
            "See write_back.py module docstring for setup instructions.\n"
        )
        sys.exit(1)

    if not os.path.exists(creds_path):
        print(f"\nERROR: Credentials file not found: {creds_path}\n")
        sys.exit(1)

    with open(creds_path) as f:
        creds_data = json.load(f)

    # Service account credentials have a "type" field
    if creds_data.get("type") == "service_account":
        return ServiceCredentials.from_service_account_file(creds_path, scopes=SCOPES)

    # OAuth credentials — check for cached token first
    if os.path.exists(TOKEN_CACHE):
        with open(TOKEN_CACHE, "rb") as f:
            creds = pickle.load(f)
        if creds and creds.valid:
            return creds
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            with open(TOKEN_CACHE, "wb") as f:
                pickle.dump(creds, f)
            return creds

    # First-time OAuth: open browser for authorization
    flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
    creds = flow.run_local_server(port=0)
    with open(TOKEN_CACHE, "wb") as f:
        pickle.dump(creds, f)
    return creds


def _detect_column_indices(header_row: list[str]) -> dict[str, int]:
    """
    Finds the column index (0-based) for each writable field by matching
    header text against known patterns. Raises SystemExit if required
    columns can't be found.
    """
    lower_headers = [h.lower() for h in header_row]
    found = {}

    for field, patterns in WRITABLE_COLUMNS.items():
        for i, h in enumerate(lower_headers):
            if any(p in h for p in patterns):
                found[field] = i
                break

    missing = [f for f in WRITABLE_COLUMNS if f not in found]
    if missing:
        print(
            f"\nERROR: Could not find columns for: {missing}\n"
            f"Headers detected: {header_row}\n"
        )
        sys.exit(1)

    return found


def write_back(
    sheets_id: str,
    vendors: list[dict],
    classifications: list[dict],
) -> None:
    """
    Updates the Department, Description, and Suggestion columns in the
    Google Sheet for every vendor that has a classification.

    Uses batch updates (single API call per sheet) to minimize quota usage.
    Vendors with encoding mismatches that can't be matched are skipped and
    reported at the end.
    """
    print(f"  Authenticating with Google Sheets...")
    creds = _get_credentials()
    client = gspread.authorize(creds)

    print(f"  Opening sheet: {sheets_id}")
    spreadsheet = client.open_by_key(sheets_id)
    worksheet = spreadsheet.get_worksheet(0)

    all_values = worksheet.get_all_values()
    if not all_values:
        print("ERROR: Sheet appears to be empty.")
        sys.exit(1)

    header_row = all_values[0]
    col_idx = _detect_column_indices(header_row)

    dept_col = col_idx["department"]
    desc_col = col_idx["description"]
    rec_col  = col_idx["recommendation"]

    # Build a lookup: vendor name (lowercase, stripped) → classification
    class_map = {
        c["vendor_name"].lower().strip(): c
        for c in classifications
    }

    # Prepare batch update cells
    updates = []
    matched = 0
    skipped = []

    for row_num, row in enumerate(all_values[1:], start=2):  # 1-based, skip header
        if not row:
            continue
        raw_name = row[0] if row else ""
        name_key = raw_name.lower().strip()

        c = class_map.get(name_key)
        if not c:
            skipped.append(raw_name)
            continue

        # gspread uses (row, col) with 1-based columns
        updates.append({
            "range": gspread.utils.rowcol_to_a1(row_num, dept_col + 1),
            "values": [[c.get("department", "")]],
        })
        updates.append({
            "range": gspread.utils.rowcol_to_a1(row_num, desc_col + 1),
            "values": [[c.get("description", "")]],
        })
        updates.append({
            "range": gspread.utils.rowcol_to_a1(row_num, rec_col + 1),
            "values": [[c.get("recommendation", "")]],
        })
        matched += 1

    if not updates:
        print("  No rows matched — nothing written.")
        return

    print(f"  Writing {matched} rows ({len(updates)} cells) in a single batch...")
    worksheet.batch_update(updates, value_input_option="RAW")

    print(f"  Done. {matched} vendors updated in Google Sheets.")
    if skipped:
        print(f"  Skipped {len(skipped)} vendors (encoding mismatch or not in classifications):")
        for name in skipped[:5]:
            print(f"    {name}")
        if len(skipped) > 5:
            print(f"    ... and {len(skipped) - 5} more")
