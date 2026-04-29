"""
fetch_data.py — Load vendor data from any CSV file or Google Sheets URL/ID.

Accepts:
  - A local CSV file path          e.g. vendors.csv
  - A Google Sheets share URL      e.g. https://docs.google.com/spreadsheets/d/...
  - A bare Google Sheets ID        e.g. 1L2u8j-3cSFLPMXbbBljI4YYXQ9COYS5YihtBhxWTPls

CSV must have at minimum two columns (names are flexible — detected automatically):
  - One column containing vendor / supplier names
  - One column containing a spend amount (USD)

Column detection priority:
  Name column  → "Vendor Name", "Vendor", "Supplier", "Company", "Name"  (case-insensitive)
  Spend column → "Cost", "Spend", "Amount", "Total", "Last 12"            (case-insensitive substring)
"""
from __future__ import annotations

import csv
import io
import re
import sys

import requests

from config import RAW_CSV

# Google Sheets export base URL
_SHEETS_EXPORT = "https://docs.google.com/spreadsheets/d/{id}/export?format=csv&gid=0"

# Column detection: ordered lists of candidate header substrings (lowercase)
_NAME_CANDIDATES  = ["vendor name", "vendor", "supplier", "company name", "company", "name"]
_SPEND_CANDIDATES = ["last 12", "cost", "spend", "amount", "total"]


def _extract_sheets_id(source: str) -> str | None:
    """Returns the Sheets ID if source looks like a Sheets URL or a bare ID."""
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", source)
    if m:
        return m.group(1)
    # Bare ID: alphanumeric + underscores/hyphens, 20+ chars
    if re.fullmatch(r"[A-Za-z0-9_-]{20,}", source):
        return source
    return None


def _detect_columns(headers: list[str]) -> tuple[str, str]:
    """
    Returns (name_col, spend_col) by matching headers against candidate lists.
    Raises SystemExit with a clear message if detection fails.
    """
    lower = [h.lower() for h in headers]

    name_col = None
    for candidate in _NAME_CANDIDATES:
        for i, h in enumerate(lower):
            if candidate in h:
                name_col = headers[i]
                break
        if name_col:
            break

    spend_col = None
    for candidate in _SPEND_CANDIDATES:
        for i, h in enumerate(lower):
            if candidate in h:
                spend_col = headers[i]
                break
        if spend_col:
            break

    if not name_col or not spend_col:
        print(
            f"\nERROR: Could not detect required columns in your CSV.\n"
            f"  Headers found: {headers}\n"
            f"  Need a name column containing one of: {_NAME_CANDIDATES}\n"
            f"  Need a spend column containing one of: {_SPEND_CANDIDATES}\n"
            f"  Rename your columns to match and retry.\n"
        )
        sys.exit(1)

    return name_col, spend_col


def parse_cost(cost_str: str) -> float:
    cleaned = re.sub(r"[\$,\s£€]", "", str(cost_str))
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _parse_csv_text(text: str) -> list[dict]:
    """Parse CSV text into vendor dicts; auto-detects name and spend columns."""
    reader = csv.DictReader(io.StringIO(text))
    headers = reader.fieldnames or []
    name_col, spend_col = _detect_columns(list(headers))

    vendors = []
    for i, row in enumerate(reader):
        name = row.get(name_col, "").strip()
        if not name:
            continue
        vendors.append({
            "vendor_name": name,
            "cost_usd": parse_cost(row.get(spend_col, "0")),
            "row_index": i + 2,
        })
    return vendors


def load_from_sheets(sheets_id: str) -> list[dict] | None:
    """Downloads CSV from Google Sheets. Returns None on failure."""
    url = _SHEETS_EXPORT.format(id=sheets_id)
    try:
        resp = requests.get(url, timeout=15, allow_redirects=True)
        if resp.status_code != 200:
            return None
        content = resp.text
        if content.strip().startswith(("<!DOCTYPE", "<html")):
            print("  Google Sheets requires authentication for this file.")
            print("  Export the sheet manually as CSV and pass it with --input path/to/file.csv")
            return None
        vendors = _parse_csv_text(content)
        if vendors:
            _save_raw_csv(vendors)
        return vendors or None
    except Exception as e:
        print(f"  Sheets download failed: {e}")
        return None


def load_from_csv(path: str) -> list[dict]:
    """Reads a local CSV file. Raises SystemExit if not found or unparseable."""
    try:
        with open(path, encoding="utf-8-sig") as f:
            text = f.read()
    except FileNotFoundError:
        print(f"\nERROR: Input file not found: {path}\n")
        sys.exit(1)
    vendors = _parse_csv_text(text)
    if not vendors:
        print(f"\nERROR: No vendor rows found in {path}.\n")
        sys.exit(1)
    _save_raw_csv(vendors)
    return vendors


def _save_raw_csv(vendors: list[dict]) -> None:
    with open(RAW_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["vendor_name", "cost_usd", "row_index"])
        writer.writeheader()
        writer.writerows(vendors)


def get_vendors(source: str) -> list[dict]:
    """
    Master entry point. Accepts a file path, Sheets URL, or Sheets ID.
    Returns vendors sorted by cost descending.
    """
    sheets_id = _extract_sheets_id(source)

    if sheets_id:
        print(f"  Detected Google Sheets ID: {sheets_id}")
        vendors = load_from_sheets(sheets_id)
        if not vendors:
            sys.exit(1)
    else:
        print(f"  Loading from CSV: {source}")
        vendors = load_from_csv(source)

    vendors.sort(key=lambda v: v["cost_usd"], reverse=True)
    return vendors
