#!/usr/bin/env python3
"""
analyze_vendors.py — Vendor Spend Analysis Pipeline

Accepts any vendor AP ledger (CSV or Google Sheets) and produces a
fully annotated XLSX with department classifications, descriptions,
Terminate/Consolidate/Optimize recommendations, Top 3 cost-reduction
opportunities, and a CEO/CFO executive memo.

Reusable across any company's vendor dataset — no hardcoded vendor names,
spend figures, or business-specific assumptions.

Usage:
  python3 analyze_vendors.py --input vendors.csv
  python3 analyze_vendors.py --input https://docs.google.com/spreadsheets/d/...
  python3 analyze_vendors.py --input 1L2u8j-3cSFLPMXbbBljI4YYXQ9COYS5YihtBhxWTPls

Skip flags (reuse cached intermediate files):
  --skip-research    reuse vendors_researched.json
  --skip-classify    reuse vendors_classified.json
  --skip-qa          skip AI QA pass
  --skip-synthesis   reuse vendors_insights.json
"""

import argparse
import json
import os
import sys

from config import CLASSIFIED_JSON, RESEARCHED_JSON, QA_JSON, INSIGHTS_JSON, OUTPUT_XLSX
from fetch_data import get_vendors
from research_vendors import research_all_vendors
from classify_vendors import classify_all_vendors
from qa_review import run_qa
from synthesize_insights import synthesize
from validate_output import validate, print_report
from build_output import build_xlsx


def _require_api_key() -> None:
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print(
            "\nERROR: ANTHROPIC_API_KEY is not set.\n"
            "Run:  export ANTHROPIC_API_KEY=sk-ant-...\n"
        )
        sys.exit(1)


def _load_cached_csv_vendors() -> list[dict]:
    import csv
    vendors = []
    with open("vendors_raw.csv", encoding="utf-8") as f:
        for i, row in enumerate(csv.DictReader(f)):
            vendors.append({
                "vendor_name": row["vendor_name"],
                "cost_usd": float(row["cost_usd"]),
                "row_index": int(row.get("row_index", i + 2)),
            })
    vendors.sort(key=lambda v: v["cost_usd"], reverse=True)
    return vendors


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Vendor Spend Analysis — reusable AP ledger classification pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--input", "-i",
        required=False,
        help="Vendor data source: local CSV path, Google Sheets URL, or Sheets ID",
    )
    parser.add_argument("--skip-research",  action="store_true")
    parser.add_argument("--skip-classify",  action="store_true")
    parser.add_argument("--skip-qa",        action="store_true")
    parser.add_argument("--skip-synthesis", action="store_true")
    args = parser.parse_args()

    # If no --input but vendors_raw.csv exists, allow re-running from cache
    if not args.input:
        if os.path.exists("vendors_raw.csv"):
            print("No --input provided; using cached vendors_raw.csv")
            args.input = "vendors_raw.csv"
        else:
            print(
                "\nERROR: --input is required.\n"
                "Provide a CSV file path, Google Sheets URL, or Sheets ID.\n"
                "Example: python3 analyze_vendors.py --input vendors.csv\n"
            )
            sys.exit(1)

    print("\n" + "=" * 64)
    print("  VENDOR SPEND ANALYSIS")
    print("=" * 64 + "\n")

    # ── 1. Fetch ──────────────────────────────────────────────────
    print("[1/6] Loading vendor data...")
    if args.input == "vendors_raw.csv" and os.path.exists("vendors_raw.csv"):
        vendors = _load_cached_csv_vendors()
        print(f"  Loaded {len(vendors)} vendors from cached vendors_raw.csv")
    else:
        vendors = get_vendors(args.input)
        print(f"  Loaded {len(vendors)} vendors")

    total_spend = sum(v["cost_usd"] for v in vendors)
    print(f"  Total TTM spend: ${total_spend:,.0f}\n")

    # ── 2. Research ───────────────────────────────────────────────
    print("[2/6] Researching vendors...")
    if args.skip_research and os.path.exists(RESEARCHED_JSON):
        with open(RESEARCHED_JSON, encoding="utf-8") as f:
            research_map = {r["vendor_name"]: r for r in json.load(f)}
        enriched = [{
            **v,
            "what_they_do":        research_map.get(v["vendor_name"], {}).get("what_they_do", ""),
            "confidence":          research_map.get(v["vendor_name"], {}).get("confidence", "LOW"),
            "needs_human_review":  research_map.get(v["vendor_name"], {}).get("needs_human_review", False),
            "web_snippet":         research_map.get(v["vendor_name"], {}).get("web_snippet", ""),
        } for v in vendors]
        print(f"  Loaded research from {RESEARCHED_JSON} (cached)\n")
    else:
        _require_api_key()
        enriched = research_all_vendors(vendors)
        high = sum(1 for v in enriched if v.get("confidence") == "HIGH")
        low  = sum(1 for v in enriched if v.get("confidence") == "LOW")
        print(f"  Research complete: {high} HIGH confidence, {low} LOW confidence\n")

    # ── 3. Classify ───────────────────────────────────────────────
    print("[3/6] Classifying vendors with Claude API...")
    if args.skip_classify and os.path.exists(CLASSIFIED_JSON):
        with open(CLASSIFIED_JSON, encoding="utf-8") as f:
            classifications = json.load(f)
        print(f"  Loaded {len(classifications)} classifications (cached)\n")
    else:
        _require_api_key()
        classifications = classify_all_vendors(enriched)
        print(f"  Classified {len(classifications)} vendors\n")

    # ── 4. QA Review ──────────────────────────────────────────────
    print("[4/6] Running AI-based QA review...")
    qa_report = None
    if args.skip_qa:
        print("  QA pass skipped (--skip-qa)\n")
        for c in classifications:
            c.setdefault("qa_reclassified", False)
            c.setdefault("qa_warn", False)
    else:
        _require_api_key()
        classifications, qa_report = run_qa(vendors, classifications)
        print(
            f"  QA complete: {qa_report['ok']} ok | "
            f"{qa_report['warn']} warn | "
            f"{qa_report['error']} error | "
            f"{qa_report['reclassified']} re-classified\n"
        )

    # ── 5. Synthesize Insights ────────────────────────────────────
    print("[5/6] Synthesizing insights (Top 3 + executive memo)...")
    if args.skip_synthesis and os.path.exists(INSIGHTS_JSON):
        with open(INSIGHTS_JSON, encoding="utf-8") as f:
            insights = json.load(f)
        print(f"  Loaded insights from {INSIGHTS_JSON} (cached)\n")
    else:
        _require_api_key()
        insights = synthesize(vendors, classifications, qa_report)
        opps = insights.get("opportunities", [])
        total_low  = sum(o.get("savings_low_usd",  0) for o in opps[:3])
        total_high = sum(o.get("savings_high_usd", 0) for o in opps[:3])
        print(f"  Identified {len(opps)} opportunities | "
              f"Est. savings: ${total_low:,.0f} – ${total_high:,.0f}\n")

    # ── 6. Validate ───────────────────────────────────────────────
    print("[6/6] Running validation checks...")
    report = validate(vendors, classifications)
    print_report(report)

    if not report["passed"]:
        print("WARNING: Validation flagged issues — review report above before submitting.\n")

    # ── Build XLSX ────────────────────────────────────────────────
    print(f"Building {OUTPUT_XLSX}...")
    build_xlsx(vendors, classifications, insights, qa_report)

    # ── Summary ───────────────────────────────────────────────────
    rec_dist = report["recommendation_distribution"]
    opps = insights.get("opportunities", [])
    total_low  = sum(o.get("savings_low_usd",  0) for o in opps[:3])
    total_high = sum(o.get("savings_high_usd", 0) for o in opps[:3])

    print("\n" + "=" * 64)
    print("  ANALYSIS COMPLETE")
    print("=" * 64)
    print(f"  Vendors analyzed:    {len(vendors)}")
    print(f"  Total TTM spend:     ${total_spend:,.0f}")
    print(f"  Terminate:           {rec_dist.get('Terminate', 0)}")
    print(f"  Consolidate:         {rec_dist.get('Consolidate', 0)}")
    print(f"  Optimize:            {rec_dist.get('Optimize', 0)}")
    if total_low or total_high:
        print(f"  Est. annual savings: ${total_low:,.0f} – ${total_high:,.0f}")
    print(f"  Output:              {OUTPUT_XLSX}")
    print("=" * 64 + "\n")


if __name__ == "__main__":
    main()
