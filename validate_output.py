"""
validate_output.py — Generic, dataset-agnostic validation rules.

These checks apply to any vendor dataset — no hardcoded vendor names or
spend amounts. The goal is to catch systematic errors in the AI pipeline,
not to assert business-specific facts.
"""
from __future__ import annotations

from config import DEPARTMENTS

VALID_RECOMMENDATIONS = {"Terminate", "Consolidate", "Optimize"}
MAX_DESCRIPTION_WORDS = 20
HIGH_SPEND_TERMINATE_THRESHOLD = 10_000  # flag Terminate on vendors above this


def validate(vendors: list[dict], classifications: list[dict]) -> dict:
    """
    Runs deterministic quality checks. Returns a report dict.

    Checks:
    - Coverage: all vendors have a classification
    - Valid departments: each classification uses a known department
    - Valid recommendations: each is Terminate / Consolidate / Optimize
    - Description length: no description exceeds MAX_DESCRIPTION_WORDS words
    - Consolidate completeness: Consolidate items should name a target in recommendation_note
    - High-spend Terminate: vendors above threshold flagged Terminate → human review
    - Department coverage: all configured departments appear at least once
    """
    class_map = {c["vendor_name"]: c for c in classifications}
    cost_map  = {v["vendor_name"]: v["cost_usd"] for v in vendors}

    vendor_names     = {v["vendor_name"] for v in vendors}
    classified_names = set(class_map.keys())

    missing = sorted(vendor_names - classified_names)
    coverage = len(classified_names & vendor_names) / len(vendor_names) if vendor_names else 0

    invalid_depts = [
        {"vendor_name": c["vendor_name"], "department": c.get("department")}
        for c in classifications
        if c.get("department") not in DEPARTMENTS
    ]

    invalid_recs = [
        {"vendor_name": c["vendor_name"], "recommendation": c.get("recommendation")}
        for c in classifications
        if c.get("recommendation") not in VALID_RECOMMENDATIONS
    ]

    desc_too_long = [
        {"vendor_name": c["vendor_name"],
         "word_count": len(c.get("description", "").split())}
        for c in classifications
        if len(c.get("description", "").split()) > MAX_DESCRIPTION_WORDS
    ]

    # Consolidate items should reference the duplicate in recommendation_note
    consolidate_missing_target = [
        {"vendor_name": c["vendor_name"], "note": c.get("recommendation_note", "")}
        for c in classifications
        if c.get("recommendation") == "Consolidate"
        and not c.get("recommendation_note", "").strip()
    ]

    # High-spend vendors flagged Terminate (may be legitimate, but flag for review)
    high_spend_terminate = [
        {
            "vendor_name": c["vendor_name"],
            "cost_usd": cost_map.get(c["vendor_name"], 0),
            "recommendation_note": c.get("recommendation_note", ""),
        }
        for c in classifications
        if c.get("recommendation") == "Terminate"
        and cost_map.get(c["vendor_name"], 0) >= HIGH_SPEND_TERMINATE_THRESHOLD
    ]

    dept_dist: dict[str, int] = {}
    rec_dist:  dict[str, int] = {}
    for c in classifications:
        d = c.get("department", "Unknown")
        r = c.get("recommendation", "Unknown")
        dept_dist[d] = dept_dist.get(d, 0) + 1
        rec_dist[r]  = rec_dist.get(r,  0) + 1

    missing_depts = [d for d in DEPARTMENTS if d not in dept_dist]

    passed = (
        coverage >= 0.98
        and not invalid_depts
        and not invalid_recs
    )

    return {
        "passed":                     passed,
        "coverage":                   coverage,
        "total_vendors":              len(vendor_names),
        "classified_count":           len(classified_names & vendor_names),
        "missing_vendors":            missing,
        "invalid_departments":        invalid_depts,
        "invalid_recommendations":    invalid_recs,
        "consolidate_missing_target": consolidate_missing_target,
        "high_spend_terminate":       high_spend_terminate,
        "description_too_long":       desc_too_long,
        "department_distribution":    dict(sorted(dept_dist.items())),
        "recommendation_distribution": rec_dist,
        "missing_departments":        missing_depts,
    }


def print_report(report: dict) -> None:
    W = 64
    print("=" * W)
    print("VALIDATION REPORT")
    print("=" * W)
    print(f"Status:     {'PASSED' if report['passed'] else 'NEEDS REVIEW'}")
    print(f"Coverage:   {report['coverage']:.1%}  "
          f"({report['classified_count']}/{report['total_vendors']} vendors)")
    print()

    print("DEPARTMENT DISTRIBUTION")
    for dept, count in report["department_distribution"].items():
        bar = "█" * (count // 3)
        print(f"  {dept:<26} {count:>3}  {bar}")
    if report["missing_departments"]:
        print(f"  [MISSING]: {', '.join(report['missing_departments'])}")
    print()

    print("RECOMMENDATION DISTRIBUTION")
    for rec, count in sorted(report["recommendation_distribution"].items()):
        print(f"  {rec:<15} {count:>3}")
    print()

    if report["high_spend_terminate"]:
        print(f"[HIGH-SPEND TERMINATE — HUMAN REVIEW RECOMMENDED]")
        for v in sorted(report["high_spend_terminate"], key=lambda x: -x["cost_usd"]):
            print(f"  ${v['cost_usd']:>10,.0f}  {v['vendor_name']}")
            if v["recommendation_note"]:
                print(f"                Reason: {v['recommendation_note']}")
        print()

    if report["consolidate_missing_target"]:
        print(f"[CONSOLIDATE WITHOUT TARGET: {len(report['consolidate_missing_target'])} item(s)]")
        for v in report["consolidate_missing_target"]:
            print(f"  {v['vendor_name']} — no target named in note")
        print()

    if report["description_too_long"]:
        print(f"[LONG DESCRIPTIONS: {len(report['description_too_long'])} vendor(s) >{MAX_DESCRIPTION_WORDS} words]")
        for v in report["description_too_long"][:5]:
            print(f"  {v['vendor_name']} ({v['word_count']} words)")
        print()

    if report["missing_vendors"]:
        print(f"[MISSING CLASSIFICATIONS: {len(report['missing_vendors'])} vendor(s)]")
        for name in report["missing_vendors"][:5]:
            print(f"  {name}")
        print()

    print("=" * W)
