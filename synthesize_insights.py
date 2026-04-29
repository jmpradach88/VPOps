"""
synthesize_insights.py — Generate Top 3 opportunities and executive memo from
classified vendor data. All outputs are derived from actual analysis results,
making the pipeline fully reusable across any vendor dataset.
"""
from __future__ import annotations

import json
import os
import time
from datetime import date

import anthropic

from config import MODEL, MAX_TOKENS, MAX_RETRIES, RETRY_DELAY, INSIGHTS_JSON

INSIGHTS_SYSTEM_PROMPT = """\
You are a VP of Operations writing for a CEO and CFO audience. Produce two deliverables \
from the vendor spend data provided:

1. TOP 3 OPPORTUNITIES — the three highest-impact cost-reduction opportunities, \
   grounded in the actual data. Name the specific vendors, justify savings in USD, \
   and state what to do.

2. EXECUTIVE MEMO fields — raw data for a one-page memo. Every field has a strict \
   length constraint; stay within it or the memo will not fit on one page.

LENGTH RULES (hard limits — do not exceed):
  - opportunity title:          ≤ 7 words, action-oriented (e.g. "Audit Salesforce Licences and Renegotiate")
  - implementation_steps:       exactly 2 strings, each ≤ 12 words
  - risks:                      1 clause, ≤ 15 words, no full stop
  - immediate_actions:          exactly 4 strings, each ≤ 12 words

IMPORTANT: Every dollar figure, vendor name, and percentage must come from the data. \
Do not invent numbers or reference vendors not in the data.

Return a single JSON object with this exact schema:
{
  "analysis_date": "YYYY-MM-DD",
  "total_vendors": <int>,
  "total_spend_usd": <float>,
  "department_summary": [
    {"department": "...", "vendor_count": <int>, "total_spend": <float>}
  ],
  "recommendation_summary": {
    "Terminate": <int>, "Consolidate": <int>, "Optimize": <int>
  },
  "opportunities": [
    {
      "rank": 1,
      "title": "...",
      "description": "one or two sentence description for the XLSX detail tab",
      "affected_vendors": ["vendor1", "vendor2"],
      "current_spend_usd": <float>,
      "savings_low_usd": <float>,
      "savings_high_usd": <float>,
      "savings_rationale": "...",
      "implementation_steps": ["step one ≤12 words", "step two ≤12 words"],
      "timeline": "...",
      "risks": "single short clause ≤15 words"
    }
  ],
  "executive_memo": {
    "from": "VP of Operations",
    "immediate_actions": [
      "action one ≤12 words",
      "action two ≤12 words",
      "action three ≤12 words",
      "action four ≤12 words"
    ]
  }
}
No markdown fences. No preamble."""


def _build_synthesis_prompt(
    vendors: list[dict],
    classifications: list[dict],
    qa_report: dict | None,
) -> str:
    cost_map = {v["vendor_name"]: v["cost_usd"] for v in vendors}
    total_spend = sum(v["cost_usd"] for v in vendors)

    # Build summary stats
    dept_totals: dict[str, dict] = {}
    rec_counts: dict[str, int] = {}
    for c in classifications:
        name = c["vendor_name"]
        dept = c.get("department", "Unknown")
        rec = c.get("recommendation", "Optimize")
        cost = cost_map.get(name, 0)

        if dept not in dept_totals:
            dept_totals[dept] = {"count": 0, "spend": 0.0}
        dept_totals[dept]["count"] += 1
        dept_totals[dept]["spend"] += cost
        rec_counts[rec] = rec_counts.get(rec, 0) + 1

    # Top 30 vendors by spend with their classifications
    top_vendors = sorted(vendors, key=lambda v: v["cost_usd"], reverse=True)[:30]
    class_map = {c["vendor_name"]: c for c in classifications}
    top_lines = []
    for v in top_vendors:
        c = class_map.get(v["vendor_name"], {})
        top_lines.append(
            f"  {v['vendor_name']} | ${v['cost_usd']:,.0f} | "
            f"{c.get('department','')} | {c.get('recommendation','')} | "
            f"{c.get('description','')} | Note: {c.get('recommendation_note','')}"
        )

    # All Consolidate vendors (key for opportunity identification)
    consolidate_vendors = [
        c for c in classifications
        if c.get("recommendation") == "Consolidate"
    ]
    consolidate_lines = []
    for c in consolidate_vendors:
        cost = cost_map.get(c["vendor_name"], 0)
        consolidate_lines.append(
            f"  {c['vendor_name']} | ${cost:,.0f} | {c.get('recommendation_note','')}"
        )

    # Optimize vendors grouped by department — surfaces multi-vendor consolidation
    # opportunities that the classifier may have rated individually as Optimize
    # rather than Consolidate (e.g. two competing audit firms both flagged Optimize).
    optimize_by_dept: dict[str, list[dict]] = {}
    for c in classifications:
        if c.get("recommendation") == "Optimize":
            dept = c.get("department", "Unknown")
            optimize_by_dept.setdefault(dept, []).append(c)

    optimize_group_lines = []
    for dept, vlist in sorted(optimize_by_dept.items(),
                               key=lambda x: sum(cost_map.get(v["vendor_name"], 0) for v in x[1]),
                               reverse=True):
        if len(vlist) < 2:
            continue  # only show departments with multiple Optimize vendors
        dept_spend = sum(cost_map.get(v["vendor_name"], 0) for v in vlist)
        optimize_group_lines.append(f"  {dept} ({len(vlist)} vendors, ${dept_spend:,.0f} combined):")
        for v in sorted(vlist, key=lambda x: cost_map.get(x["vendor_name"], 0), reverse=True)[:8]:
            cost = cost_map.get(v["vendor_name"], 0)
            optimize_group_lines.append(
                f"    {v['vendor_name']} | ${cost:,.0f} | {v.get('description', '')}"
            )

    # Department spend summary
    dept_lines = sorted(dept_totals.items(), key=lambda x: x[1]["spend"], reverse=True)
    dept_summary = "\n".join(
        f"  {dept}: {data['count']} vendors, ${data['spend']:,.0f}"
        for dept, data in dept_lines
    )

    today = date.today().isoformat()

    prompt = f"""Analyze this vendor spend data and produce the Top 3 opportunities \
and executive memo as specified.

DATASET OVERVIEW
  Analysis date:    {today}
  Total vendors:    {len(vendors)}
  Total TTM spend:  ${total_spend:,.0f}
  Recommendations:  Terminate={rec_counts.get('Terminate',0)}, \
Consolidate={rec_counts.get('Consolidate',0)}, \
Optimize={rec_counts.get('Optimize',0)}

SPEND BY DEPARTMENT
{dept_summary}

TOP 30 VENDORS BY SPEND
{chr(10).join(top_lines)}

ALL CONSOLIDATE-FLAGGED VENDORS ({len(consolidate_vendors)} total)
{chr(10).join(consolidate_lines) if consolidate_lines else '  (none flagged)'}

OPTIMIZE VENDORS GROUPED BY DEPARTMENT
(Multiple Optimize vendors in the same department often signal a consolidation \
opportunity even if not individually flagged as Consolidate. Review for overlap.)
{chr(10).join(optimize_group_lines) if optimize_group_lines else '  (none)'}

Use this data to identify the three highest-impact opportunities. \
Prioritize by dollar impact. Consolidation opportunities across multiple Optimize \
vendors in the same department are valid — treat them as Consolidate opportunities \
if the vendors clearly provide overlapping services. Be specific about which vendors \
to act on and why."""

    return prompt


def synthesize(
    vendors: list[dict],
    classifications: list[dict],
    qa_report: dict | None = None,
) -> dict:
    """
    Generates Top 3 opportunities and executive memo from the classified data.
    Caches result to INSIGHTS_JSON to avoid re-running on subsequent builds.
    """
    if os.path.exists(INSIGHTS_JSON):
        with open(INSIGHTS_JSON, encoding="utf-8") as f:
            cached = json.load(f)
        # Invalidate cache if vendor count changed
        if cached.get("total_vendors") == len(vendors):
            print("  Loaded insights from cache.")
            return cached

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    client = anthropic.Anthropic(api_key=api_key)
    prompt = _build_synthesis_prompt(vendors, classifications, qa_report)

    for attempt in range(MAX_RETRIES):
        try:
            response = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                system=[
                    {
                        "type": "text",
                        "text": INSIGHTS_SYSTEM_PROMPT,
                        "cache_control": {"type": "ephemeral"},
                    }
                ],
                messages=[{"role": "user", "content": prompt}],
            )
            raw = response.content[0].text.strip()
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"):
                    raw = raw[4:]
                raw = raw.strip()
            if raw.endswith("```"):
                raw = raw[:-3].strip()

            insights = json.loads(raw)

            with open(INSIGHTS_JSON, "w", encoding="utf-8") as f:
                json.dump(insights, f, indent=2, ensure_ascii=False)

            return insights

        except (json.JSONDecodeError, ValueError) as e:
            print(f"  Synthesis parse error (attempt {attempt+1}): {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)
        except anthropic.RateLimitError:
            wait = RETRY_DELAY * (2 ** attempt)
            print(f"  Rate limited. Waiting {wait}s...")
            time.sleep(wait)
        except anthropic.APIStatusError as e:
            print(f"  API error {e.status_code}: {e.message}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)

    raise RuntimeError("Insight synthesis failed after all retries.")
