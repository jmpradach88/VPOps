"""
qa_review.py — AI-based QA pass over all vendor classifications.

Every classification is reviewed by a second Claude call acting as a
senior procurement auditor. Errors trigger automatic re-classification
with the QA feedback injected as additional context.
"""
from __future__ import annotations

import json
import os
import time

import anthropic

from config import (
    MODEL,
    MAX_TOKENS,
    BATCH_SIZE,
    MAX_RETRIES,
    RETRY_DELAY,
    QA_JSON,
    CLASSIFIED_JSON,
    DEPARTMENTS,
)

QA_SYSTEM_PROMPT = """\
You are a senior procurement auditor reviewing vendor classifications produced by \
an AI assistant. Your job is to flag errors — not rewrite everything.

For each vendor classification, evaluate:
1. DEPARTMENT FIT: Is the assigned department correct given the vendor name, \
   description, and spend level? Flag if it should clearly be a different department.
2. DESCRIPTION QUALITY: Is the description specific and factual? \
   Flag if it is vague (e.g., "provides business services"), generic, or \
   potentially hallucinated (states facts that are likely wrong).
3. RECOMMENDATION CONSISTENCY:
   - "Terminate" on a vendor with spend >$10,000 requires a clear reason \
     (duplicate, defunct, one-off). Flag if the note is weak or missing.
   - "Consolidate" must name the specific duplicate in the note. \
     Flag if it just says "consolidate" with no target.
   - "Optimize" applied to a one-time purchase or a tiny vendor (<$200) \
     is suspicious — flag it.
4. FACTUAL ACCURACY: Flag any description that contradicts widely-known facts \
   (e.g., AWS described as "HR software", Salesforce described as "cloud hosting").

Severity definitions:
  "ok"   — no issues; classification looks correct
  "warn" — minor concern that should be noted but does not require re-classification
  "error"— clear misclassification or fabricated description; must be re-classified

Return ONLY a JSON array, one object per vendor, same order as input.
Schema:
[{
  "vendor_name": "...",
  "qa_passed": true | false,
  "issues": ["issue description 1", ...],
  "severity": "ok | warn | error"
}]
No markdown fences. No preamble."""


def _build_qa_prompt(batch: list[dict]) -> str:
    lines = []
    for i, item in enumerate(batch):
        lines.append(
            f"{i+1}. {item['vendor_name']} | ${item.get('cost_usd', 0):,.0f} | "
            f"Dept: {item.get('department','')} | "
            f"Rec: {item.get('recommendation','')} | "
            f"Desc: {item.get('description','')} | "
            f"Note: {item.get('recommendation_note','')}"
        )
    return "Review these vendor classifications:\n\n" + "\n".join(lines)


def _run_qa_batch(
    client: anthropic.Anthropic,
    batch: list[dict],
    batch_num: int,
    total: int,
) -> list[dict]:
    for attempt in range(MAX_RETRIES):
        try:
            response = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                system=[
                    {
                        "type": "text",
                        "text": QA_SYSTEM_PROMPT,
                        "cache_control": {"type": "ephemeral"},
                    }
                ],
                messages=[{"role": "user", "content": _build_qa_prompt(batch)}],
            )
            raw = response.content[0].text.strip()
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"):
                    raw = raw[4:]
                raw = raw.strip()
            if raw.endswith("```"):
                raw = raw[:-3].strip()
            data = json.loads(raw)
            if len(data) != len(batch):
                raise ValueError(f"Expected {len(batch)} QA items, got {len(data)}")
            return data
        except (json.JSONDecodeError, ValueError) as e:
            print(f"    QA batch {batch_num} parse error (attempt {attempt+1}): {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)
        except anthropic.RateLimitError:
            wait = RETRY_DELAY * (2**attempt)
            print(f"    Rate limited. Waiting {wait}s...")
            time.sleep(wait)
        except anthropic.APIStatusError as e:
            print(f"    API error {e.status_code}: {e.message}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)

    # Fallback: mark all as warn so pipeline continues
    return [
        {"vendor_name": item["vendor_name"], "qa_passed": True,
         "issues": ["QA batch failed; manual review recommended"], "severity": "warn"}
        for item in batch
    ]


def _reclassify_errors(
    client: anthropic.Anthropic,
    error_items: list[dict],
    qa_flags: dict[str, dict],
) -> list[dict]:
    """
    Re-classifies vendors flagged as 'error' by QA, injecting QA feedback
    as additional context into the prompt.
    """
    from classify_vendors import CLASSIFICATION_SYSTEM_PROMPT

    lines = []
    for i, item in enumerate(error_items):
        flag = qa_flags.get(item["vendor_name"], {})
        issues_str = "; ".join(flag.get("issues", []))
        lines.append(
            f"{i+1}. {item['vendor_name']} | ${item.get('cost_usd', 0):,.0f}"
            + (f" | Context: {item.get('what_they_do','')}" if item.get('what_they_do') else "")
            + f"\n   PREVIOUS CLASSIFICATION WAS WRONG — QA ISSUES: {issues_str}"
            + f"\n   Previous dept={item.get('department','')} rec={item.get('recommendation','')}"
        )
    user_content = (
        "Re-classify these vendors. The previous classifications had errors. "
        "QA issues are noted per vendor — fix them.\n\n"
        + "\n".join(lines)
    )

    for attempt in range(MAX_RETRIES):
        try:
            response = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                system=[{"type": "text", "text": CLASSIFICATION_SYSTEM_PROMPT,
                          "cache_control": {"type": "ephemeral"}}],
                messages=[{"role": "user", "content": user_content}],
            )
            raw = response.content[0].text.strip()
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"):
                    raw = raw[4:]
                raw = raw.strip()
            if raw.endswith("```"):
                raw = raw[:-3].strip()
            data = json.loads(raw)
            if len(data) != len(error_items):
                raise ValueError(f"Re-classify expected {len(error_items)}, got {len(data)}")
            return data
        except (json.JSONDecodeError, ValueError) as e:
            print(f"    Re-classify parse error (attempt {attempt+1}): {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)
        except (anthropic.RateLimitError, anthropic.APIStatusError):
            time.sleep(RETRY_DELAY)

    return []  # give up; originals kept with warn flag


def run_qa(vendors: list[dict], classifications: list[dict]) -> tuple[list[dict], dict]:
    """
    Runs the full QA pipeline.

    Returns:
      (final_classifications, qa_report)

    qa_report keys: total, ok, warn, error, reclassified, qa_results (list)
    """
    # Build a cost lookup
    cost_lookup = {v["vendor_name"]: v["cost_usd"] for v in vendors}

    # Merge cost into classification records for the QA prompt
    merged = []
    for c in classifications:
        merged.append({**c, "cost_usd": cost_lookup.get(c["vendor_name"], 0)})

    # Load cached QA results
    qa_cache: dict[str, dict] = {}
    if os.path.exists(QA_JSON):
        with open(QA_JSON, encoding="utf-8") as f:
            for item in json.load(f):
                qa_cache[item["vendor_name"]] = item

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    client = anthropic.Anthropic(api_key=api_key)

    batches = [merged[i: i + BATCH_SIZE] for i in range(0, len(merged), BATCH_SIZE)]
    total_batches = len(batches)

    for batch_num, batch in enumerate(batches, 1):
        remaining = [item for item in batch if item["vendor_name"] not in qa_cache]
        if not remaining:
            print(f"  QA batch {batch_num}/{total_batches}: skipped (cached)")
            continue

        print(f"  QA batch {batch_num}/{total_batches}: {len(remaining)} vendors...")
        results = _run_qa_batch(client, remaining, batch_num, total_batches)

        for item in results:
            qa_cache[item["vendor_name"]] = item

        with open(QA_JSON, "w", encoding="utf-8") as f:
            json.dump(list(qa_cache.values()), f, indent=2, ensure_ascii=False)

    # Identify errors that need re-classification
    error_vendors = [
        m for m in merged
        if qa_cache.get(m["vendor_name"], {}).get("severity") == "error"
    ]

    reclassified_names: set[str] = set()
    if error_vendors:
        print(f"  Re-classifying {len(error_vendors)} vendors flagged as errors by QA...")
        new_classifications = _reclassify_errors(client, error_vendors, qa_cache)
        # Update the merged list and log changes
        class_map = {c["vendor_name"]: c for c in merged}
        for new_c in new_classifications:
            name = new_c["vendor_name"]
            old = class_map.get(name, {})
            if old.get("department") != new_c.get("department") or \
               old.get("recommendation") != new_c.get("recommendation"):
                print(
                    f"    Re-classified: {name} | "
                    f"{old.get('department')}→{new_c.get('department')} | "
                    f"{old.get('recommendation')}→{new_c.get('recommendation')}"
                )
                reclassified_names.add(name)
            class_map[name] = {**old, **new_c}
        merged = list(class_map.values())

        # Update CLASSIFIED_JSON with re-classified items
        updated_classified = []
        with open(CLASSIFIED_JSON, encoding="utf-8") as f:
            existing = {item["vendor_name"]: item for item in json.load(f)}
        for item in merged:
            existing[item["vendor_name"]] = {
                k: v for k, v in item.items()
                if k in {"vendor_name", "department", "description",
                         "recommendation", "recommendation_note"}
            }
        with open(CLASSIFIED_JSON, "w", encoding="utf-8") as f:
            json.dump(list(existing.values()), f, indent=2, ensure_ascii=False)

    # Build QA report
    total = len(merged)
    ok_count = sum(1 for v in qa_cache.values() if v.get("severity") == "ok")
    warn_count = sum(1 for v in qa_cache.values() if v.get("severity") == "warn")
    error_count = sum(1 for v in qa_cache.values() if v.get("severity") == "error")

    qa_report = {
        "total": total,
        "ok": ok_count,
        "warn": warn_count,
        "error": error_count,
        "reclassified": len(reclassified_names),
        "reclassified_names": list(reclassified_names),
        "qa_results": list(qa_cache.values()),
    }

    # Return final classifications (strip cost field added for QA)
    final = []
    for item in merged:
        final.append({
            k: v for k, v in item.items()
            if k in {"vendor_name", "department", "description",
                     "recommendation", "recommendation_note"}
        })

    # Tag re-classified items so XLSX can flag them
    for item in final:
        if item["vendor_name"] in reclassified_names:
            item["qa_reclassified"] = True
        else:
            severity = qa_cache.get(item["vendor_name"], {}).get("severity", "ok")
            item["qa_reclassified"] = False
            item["qa_warn"] = severity == "warn"

    return final, qa_report
