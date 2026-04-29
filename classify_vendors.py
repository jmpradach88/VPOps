"""
classify_vendors.py — Classify vendors via Claude API with prompt caching.

Each vendor receives:
  - department        (one of the configured categories)
  - description       (1-line, ≤15 words)
  - recommendation    (Terminate | Consolidate | Optimize)
  - recommendation_note (1-sentence rationale)

Duplicate detection is performed by Claude itself: it is instructed to
look for vendors with similar names or overlapping services within each
batch and flag them for consolidation. No hardcoded duplicate pairs needed.
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
    CLASSIFIED_JSON,
    DEPARTMENTS,
)

_DEPT_LIST = "\n".join(f"- {d}" for d in DEPARTMENTS)

CLASSIFICATION_SYSTEM_PROMPT = f"""\
You are a VP of Operations performing a vendor spend analysis for an international \
technology company. Your task is to classify vendors from the accounts payable ledger.

DEPARTMENTS — assign exactly one per vendor:
{_DEPT_LIST}

Department definitions:
- Engineering: dev tools, cloud infrastructure, CI/CD, monitoring, code repositories
- Facilities: office space (coworking, leases), utilities, maintenance, on-site catering, \
  parking, office supplies
- G&A: HR tools, payroll, employee benefits/insurance, corporate travel management, \
  IT hardware, telecom/mobile, general admin
- Legal: law firms, legal counsel, IP/patent, compliance, notary, regulatory bodies
- M&A: investment banks, M&A advisors, due diligence firms, data room providers, \
  corporate finance advisory
- Marketing: advertising, SEO/SEM, content/PR, events, social media, brand/design agencies
- SaaS: cross-departmental business platforms (CRM, ERP, FP&A, project management, \
  workflow automation)
- Product: product management tools, UX/design platforms, user research, certification
- Professional Services: IT consulting, staff augmentation, outsourced development, \
  system integration, recruitment agencies
- Sales: sales intelligence, outbound prospecting, sales enablement, lead generation
- Support: customer support platforms, QA tools, customer success software
- Finance: accounting firms, auditors, financial advisors, tax consultants, \
  payment processing, payroll bureaus

RECOMMENDATION CRITERIA:
- Terminate: no identifiable recurring business value; one-off purchases (restaurant, \
  bakery, event venue, retail store, hotel stay); spend under $500 with no clear recurring \
  purpose; or clearly superseded by a higher-spend entry doing the same thing
- Consolidate: this vendor overlaps with another vendor on the list providing the same \
  service — name the other vendor in recommendation_note. Look for: same parent company \
  under different legal entities, same product category with redundant subscriptions, \
  duplicate travel/expense platforms, multiple office providers in the same city
- Optimize: strategic and necessary vendor, but spend is high relative to market, \
  contract should be renegotiated, volume discounts sought, or usage right-sized. \
  Apply to all recurring vendors above $10,000/year not flagged for Consolidate or Terminate.

DUPLICATE DETECTION — actively look within the provided list for:
  - Same company name in different legal forms (e.g. "Acme Ltd" and "Acme Inc")
  - Same service category with two active subscriptions (e.g. two CRM tools, \
    two travel management platforms, two cloud providers)
  - Multiple office/coworking vendors in the same geography
  When found, flag both as Consolidate and name the other in recommendation_note.

OUTPUT FORMAT — return ONLY a JSON array, one object per vendor, same order as input:
[{{
  "vendor_name": "<exact name from input>",
  "department": "<one department from the list>",
  "description": "<single sentence, max 15 words, specific and factual>",
  "recommendation": "Terminate | Consolidate | Optimize",
  "recommendation_note": "<1-sentence rationale>"
}}]
No markdown fences. No preamble. No trailing text."""


def _build_user_prompt(batch: list[dict]) -> str:
    lines = []
    for i, v in enumerate(batch):
        context = v.get("what_they_do", "")
        conf = v.get("confidence", "")
        line = f"{i+1}. {v['vendor_name']} | ${v['cost_usd']:,.0f}"
        if context:
            line += f" | Context: {context}"
        if conf == "LOW":
            line += " [LOW CONFIDENCE — use best judgement]"
        lines.append(line)
    return f"Classify these {len(batch)} vendors:\n\n" + "\n".join(lines)


def _parse_response(raw: str, expected: int) -> list[dict]:
    text = raw.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
        text = text.strip()
    if text.endswith("```"):
        text = text[:-3].strip()

    data = json.loads(text)
    if not isinstance(data, list):
        raise ValueError(f"Expected JSON array, got {type(data).__name__}")
    if len(data) != expected:
        raise ValueError(f"Expected {expected} items, got {len(data)}")

    required = {"vendor_name", "department", "description", "recommendation", "recommendation_note"}
    for i, item in enumerate(data):
        missing = required - set(item.keys())
        if missing:
            raise ValueError(f"Item {i} missing keys: {missing}")
    return data


def _classify_batch(
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
                        "text": CLASSIFICATION_SYSTEM_PROMPT,
                        "cache_control": {"type": "ephemeral"},
                    }
                ],
                messages=[{"role": "user", "content": _build_user_prompt(batch)}],
            )
            usage = response.usage
            cache_read = getattr(usage, "cache_read_input_tokens", 0)
            if cache_read > 0:
                print(f"    (cache hit: {cache_read:,} tokens saved)")
            return _parse_response(response.content[0].text, len(batch))
        except (json.JSONDecodeError, ValueError) as e:
            print(f"    Batch {batch_num} parse error (attempt {attempt+1}): {e}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)
        except anthropic.RateLimitError:
            wait = RETRY_DELAY * (2 ** attempt)
            print(f"    Rate limited. Waiting {wait}s...")
            time.sleep(wait)
        except anthropic.APIStatusError as e:
            print(f"    API error {e.status_code}: {e.message}")
            if attempt < MAX_RETRIES - 1:
                time.sleep(RETRY_DELAY)

    raise RuntimeError(f"Batch {batch_num}/{total} failed after {MAX_RETRIES} attempts")


def classify_all_vendors(vendors: list[dict]) -> list[dict]:
    """
    Classifies all vendors in batches with crash-recovery caching.
    Returns classifications in the same order as the input vendor list.
    """
    completed: dict[str, dict] = {}
    if os.path.exists(CLASSIFIED_JSON):
        with open(CLASSIFIED_JSON, encoding="utf-8") as f:
            for item in json.load(f):
                completed[item["vendor_name"]] = item

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    client = anthropic.Anthropic(api_key=api_key)

    batches = [vendors[i: i + BATCH_SIZE] for i in range(0, len(vendors), BATCH_SIZE)]
    total_batches = len(batches)

    for batch_num, batch in enumerate(batches, 1):
        remaining = [v for v in batch if v["vendor_name"] not in completed]
        if not remaining:
            print(f"  Classify batch {batch_num}/{total_batches}: skipped (cached)")
            continue

        print(f"  Classify batch {batch_num}/{total_batches}: {len(remaining)} vendors...")
        results = _classify_batch(client, remaining, batch_num, total_batches)

        for item in results:
            completed[item["vendor_name"]] = item

        with open(CLASSIFIED_JSON, "w", encoding="utf-8") as f:
            json.dump(list(completed.values()), f, indent=2, ensure_ascii=False)

    ordered = []
    for v in vendors:
        rec = completed.get(v["vendor_name"])
        if rec:
            ordered.append(rec)
        else:
            ordered.append({
                "vendor_name": v["vendor_name"],
                "department": "G&A",
                "description": "Vendor classification unavailable — manual review required.",
                "recommendation": "Optimize",
                "recommendation_note": "Could not classify; manual review required.",
            })
    return ordered
