"""
research_vendors.py — Enrich vendor records with verified descriptions.

Strategy:
  - All vendors: Claude training-knowledge pass (batch of 50) → confidence HIGH/LOW
  - LOW-confidence vendors: DuckDuckGo instant-answer lookup for a snippet
  - Results cached in vendors_researched.json to survive restarts
"""
from __future__ import annotations

import json
import os
import time

import anthropic
import requests

from config import (
    MODEL,
    MAX_TOKENS,
    BATCH_SIZE,
    MAX_RETRIES,
    RETRY_DELAY,
    RESEARCHED_JSON,
    RESEARCH_THRESHOLD,
)

RESEARCH_SYSTEM_PROMPT = """\
You are a business intelligence analyst with broad knowledge of global software, \
professional services, and facilities companies.

For each vendor listed below, provide:
1. what_they_do — one sentence (max 15 words) describing the vendor's core product or service, \
   based on your training knowledge.
2. confidence — "HIGH" if you are certain this is a well-known company whose purpose you know, \
   "LOW" if the vendor is local, obscure, or you are genuinely unsure.
3. needs_human_review — true only if confidence is LOW AND the vendor has significant spend \
   (this will trigger a web lookup).

Return ONLY a JSON array, one object per vendor, in the same order as the input.
Schema: [{"vendor_name": "...", "what_they_do": "...", "confidence": "HIGH|LOW", \
"needs_human_review": true|false}]
No markdown fences, no preamble."""


def _ddg_search(vendor_name: str) -> str:
    """
    Queries DuckDuckGo Instant Answer API for a short description of the vendor.
    Returns a snippet string or empty string on failure.
    """
    try:
        resp = requests.get(
            "https://api.duckduckgo.com/",
            params={"q": vendor_name, "format": "json", "no_html": 1, "skip_disambig": 1},
            timeout=8,
        )
        if resp.status_code != 200:
            return ""
        data = resp.json()
        # AbstractText is the best signal; RelatedTopics as fallback
        snippet = data.get("AbstractText", "").strip()
        if not snippet:
            topics = data.get("RelatedTopics", [])
            if topics and isinstance(topics[0], dict):
                snippet = topics[0].get("Text", "").strip()
        return snippet[:300] if snippet else ""
    except Exception:
        return ""


def _research_batch(
    client: anthropic.Anthropic,
    batch: list[dict],
) -> list[dict]:
    """
    Sends one batch of vendors to Claude for knowledge-based research.
    Returns list of research dicts.
    """
    lines = [
        f"{i+1}. {v['vendor_name']} | ${v['cost_usd']:,.0f}"
        for i, v in enumerate(batch)
    ]
    user_content = "Research these vendors:\n\n" + "\n".join(lines)

    for attempt in range(MAX_RETRIES):
        try:
            response = client.messages.create(
                model=MODEL,
                max_tokens=MAX_TOKENS,
                system=[
                    {
                        "type": "text",
                        "text": RESEARCH_SYSTEM_PROMPT,
                        "cache_control": {"type": "ephemeral"},
                    }
                ],
                messages=[{"role": "user", "content": user_content}],
            )
            raw = response.content[0].text.strip()
            # Strip markdown fences if present
            if raw.startswith("```"):
                raw = raw.split("```")[1]
                if raw.startswith("json"):
                    raw = raw[4:]
                raw = raw.strip()
            if raw.endswith("```"):
                raw = raw[:-3].strip()
            result = json.loads(raw)
            if len(result) != len(batch):
                raise ValueError(
                    f"Expected {len(batch)} items, got {len(result)}"
                )
            return result
        except (json.JSONDecodeError, ValueError) as e:
            print(f"    Research batch parse error (attempt {attempt+1}): {e}")
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

    # Fallback: return placeholder so pipeline doesn't break
    return [
        {
            "vendor_name": v["vendor_name"],
            "what_they_do": "",
            "confidence": "LOW",
            "needs_human_review": v["cost_usd"] >= RESEARCH_THRESHOLD,
        }
        for v in batch
    ]


def research_all_vendors(vendors: list[dict]) -> list[dict]:
    """
    Runs the full research pipeline:
    1. Claude knowledge pass for all vendors (batched).
    2. DuckDuckGo web lookup for LOW-confidence, high-spend vendors.
    3. Merges results back, keyed by vendor_name.
    4. Saves to vendors_researched.json after each batch (crash recovery).

    Returns enriched vendor list: original fields + what_they_do, confidence,
    needs_human_review, web_snippet.
    """
    # Load any previously completed research
    completed: dict[str, dict] = {}
    if os.path.exists(RESEARCHED_JSON):
        with open(RESEARCHED_JSON, encoding="utf-8") as f:
            for item in json.load(f):
                completed[item["vendor_name"]] = item

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    client = anthropic.Anthropic(api_key=api_key)

    # Split into batches
    batches = [vendors[i: i + BATCH_SIZE] for i in range(0, len(vendors), BATCH_SIZE)]
    total_batches = len(batches)

    for batch_num, batch in enumerate(batches, 1):
        # Skip vendors already researched
        remaining = [v for v in batch if v["vendor_name"] not in completed]
        if not remaining:
            print(f"  Research batch {batch_num}/{total_batches}: skipped (cached)")
            continue

        print(f"  Research batch {batch_num}/{total_batches}: {len(remaining)} vendors...")
        results = _research_batch(client, remaining)

        for item in results:
            completed[item["vendor_name"]] = item

        # Crash-recovery save after each batch
        with open(RESEARCHED_JSON, "w", encoding="utf-8") as f:
            json.dump(list(completed.values()), f, indent=2, ensure_ascii=False)

    # Web lookup pass for LOW-confidence + high-spend vendors
    print("  Running web lookup for LOW-confidence / high-spend vendors...")
    changed = False
    for vendor in vendors:
        name = vendor["vendor_name"]
        rec = completed.get(name, {})
        # Lookup if: flagged for review, OR LOW confidence with spend above threshold
        needs_lookup = rec.get("needs_human_review", False) or (
            rec.get("confidence", "LOW") == "LOW"
            and vendor["cost_usd"] >= RESEARCH_THRESHOLD
        )
        if needs_lookup and not rec.get("web_snippet"):
            snippet = _ddg_search(name)
            if name not in completed:
                # Encoding mismatch between CSV and Claude's returned name — skip web update
                continue
            completed[name]["web_snippet"] = snippet
            if snippet:
                completed[name]["what_they_do"] = (
                    snippet[:120].rstrip(".") + "."
                    if len(snippet) > 120
                    else snippet
                )
                completed[name]["confidence"] = "MEDIUM"
            changed = True
            time.sleep(0.3)

    if changed:
        with open(RESEARCHED_JSON, "w", encoding="utf-8") as f:
            json.dump(list(completed.values()), f, indent=2, ensure_ascii=False)

    # Merge research back into vendor dicts
    enriched = []
    for v in vendors:
        rec = completed.get(v["vendor_name"], {})
        enriched.append({
            **v,
            "what_they_do": rec.get("what_they_do", ""),
            "confidence": rec.get("confidence", "LOW"),
            "needs_human_review": rec.get("needs_human_review", False),
            "web_snippet": rec.get("web_snippet", ""),
        })

    return enriched
