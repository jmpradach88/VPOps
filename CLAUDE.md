# CLAUDE.md — Instructions for working on this project

## What this project is

A CLI pipeline that classifies an AP vendor ledger, identifies cost-reduction opportunities,
and produces an executive-ready XLSX. It runs against any company's vendor data — nothing
is hardcoded to a specific dataset.

The target audience for the *code* is a future operator or engineer who needs to understand,
modify, or re-run it quickly. Clarity and simplicity are the primary quality bar.

---

## Coding guidelines

**Prefer clarity over cleverness.**
If a straightforward loop does the job, use it. Avoid nested comprehensions, chained
transformations, or abstractions that require mental unpacking. A future operator skimming
this at 11pm before a deadline should understand what a function does in one read.

**Keep functions small and focused.**
Each function should do one thing. If a function is doing two things, split it.
If you need a comment to explain what a block does, that block is probably a function.

**Name things honestly.**
`classify_batch` does what it says. `_build_user_prompt` builds a user prompt.
Avoid vague names like `process`, `handle`, or `run` without a subject noun.

**No premature abstraction.**
This project has one job. Don't introduce base classes, registries, plugin systems,
or factory patterns unless a concrete second use case exists. Three similar lines
is better than a premature abstraction.

**No silent failures.**
Every error path should print something useful and either raise or exit with a clear
message. Never swallow exceptions with a bare `except: pass`.

---

## Security

**Never log or print the API key.**
The `ANTHROPIC_API_KEY` is read from the environment. It must never appear in logs,
output files, print statements, or committed files.

**Never commit intermediate data files.**
`vendors_raw.csv`, `vendors_classified.json`, `vendors_researched.json`, `vendors_qa.json`,
and `vendors_insights.json` may contain real vendor spend data. They are in `.gitignore`
and must stay there. Do not add exceptions.

**Validate external inputs at the boundary.**
`fetch_data.py` is the only place that touches external data (CSV files, HTTP responses).
Column detection, cost parsing, and encoding handling all live there. Downstream code
trusts that the vendor list it receives is clean.

**No shell injection.**
If you add any `subprocess` or `os.system` calls, use argument lists — never string
interpolation into shell commands.

---

## API cost discipline

This pipeline calls the Anthropic API multiple times. Every change that touches
a Claude call should preserve or improve cost efficiency.

**Before touching any Claude call, grep for all other Claude calls and check if the same change applies.**
All API call parameters (`max_tokens`, `model`, `cache_control`) are defined once in `config.py`
and imported everywhere. If you find yourself changing a hardcoded value in one module,
run `grep -n "max_tokens\|model=\|cache_control" *.py` first to catch every instance.

**Always use prompt caching on system prompts.**
The pattern is established in every module that calls Claude:
```python
system=[{"type": "text", "text": SYSTEM_PROMPT, "cache_control": {"type": "ephemeral"}}]
```
Do not remove `cache_control`. It saves ~85% of input tokens from batch 2 onward.

**Batch size is 50 for a reason.**
50 vendors per call keeps output JSON reliably within the `MAX_TOKENS` (8,192) budget.
Going higher risks truncation and wastes a full retry. Going lower doubles API calls.
If you change `BATCH_SIZE`, test with a real run and document why.

**Don't add API calls without documenting the cost impact.**
Every new Claude call should have an estimated token count noted in a comment
or in the Methodology tab. The current pipeline costs ~$0.36 for 400 vendors.
New stages should not materially exceed that without justification.

---

## Architecture tradeoffs

**Why separate research and classify steps?**
Research establishes what each vendor *actually does* before classification happens.
This grounds the descriptions in fact rather than name inference. The tradeoff is
one extra set of API calls, but it eliminates hallucinated descriptions for
high-spend vendors — worth it.

**Why is synthesize_insights.py a separate module?**
The Top 3 opportunities and executive memo require seeing the full classified dataset
to be meaningful. Running them as a final synthesis step (rather than inline during
classification) means Claude has the complete picture when it makes recommendations.
It also makes the output fully data-driven and reusable.

**Why is write-back an opt-in flag rather than the default?**
XLSX output is self-contained, portable, and auditable — it works without any
Google credentials. Write-back (`--write-back`) requires OAuth or a service account,
creates a live Google Doc, and makes network calls that can fail. Keeping it opt-in
means the core analysis pipeline has no external credential dependency. The implementation
lives in `write_back.py` and is never imported unless `--write-back` is passed.

**Why DuckDuckGo for web lookups instead of a more capable search API?**
DuckDuckGo Instant Answer is free, requires no API key, and is sufficient for
confirming what a well-known company does. For obscure vendors, it often returns
nothing — which is the right signal (flag for human review). A paid search API
would improve coverage but adds cost and a credential dependency.

**Why crash recovery on every stage?**
A 400-vendor run takes 7–10 minutes. If it fails at batch 7 of 8, you don't want
to re-run from scratch. Every stage saves progress incrementally to a JSON cache
and resumes from the last completed batch. Do not remove this pattern.

---

## What to do when modifying this project

- If you change the department taxonomy, update `DEPARTMENTS` in `config.py` only.
  The system prompt is built from that list dynamically.
- If you add a new pipeline stage, add a corresponding `--skip-*` flag and a cache file.
- If you change the output schema from any Claude call, update the JSON schema comment
  in the relevant system prompt and verify the downstream parser still works.
- If you add a new output column to the XLSX, add it in `build_output.py` only —
  do not add display logic to other modules.
- Run `python3 -c "import config, fetch_data, research_vendors, classify_vendors,
  qa_review, synthesize_insights, build_output, validate_output, write_back, analyze_vendors"`
  after any change to confirm all modules still import cleanly.

---

## What not to do

- Do not add hardcoded vendor names, spend amounts, or company-specific logic anywhere.
  The pipeline must remain reusable across any AP dataset.
- Do not add a `requirements.txt` that pins transitive dependencies. Hard dependencies
  are `anthropic` (core pipeline) and `gspread google-auth google-auth-oauthlib
  google-api-python-client` (write-back only). Everything else (`openpyxl`, `requests`)
  ships with standard Python environments. Keep install instructions in README only.
- Do not change the batch-level JSON output format without updating the corresponding
  `_parse_response` function. Mismatched schemas are the most common source of bugs.
- Do not use `print` for debugging and leave it in. Use it for user-facing progress
  output only (`[2/6] Researching vendors...`). Remove debug prints before committing.
