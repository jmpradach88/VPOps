# Vendor Spend Analysis

Classifies an AP vendor ledger, identifies Top 3 cost-reduction opportunities, and
produces a CEO/CFO-ready XLSX — in a single CLI run against any company's vendor data.

Built with the **Claude Code CLI** using the Anthropic SDK (`claude-sonnet-4-6`).
All classification, QA review, and synthesis is performed programmatically via the API —
no manual data entry.

**Input:** [`vendors_raw.csv`](vendors_raw.csv) — AP ledger (386 vendors, $7.9M TTM spend)  
**Source sheet:** [Google Sheets](https://docs.google.com/spreadsheets/d/1L2u8j-3cSFLPMXbbBljI4YYXQ9COYS5YihtBhxWTPls)  
**Output:** [`VendorAnalysis_Output.xlsx`](VendorAnalysis_Output.xlsx) — classified vendor data, Top 3 opportunities, methodology, memo link

## Setup

```bash
pip install anthropic gspread google-auth google-auth-oauthlib google-api-python-client
export ANTHROPIC_API_KEY=sk-ant-...
```

`openpyxl` and `requests` are included in standard Anaconda/Python installs.

The Google Sheets/Docs packages (`gspread`, `google-auth`, `google-api-python-client`) are only
required if you use the `--write-back` flag. The core analysis pipeline works without them.

## Usage

```bash
# CSV file
python3 analyze_vendors.py --input vendors.csv

# Google Sheets URL or ID
python3 analyze_vendors.py --input https://docs.google.com/spreadsheets/d/YOUR_ID/...
python3 analyze_vendors.py --input YOUR_SHEET_ID

# Re-run and skip expensive stages using cached files
python3 analyze_vendors.py --input vendors.csv --skip-research --skip-classify --skip-qa

# Write results back to the source Google Sheet (requires GOOGLE_CREDENTIALS_FILE)
# Also creates a Google Doc for the executive memo and links it in the Recommendations tab
export GOOGLE_CREDENTIALS_FILE=/path/to/credentials.json
python3 analyze_vendors.py --input YOUR_SHEET_ID --skip-research --skip-classify --skip-qa --skip-synthesis --write-back
```

## Input requirements

Any CSV with at minimum:
- A vendor/supplier name column
- A spend/cost amount column (USD)

Column names are auto-detected — no preprocessing required.

## Output

`VendorAnalysis_Output.xlsx` with four tabs:

| Tab | Contents |
|-----|----------|
| Vendor Analysis | Every vendor: Department, Description, Recommendation, QA Flag |
| Top 3 Opportunities | Highest-impact savings with $ estimates, affected vendors, timeline, risks |
| Methodology | Pipeline docs, prompts used, QA evidence with re-classification examples |
| Recommendations | Google Doc link (requires `--write-back`) |

With `--write-back`, results are also written to the source Google Sheet and a
one-page executive memo is created as a shared Google Doc (anyone with link can view).

## Pipeline

```
Fetch → Research → Classify → QA Review → Synthesize → Validate → Build XLSX
```

1. **Fetch** — loads CSV or Google Sheets, auto-detects columns
2. **Research** — Claude knowledge pass per vendor (batches of 50); DuckDuckGo web lookup for
   low-confidence vendors above $20K spend
3. **Classify** — Claude API with prompt caching (`cache_control: ephemeral`); assigns department,
   one-line description, and Terminate / Consolidate / Optimize with justification note
4. **QA Review** — second independent Claude pass audits every classification against four criteria:
   department fit, description specificity, recommendation consistency, factual accuracy.
   Errors trigger automatic re-classification with QA feedback injected as context.
5. **Synthesize** — Claude analyzes the full classified dataset to generate Top 3 opportunities
   and executive memo from actual data; no hardcoded outputs
6. **Validate** — deterministic rule checks: 100% coverage, valid fields, Consolidate notes name
   a target, high-spend Terminate flags surfaced for human review

Estimated runtime: ~8 minutes for 400 vendors. Estimated API cost: ~$0.36.

## Files

```
analyze_vendors.py      Entry point — orchestrates all pipeline stages
config.py               Constants: model, departments, thresholds, file paths
fetch_data.py           CSV / Sheets loader; column auto-detection
research_vendors.py     Vendor knowledge enrichment via Claude + DuckDuckGo
classify_vendors.py     Claude classification with prompt caching + crash recovery
qa_review.py            AI QA audit pass; auto re-classification of errors
synthesize_insights.py  Top 3 opportunities + executive memo generation
build_output.py         XLSX builder (4 tabs)
validate_output.py      Deterministic rule-based quality checks
write_back.py           Google Sheets write-back + Google Doc memo creation
CLAUDE.md               Project guidelines for AI-assisted development
```

## Intermediate files (gitignored — may contain vendor spend data)

```
vendors_raw.csv             Raw AP data from source
vendors_researched.json     Research cache
vendors_classified.json     Classification cache (crash recovery)
vendors_qa.json             QA results cache
vendors_insights.json       Synthesized opportunities + memo cache
```
