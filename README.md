# Vendor Spend Analysis

Classifies an AP vendor ledger, identifies Top 3 cost-reduction opportunities, and
produces a CEO/CFO-ready XLSX — in a single CLI run against any company's vendor data.

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
| Top 3 Opportunities | Highest-impact savings with $ estimates, actions, timeline |
| Methodology | Pipeline docs, QA statistics, tools used, limitations |
| Recommendations | Link to Google Doc memo + summary (requires `--write-back`) |

## Pipeline

```
Fetch → Research → Classify → QA Review → Synthesize → Validate → Build XLSX
```

1. **Fetch** — loads CSV or Google Sheets, auto-detects columns
2. **Research** — Claude knowledge pass per vendor; web lookup for low-confidence / high-spend vendors
3. **Classify** — Claude API in batches of 50, prompt-cached; assigns department, description, recommendation
4. **QA Review** — second Claude pass reviews every classification; auto re-classifies errors
5. **Synthesize** — Claude generates Top 3 opportunities and executive memo from actual data
6. **Validate** — deterministic checks (coverage, valid fields, high-spend flags)

Estimated runtime: ~8 minutes for 400 vendors. Estimated API cost: ~$0.36.

## Files

```
analyze_vendors.py      Entry point
config.py               Constants (model, departments, thresholds)
fetch_data.py           CSV / Sheets loader
research_vendors.py     Vendor knowledge enrichment
classify_vendors.py     Claude classification
qa_review.py            AI QA pass
synthesize_insights.py  Opportunities + exec memo generation
build_output.py         XLSX builder
validate_output.py      Quality checks
CLAUDE.md               Project guidelines for AI-assisted development
```
