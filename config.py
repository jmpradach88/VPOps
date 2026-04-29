"""
config.py — Central constants for the vendor spend analysis pipeline.

All values here are generic and reusable across any company's vendor list.
No spreadsheet-specific data lives here — pass your input via CLI --input.
"""

# ── Model & API ───────────────────────────────────────────────────────────────
MODEL = "claude-sonnet-4-6"
MAX_TOKENS = 8192      # per API call; 4096 truncates at 50-vendor batches (~80 tokens/vendor)
BATCH_SIZE = 50        # vendors per Claude API call; 50 is the reliability sweet spot
MAX_RETRIES = 3
RETRY_DELAY = 5        # seconds; doubled on each rate-limit retry

# Vendors above this annual spend get a web research pass in addition to
# Claude's training-knowledge pass. Below this, name inference is sufficient.
RESEARCH_THRESHOLD = 20_000

# ── File paths (all generated; none are inputs) ───────────────────────────────
OUTPUT_XLSX      = "VendorAnalysis_Output.xlsx"
RAW_CSV          = "vendors_raw.csv"
CLASSIFIED_JSON  = "vendors_classified.json"
RESEARCHED_JSON  = "vendors_researched.json"
QA_JSON          = "vendors_qa.json"
INSIGHTS_JSON    = "vendors_insights.json"

# ── Department taxonomy ───────────────────────────────────────────────────────
# Standard taxonomy for a multi-product SaaS / technology company.
# Edit this list if your org uses different department names.
DEPARTMENTS = [
    "Engineering",
    "Facilities",
    "G&A",
    "Legal",
    "M&A",
    "Marketing",
    "SaaS",
    "Product",
    "Professional Services",
    "Sales",
    "Support",
    "Finance",
]
