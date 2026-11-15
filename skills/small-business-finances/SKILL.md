---
name: small-business-finances
description: >-
  Turn small-business bank statements (PDF or CSV) into reconciled cash-basis books,
  an interactive browser dashboard with double-click transaction drill-down and live
  in-UI classification, a concise Summary P&L (Short P&L), an accountant-ready Excel
  workbook, and a strictly factual Financial Health & Cost Report. Use whenever a
  user wants to do bookkeeping from bank statements, understand business revenue,
  costs, and profit over time, review unclassified transactions, or generate small
  business financial reports.
---

# Small-Business Finances & Interactive Bookkeeping Skill

This skill pairs **LLM judgment** (reading messy bank statement formats, grouping counterparties, and asking plain-English questions) with **deterministic Python scripts** (exact balance-chain math, transfer elimination, interactive HTML rendering, and `openpyxl` workbook generation).

---

## Core Design Principles & Hard Guardrails

1. **Progressive Disclosure for Small Business Owners:**
   Lead with clarity—*how much money came in, where costs went, what profit is, and how that bridges to bank cash*—using the **Interactive HTML Dashboard** and **Summary P&L**, while keeping full audit schedules inside the companion Excel workbook.
2. **Strict "Facts Only, Zero Suggestions" Policy:**
   In all Financial Health outputs and chat summaries, report **only objective facts** (liquidity, cost shares, monthly volatility, fee friction, and data completeness). **Never** emit unsolicited business advice, cost-cutting recommendations, or operational suggestions (e.g., never write *"You should reduce marketing spend"* or *"Consider switching from fuel to solar"*). See [references/financial_health_spec.md](./references/financial_health_spec.md).
3. **Conservative Cash-Basis Accounting:**
   Never assume an ambiguous bank inflow is revenue or an ambiguous outflow is an operating expense. Keep uncertain items in `Needs Review` and show them transparently in the **Cash Reconciliation Bridge** until the owner confirms them. See [references/accounting_policies.md](./references/accounting_policies.md).
4. **Zero Mental Math & Verification-Driven Discipline (Kajiyama, 2026):**
   Never sum transactions manually in prompt reasoning and never trust unverified LLM labels. Always:
   - Inspect the machine-readable contract (`python scripts/_contract.py --json`) and runtime (`python scripts/_contract.py doctor`).
   - Pre-measure statement files before parsing (`python scripts/probe_statements.py --dir <statements_dir> --json`), explicitly marking encrypted/unreadable PDFs as `unknown_encrypted`.
   - Enforce provenance on classifications: every non-`Needs Review` row must carry deterministic narration support or explicit owner confirmation in `classifications.json`.
   - Confirm `verify_package.py` reports `"status": "PASSED"` and `"math_integrity_verdict": "VERIFIED"`, while honoring the completeness cap (`"PROVISIONAL_REVIEW_REQUIRED"` whenever any `Needs Review` items, missing dates, or encrypted backups exist) and co-reporting `unverified_range`.
5. **Durable Owner Classifications (`classifications.json`):**
   Store every owner confirmation in `classifications.json` matching [references/classifications.schema.json](./references/classifications.schema.json) (either exported by the owner from the HTML dashboard or recorded from chat answers) so re-running scripts never overwrites user work.

---

## Prerequisites

Ensure Python 3.10+ is available with `openpyxl` (and `pdfplumber` / `pypdf` when extracting PDF bank statements):

```bash
pip install openpyxl pdfplumber pypdf
```

---

## End-to-End Workflow

### Step 0: Runtime & Statement Pre-Measurement (Probe First)
1. Check dependencies and contract:
   ```bash
   python scripts/_contract.py doctor
   python scripts/_contract.py --json
   ```
2. Pre-measure the raw statement directory before parsing:
   ```bash
   python scripts/probe_statements.py --dir <statements_dir> --json
   ```
   Record any `unknown_encrypted` or unreadable files rather than silently dropping them.

### Step 1: Intake & Business Profile Setup
1. Inventory all supplied bank/wallet statement files (PDFs, CSVs, XLSX) and note covered accounts, currencies, and date ranges.
2. Identify the business profile and choose the appropriate `direct_cost_label` from [references/industry_coa_templates.md](./references/industry_coa_templates.md):
   - Service studio / agency / clinic → `"Cost of Providing Services"`
   - Retail / e-commerce / physical goods → `"Cost of Goods Sold (COGS)"`
   - Food / bakery / hospitality → `"Cost of Food, Ingredients & Packaging"`
3. Initialize `classifications.json` if one does not yet exist:
   ```json
   {
     "schema_version": "1.0",
     "business_profile": {
       "business_name": "Example Studio",
       "legal_entity": "EXAMPLE STUDIO LTD",
       "business_type": "service",
       "industry": "Services",
       "currency": "USD",
       "direct_cost_label": "Cost of Providing Services",
       "reporting_basis": "Cash basis"
     },
     "category_group_overrides": {},
     "counterparty_rules": [],
     "transaction_overrides": {}
   }
   ```

### Step 2: Statement Extraction & Balance-Chain Reconciliation
1. Parse each readable statement into normalized JSON matching [references/parsed_statements.schema.json](./references/parsed_statements.schema.json):
   - Capture header controls (`opening`, `closing`, `credits`, `debits`, `start`, `end`, `account`, `currency`).
   - Extract each transaction (`seq`, `date`, `datetime`, `description`, `debit`, `credit`, `balance`, `charge`, `counterparty`, `source_account`, `reference`, `pages`).
   - Apply conservative rule-based `section` and `category` only where statement narrations clearly establish business purpose (`provenance: "statement_narration_rule"`); leave ambiguous items as `section: "Needs Review"`. Unverified claims without narration support are automatically downgraded by `reconcile_ledger.py`.
2. Pre-measure extracted JSON and run deterministic reconciliation (`books_data.json` conforms to [references/books_data.schema.json](./references/books_data.schema.json)):
   ```bash
   python scripts/probe_statements.py --parsed work/parsed_statements.json --json
   python scripts/reconcile_ledger.py \
     --parsed work/parsed_statements.json \
     --out work/books_data.json \
     --entity-keyword "EXAMPLE STUDIO" \
     --json
   ```
3. Inspect `vdc_report` in `books_data.json`: confirm zero `running_breaks` and zero header control deltas (`credits_diff`, `debits_diff`, `ending_diff`). Any detected encrypted backups, missing calendar days, mixed final-month cutoffs, or unmatched own-account transfers are auto-populated under `"issues"` and `"unverified_range"`.

### Step 3: Build & Verify All Deliverables
Run the single orchestrator command to generate and verify the entire 4-file output package:

```bash
python scripts/build_all.py \
  --books work/books_data.json \
  --classifications work/classifications.json \
  --parsed work/parsed_statements.json \
  --out-dir output \
  --slug Example_Studio_2026
```

`build_all.py` generates and runs automated verification (`verify_package.py`) on:
1. `output/<slug>_Financial_Dashboard.html` — Self-contained interactive visual dashboard (KPIs, monthly SVG chart, cost bars, expandable Short P&L with **double-click transaction drill-down**, Financial Health view, and an interactive **Classify Unknowns** workspace with live P&L recalculation and `classifications.json` export).
2. `output/<slug>_Financial_Health_Report.md` — Shareable, strictly factual Financial Health & Cost Concentration report (automatically scanned to enforce zero advisory language).
3. `output/<slug>_Cash_Books.xlsx` — 14-sheet formula-driven workbook ordered for small business owners first (`Executive Overview`, `Summary P&L`, `Financial Health`), followed by full accounting schedules (`Detailed P&L`, `Cash Flow`, `Master Ledger`, `Needs Review`) and audit reconciliations.
4. `output/<slug>_Cash_Ledger.csv` — UTF-8 CSV snapshot of all categorized ledger components.

### Step 4: Owner Handoff & Iterative Review Loop
1. **Present the Headline Summary & Cash Bridge** clearly in chat (showing Confirmed Revenue, Direct Costs, Gross Profit, Operating Expenses, Net Profit, and how Unclassified + Non-P&L Cash bridge to the exact Bank Cash Change).
2. **Point the owner to `<slug>_Financial_Dashboard.html`** first so they can visually explore trends, double-click any line item to inspect transactions, or tag unknown counterparties directly in the browser.
3. **Ask at most 2–3 high-impact classification questions at a time** (using the largest `smart_review_groups` by amount—often 2–3 recurring counterparties resolve 40–70% of unclassified value).
4. **Apply Owner Confirmations:** Whenever the user answers in chat or provides an exported `classifications.json` from the dashboard, merge those entries into `work/classifications.json` and re-run `scripts/build_all.py`. Verify that reconciled bank cash never changes when classifications update.
5. **Development-Time Evaluation (Optional / CI):** Run `python scripts/_contract.py check-drift` and `python scripts/probe_sensitivity.py --json` to verify schema/doc alignment and 10/10 seeded-defect sensitivity across `generate_dashboard.py`, `generate_health_report.py`, `generate_workbook.py`, `reconcile_ledger.py`, and `verify_package.py`.
