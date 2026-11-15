# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Users

- **Primary — Non-accountant small-business owners and founders** (service studios, retail stores, clinics, agencies, hospitality): They have raw monthly or consolidated bank and POS wallet statements (PDF, CSV, or XLSX, often across multiple accounts and currencies such as NGN, USD, GBP, or EUR) and need an immediate, plain-English answer to four questions—*how much money came in, where costs went, what profit they kept, and how that profit connects to their actual bank balance*—plus a fast way to label unrecognized counterparties without learning accounting jargon or navigating a 14-tab spreadsheet.
- **Secondary — Accountants, bookkeepers, and tax preparers**: They receive the companion 14-sheet formula-driven Excel workbook (`*_Cash_Books.xlsx`) and categorized CSV ledger (`*_Cash_Ledger.csv`) with full audit trails, statement-level balance reconciliations, internal transfer pair eliminations, and split embedded fees.

## Product Purpose

AI Small-Business Books (`small-business-finances`) turns raw small-business bank statements into reconciled cash-basis books and progressive-disclosure financial insights. Success means a small-business owner can open a single standalone HTML file, understand their revenue, cost structure, provisional vs. final profit, and bank cash reconciliation in seconds, label ambiguous transactions in-browser with live P&L updates, and hand off an accountant-ready package that reconciles to `0.01`.

Scope of web surfaces in this repository:
1. **Interactive Financial Overview Dashboard** (`skills/small-business-finances/templates/dashboard.html` compiled into `docs/index.html` and client `*_Financial_Dashboard.html` deliverables).
2. **Public Landing / Showcase Page** for the portable `small-business-finances` AI agent skill and live interactive demo.

## Positioning

Pairs **LLM judgment** (extracting messy statement formats, clustering recurring counterparties, and asking high-leverage plain-English classification questions) with **deterministic Python verification and reporting** (`reconcile_ledger.py`, `ledger_engine.py`, `build_all.py`, `verify_package.py`)—bridging the gap between hallucination-prone chat summaries and overwhelming accountant spreadsheets. Every number in the interactive browser dashboard, Markdown health report, and 14-sheet Excel workbook is computed from a single reconciled ledger and mathematically verified to `0.01`.

## Operating Context

- **Dual-surface workflow**: An AI coding/finance agent runs the skill pipeline on supplied bank statements and outputs a four-file package (`*_Financial_Dashboard.html`, `*_Financial_Health_Report.md`, `*_Cash_Books.xlsx`, `*_Cash_Ledger.csv`).
- **Offline & static hosting**: The HTML dashboard is opened directly from disk (`file://`) by business owners or hosted statically on GitHub Pages (`docs/index.html`).
- **Iterative owner-in-the-loop classification**: Owners inspect transactions via single-click slide-over drawers, label unknown counterparty groups in the browser (persisted in `localStorage` and exported as `classifications.json`), and feed `classifications.json` back into the Python pipeline so re-runs never overwrite owner work.

## Capabilities and Constraints

- **Single-file, zero-server HTML dashboard**: `skills/small-business-finances/templates/dashboard.html` must remain self-contained (inline CSS and vanilla JS, with ledger data embedded in `<script id="books-payload" type="application/json">{{BOOKS_JSON_PAYLOAD}}</script>`) so generated HTML files work offline over `file://` without a backend or external build step.
- **Strict "Facts Only, Zero Suggestions" guardrail**: All Financial Health & Cost Report views (HTML Tab 2 and `*_Financial_Health_Report.md`) must report strictly objective facts (liquidity, cost concentration, monthly rhythm, banking friction, and data completeness) and pass `verify_package.py`'s automated scan forbidding prescriptive or advisory phrasing (e.g., *"you should"*, *"we recommend"*, *"consider cutting"*).
- **Conservative cash-basis accounting & Cash Bridge**: Ambiguous bank inflows or outflows are never guessed as revenue or expense; they remain in `Needs Review` and are explicitly bridged between Provisional Net Profit and Actual Net Bank Cash Change (`Starting Bank Cash + Net Bank Cash Move = Closing Bank Cash`).
- **Exact `0.01` reconciliation invariance**: Reclassifying unknown transactions in the UI updates Revenue, Costs, and Profit live, but must never alter reconciled bank cash balances or actual net bank cash change.
- **Privacy guardrail**: Real client bank statements (`*Bank_Statements*/`, `/statements/`, `/data/`, `/output/`, `/tmp/`) are gitignored and must never be committed or referenced in public demo surfaces.

## Brand Commitments

- **Voice & Terminology**: Plain-English, calm, trustworthy, and non-judgmental for non-accountant business owners—leading with relatable terms (*Money In (Sales)*, *Money Out (Costs)*, *Profit Kept*, *Money in Bank (Cash)*, *Day-to-Day Operating Overhead*, *Cash Bridge*) paired with standard accounting subtitles so owners and accountants stay aligned.

## Evidence on Hand

- **Skill & Deterministic Engine**: `skills/small-business-finances/SKILL.md`, `skills/small-business-finances/scripts/reconcile_ledger.py`, `ledger_engine.py`, `generate_dashboard.py`, `generate_health_report.py`, `generate_workbook.py`, `build_all.py`, and `verify_package.py`.
- **Dashboard Template**: `skills/small-business-finances/templates/dashboard.html`.
- **Synthetic Live Demo ("Aura Studio & Spa")**: 100% synthetic two-account dataset (Jan–Jun 2026, USD, 71 bank rows / 89 ledger components, 6 matched internal transfers, 4 smart review groups) generated by `examples/generate_demo_data.py` into `docs/index.html` and `examples/output/Aura_Studio_2026_*`.
- **Absences not to fabricate**: Do not invent customer testimonials, third-party certifications, user counts, or pricing tiers unless explicitly supplied by the user.

## Product Principles

1. **Labels and Meaning Before Numbers**: Non-accountant owners must always read *what* a figure represents in plain English before being hit with raw digits; numbers support comprehension rather than overpowering the surrounding copy.
2. **Progressive Disclosure from Headline to Audit Trail**: Start with four headline answers (Money In, Money Out, Profit Kept, Money in Bank) and visual monthly/cost rhythm, then let a single click reveal subcategories, exact bank statement rows, and statement page references.
3. **Radical Honesty About Uncertainty**: Never disguise unverified bank entries inside revenue or expenses; surface unlabeled inflows and outflows visibly and show the exact mathematical bridge between provisional profit and reconciled bank cash.
4. **Facts Without Presumption**: Present cost concentration, fee friction, and cash rhythm with clinical clarity—never lecture the business owner with unsolicited operational advice.
