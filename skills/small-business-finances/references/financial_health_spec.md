# Financial Health & Cost Report Specification (Strict Non-Advisory Standard)

## Purpose
Small business owners need a clear, factual mirror of **where their money came from, where costs are concentrated, how cash moved across months, and how complete their books are**—without an AI making assumptions about their operations or offering unsolicited advice.

---

## Hard Non-Advisory Guardrail ("No Suggestions")

1. **Report Facts, Proportions, Ranks, and Comparisons Only:**
   - **Allowed (Factual):** *"Power, Fuel & Maintenance totaled ₦188,608 across 27 transactions, representing 18.8% of classified P&L costs."*
   - **Allowed (Factual):** *"July recorded the highest full-month classified spend (₦290,701.85) and the lowest full-month closing bank balance (₦6,102.42)."*
   - **Allowed (Factual):** *"89.0% of external cash inflows by value (₦946,330.00 across 52 receipts) remain in Needs Review pending owner classification."*
2. **Forbidden Prescriptive / Advisory Language:**
   - **Never** tell the owner to cut, reduce, renegotiate, invest, switch vendors, raise prices, or restructure staffing.
   - **Forbidden phrases** (enforced automatically by `scripts/verify_package.py`):
     - `"you should"`, `"we recommend"`, `"recommended action"`
     - `"consider cutting"`, `"consider reducing"`, `"try to reduce"`
     - `"suggestion:"`, `"suggestions:"`, `"our advice"`

---

## Required Report Sections

1. **Summary Profit & Loss and Cash Bridge:** Concise rolled-up P&L (`Revenue` → `Direct Costs / COGS` → `Gross Profit` → 5 `Operating Expense` buckets → `Operating Profit` → `Taxes & Bank Fees` → `Net Profit`) followed immediately by the Cash Reconciliation Bridge to `Actual Net Bank Cash Change`.
2. **Cash Position & Liquidity Trajectory:** Opening cash, latest closing cash per account, net external cash change, external inflow-to-outflow ratio, highest and lowest full-month closing balances, and partial/mixed-cutoff disclosures.
3. **Where Costs Are (Cost Concentration):** Percentage split between Direct Costs, Operating Overhead, and Taxes/Bank Fees; ranked Cost Clusters; Top 10 individual expense categories with transaction counts and % share.
4. **Monthly Cash & Cost Rhythm:** Average full-month spend, highest/lowest spend full months, and a month-by-month progression table.
5. **Banking Friction, Transfers & Data Completeness:** Count and sum of embedded processing fees, direct bank charges, statutory taxes/levies, matched internal transfers excluded, and exact % of inflow/outflow value classified.
