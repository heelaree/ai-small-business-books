# Cash-Basis Small Business Accounting Policies

## 1. Recognition Basis
- **Cash Basis:** Record business receipts and payments on the date cash settles on the bank statement (including settlement of prior-period bills).
- **No Unverified Accruals:** Do not invent accounts receivable, accounts payable, inventory valuations, depreciation schedules, or opening equity balances from bank statements alone.
- **Stock vs. Flow Discipline:** Cumulative P&L and cash flow lines sum across months. **Never sum monthly opening or closing cash balances**—cumulative opening cash is the earliest period's opening balance, and cumulative closing cash is each account's latest supplied closing balance.

## 2. Standard Section Taxonomy

| Section | P&L Impact? | External Cash Impact? | Purpose |
| :--- | :---: | :---: | :--- |
| `Revenue` | Yes (+) | Yes | Confirmed customer sales, service receipts, commissions, less customer refunds (`-`). |
| `Direct costs` | Yes (−) | Yes | Direct cost of delivering services or goods (COGS, supplies, service subcontractors, packaging, outbound customer delivery). |
| `Operating expenses` | Yes (−) | Yes | Overhead grouped into 5 universal Summary P&L buckets: *People & Staff*, *Power, Fuel & Maintenance*, *Marketing & Promotion*, *Premises, Cleaning & Security*, and *Admin, Comms & Operations*. |
| `Taxes` | Yes (−) | Yes | Transaction VAT, stamp duties, business premises levies, and statutory remittances shown on statements. |
| `Bank fees` | Yes (−) | Yes | Account maintenance, SMS alerts, commissions, and split embedded transfer/POS processing charges. |
| `Internal transfers` | No | No (Nets to 0) | Matched same-day principal transfers between supplied accounts of the same business. |
| `Unmatched own-account transfers` | No | Yes | Transfers to/from other accounts naming the business whose statements were not supplied. |
| `Owner transactions` | No | Yes | Confirmed owner capital contributions or personal drawings. |
| `Loans` | No | Yes | Confirmed loan principal proceeds or principal repayments (interest, once separated, enters expenses). |
| `Asset purchases` | No | Yes | Confirmed long-term equipment, meter, vehicle, or capital asset purchases. |
| `Needs Review` | No (Disclosed in Bridge) | Yes | Unidentified inflows or outflows lacking narration or documentary proof of nature. |

## 3. Embedded Fee Splitting & Reversals
- Many banks/fintechs (e.g., Moniepoint, POS gateways) bundle a processing charge inside a single settlement debit or withhold a fee from a credit.
- Split each affected row into:
  - `Principal (-P)`: Gross business receipt or payment amount.
  - `Embedded fee (-F)`: Negative `Bank fees` -> `Embedded processing charges`.
- Invariant: `Principal + Embedded fee == Net Bank Settlement` (`credit - debit`) to `0.01`.
- Failed-payment reversals (e.g., reversed utility or ad debits) offset the original expense category rather than inflating `Revenue`.

## 4. Internal Transfer Matching
- Pair cross-account transfers only when **both** business bank statements are in scope, dates match (or clear within settlement window), and principal amounts match to `0.01`.
- Exclude matched principal from P&L and external cash flow; keep any transfer fees or stamp duties in `Bank fees` / `Taxes`.
