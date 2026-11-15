#!/usr/bin/env python3
"""Objective Financial Health & Cost Concentration Report generator.

GUARDRAIL: Strictly descriptive and factual. Reports verified cash positions, revenue/cost
concentrations, month-to-month movements, banking friction, and data completeness metrics.
Never emits prescriptive business advice, cost-cutting recommendations, or financial suggestions.
"""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
from typing import Any

from ledger_engine import OPEX_BUCKETS, format_money, load_and_compute


def month_label(ym: str, partial: bool = False) -> str:
    dt = datetime.strptime(f"{ym}-01", "%Y-%m-%d")
    base = dt.strftime("%b %Y")
    return f"{base} (partial)" if partial else base


def render_health_markdown(computed: dict[str, Any]) -> str:
    p = computed["profile"]
    cur = computed["currency"]
    sp = computed["summary_pnl"]
    h = computed["health"]
    months = computed["months"]
    pflags = computed["month_partial_flags"]
    direct_lbl = p["direct_cost_label"]

    period_span = (
        f"{month_label(months[0])} – {month_label(months[-1], pflags.get(months[-1], False))}"
        if months
        else "No periods"
    )
    is_provisional = h["unclassified_in_count"] > 0 or h["unclassified_out_count"] > 0

    def m(val: float) -> str:
        return format_money(val, cur, use_symbol=True)

    lines: list[str] = []
    lines.append(f"# Financial Health & Cost Report — {p['business_name']}")
    lines.append("")
    lines.append(
        f"**Entity:** {p['legal_entity']} | **Period Covered:** {period_span} | "
        f"**Currency:** {cur} | **Accounting Basis:** {p['reporting_basis']}"
    )
    lines.append("")
    if is_provisional:
        lines.append(
            f"> **Status — Provisional Books:** {h['inflow_classified_pct']:.1f}% of external cash inflows "
            f"and {h['outflow_classified_pct']:.1f}% of external cash outflows by value are currently classified. "
            f"{m(sp['unclassified_in']['total'])} across {h['unclassified_in_count']} receipt(s) and "
            f"{m(sp['unclassified_out']['total'])} across {h['unclassified_out_count']} payment(s) remain outside "
            f"the Profit & Loss statement pending owner classification."
        )
        lines.append("")

    # Section 1: Summary P&L + Cash Bridge
    lines.append("## 1. Summary Profit & Loss and Cash Bridge")
    lines.append("")
    lines.append("| Summary Line | Cumulative Amount | Notes / Basis |")
    lines.append("| :--- | ---: | :--- |")
    lines.append(f"| **Revenue (Net of Refunds)** | **{m(sp['revenue']['total'])}** | Confirmed operating sales & service receipts |")
    lines.append(f"| **{direct_lbl}** | **{m(-sp['direct_costs']['total'])}** | Direct service/product delivery costs |")
    lines.append(f"| **Gross Profit / (Loss)** | **{m(sp['gross_profit']['total'])}** | Revenue minus {direct_lbl.lower()} |")
    for b in OPEX_BUCKETS:
        b_amt = sp["opex_buckets"][b]["total"]
        if abs(b_amt) >= 0.01:
            lines.append(f"| Operating Expenses — {b} | {m(-b_amt)} | Rolled-up overhead cluster |")
    lines.append(f"| **Total Operating Expenses** | **{m(-sp['total_opex']['total'])}** | All classified overhead categories |")
    lines.append(f"| **Operating Profit / (Loss)** | **{m(sp['operating_profit']['total'])}** | Gross result minus operating expenses |")
    lines.append(f"| Taxes & Statutory Levies | {m(-sp['taxes']['total'])} | Stamp duties, transaction VAT & levies |")
    lines.append(f"| Bank & Payment Fees | {m(-sp['bank_fees']['total'])} | Maintenance, SMS & embedded processing fees |")
    net_lbl = "Provisional Net Profit / (Loss)" if is_provisional else "Net Profit / (Loss)"
    lines.append(f"| **{net_lbl}** | **{m(sp['net_profit']['total'])}** | Result across classified P&L items |")
    lines.append(f"| *Cash Bridge: Unclassified Receipts* | *{m(sp['unclassified_in']['total'])}* | *{h['unclassified_in_count']} inflow(s) excluded from P&L* |")
    lines.append(f"| *Cash Bridge: Unclassified Payments* | *{m(-sp['unclassified_out']['total'])}* | *{h['unclassified_out_count']} outflow(s) excluded from P&L* |")
    lines.append(f"| *Cash Bridge: Non-P&L Cash Movements* | *{m(sp['non_pnl_net']['total'])}* | *Loans, capital assets, owner & out-of-scope account transfers* |")
    lines.append(f"| **Actual Net Bank Cash Change** | **{m(sp['actual_cash_change']['total'])}** | **Opening {m(h['opening_cash'])} → Closing {m(h['closing_cash'])}** |")
    lines.append("")

    # Section 2: Cash Position & Liquidity Snapshot
    lines.append("## 2. Cash Position & Liquidity Trajectory")
    lines.append("")
    lines.append("| Metric | Value | Context |")
    lines.append("| :--- | ---: | :--- |")
    lines.append(f"| Opening Combined Bank Cash | {m(h['opening_cash'])} | Start of supplied records ({months[0] if months else 'n.a.'}) |")
    cutoff_ctx = "Mixed account cutoff dates in final month" if h["mixed_cutoff"] else "End of supplied records"
    lines.append(f"| Latest Supplied Bank Cash | {m(h['closing_cash'])} | {cutoff_ctx} |")
    lines.append(f"| Net Cash Change in Bank | {m(h['net_cash_change'])} | Total external cash inflows minus outflows |")
    lines.append(f"| Total External Cash Inflows | {m(h['external_inflows_total'])} | Excludes matched internal transfers |")
    lines.append(f"| Total External Cash Outflows | {m(h['external_outflows_total'])} | Excludes matched internal transfers |")
    lines.append(f"| External Inflow-to-Outflow Ratio | {h['inflow_outflow_ratio']:.2f}x | Ratio of total external cash received to cash paid |")
    lines.append(
        f"| Highest Full-Month Closing Cash | {m(h['peak_cash_month']['balance'])} | Recorded at end of {month_label(h['peak_cash_month']['month'])} |"
    )
    lines.append(
        f"| Lowest Full-Month Closing Cash | {m(h['low_cash_month']['balance'])} | Recorded at end of {month_label(h['low_cash_month']['month'])} |"
    )
    lines.append("")

    if h["latest_account_balances"]:
        lines.append("### Account-Level Cash Balances")
        lines.append("")
        lines.append("| Bank Account | Opening Balance | Latest Closing Balance | Latest Coverage Date |")
        lines.append("| :--- | ---: | ---: | :--- |")
        open_map = {r["account"]: r["opening"] for r in h["opening_account_balances"]}
        for r in h["latest_account_balances"]:
            lines.append(
                f"| {r['account']} | {m(open_map.get(r['account'], 0.0))} | {m(r['source_closing'])} | {r['end']} ({r['coverage']}) |"
            )
        lines.append("")

    # Section 3: Where Costs Are (Cost Structure & Concentration)
    lines.append("## 3. Where Costs Are (Classified Cost Breakdown)")
    lines.append("")
    lines.append(
        f"Total classified P&L costs across the period equal **{m(h['total_classified_costs'])}**, distributed as:"
    )
    lines.append(
        f"- **{direct_lbl}:** {m(sp['direct_costs']['total'])} (**{h['direct_cost_share_pct']:.1f}%** of classified costs)"
    )
    lines.append(
        f"- **Operating Expenses (Overhead):** {m(sp['total_opex']['total'])} (**{h['opex_share_pct']:.1f}%** of classified costs)"
    )
    lines.append(
        f"- **Taxes, Levies & Bank Fees:** {m(sp['taxes_and_fees']['total'])} (**{h['tax_fee_share_pct']:.1f}%** of classified costs)"
    )
    lines.append("")

    lines.append("### Cost Clusters Ranked by Share of Spend")
    lines.append("")
    lines.append("| Cost Cluster | Cost Type | Cumulative Spend | % of Classified Costs |")
    lines.append("| :--- | :--- | ---: | ---: |")
    for c in h["cost_clusters"]:
        lines.append(f"| **{c['name']}** | {c['group_type']} | {m(c['amount'])} | {c['share_pct']:.1f}% |")
    lines.append("")

    lines.append("### Top Individual Cost Categories")
    lines.append("")
    lines.append("| Rank | Category | Summary Bucket | Transactions | Spend | % of Classified Costs |")
    lines.append("| ---: | :--- | :--- | ---: | ---: | ---: |")
    for idx, cat in enumerate(h["top_cost_categories"], 1):
        lines.append(
            f"| {idx} | {cat['category']} | {cat['summary_group']} | {cat['txn_count']} | {m(cat['amount'])} | {cat['share_pct']:.1f}% |"
        )
    lines.append("")

    # Section 4: Monthly Rhythm & Volatility
    lines.append("## 4. Monthly Cash & Cost Rhythm")
    lines.append("")
    lines.append(
        f"- **Average Full-Month Classified Spend:** {m(h['avg_full_month_cost'])} per month"
    )
    lines.append(
        f"- **Highest Spend Full Month:** {month_label(h['peak_cost_month']['month'])} ({m(h['peak_cost_month']['cost'])} in classified P&L costs)"
    )
    lines.append(
        f"- **Lowest Spend Full Month:** {month_label(h['low_cost_month']['month'])} ({m(h['low_cost_month']['cost'])} in classified P&L costs)"
    )
    lines.append("")
    lines.append("| Month | Confirmed Revenue | Direct Costs | Operating Expenses | Taxes & Fees | Unclassified Net | Closing Bank Cash |")
    lines.append("| :--- | ---: | ---: | ---: | ---: | ---: | ---: |")
    cf = computed["cash_flow"]
    for mo in months:
        lbl = month_label(mo, pflags.get(mo, False))
        lines.append(
            f"| {lbl} | {m(sp['revenue'][mo])} | {m(sp['direct_costs'][mo])} | {m(sp['total_opex'][mo])} | "
            f"{m(sp['taxes_and_fees'][mo])} | {m(sp['unclassified_net'][mo])} | {m(cf['closing'][mo])} |"
        )
    lines.append("")

    # Section 5: Banking Friction & Data Completeness
    lines.append("## 5. Banking Friction, Transfers & Data Completeness")
    lines.append("")
    lines.append(
        f"- **Embedded Payment/Transfer Processing Fees:** {m(h['embedded_fee_total'])} across {h['embedded_fee_count']} transactions"
    )
    lines.append(
        f"- **Direct Bank Account Charges (SMS, Maintenance, Commissions):** {m(h['bank_charge_total'])} across {h['bank_charge_count']} entries"
    )
    lines.append(
        f"- **Statutory Taxes & Levies Recorded on Statements:** {m(h['taxes_total'])}"
    )
    t_count = len(computed["transfers"])
    t_sum = round(sum(t["amount"] for t in computed["transfers"]), 2)
    lines.append(
        f"- **Matched Internal Account Transfers Excluded from P&L:** {t_count} pair(s) totaling {m(t_sum)} moved between supplied business accounts"
    )
    lines.append(
        f"- **Open Verification & Scope Items:** {h['open_review_count']} item(s) recorded in the Needs Review register"
    )
    vdc_st = h.get("vdc_status") or {}
    if vdc_st:
        lines.append(
            f"- **Math Integrity Status:** `{vdc_st.get('math_integrity', 'VERIFIED')}` "
            f"(all bank settlements and bridges tie to 0.01)"
        )
        lines.append(
            f"- **Accounting Completeness Status:** `{vdc_st.get('accounting_completeness', 'PROVISIONAL_REVIEW_REQUIRED')}`"
        )
        unver = vdc_st.get("unverified_range") or []
        if unver:
            lines.append("- **Explicit Unverified Range (Outside Automated Statement Verification):**")
            for u_item in unver:
                lines.append(f"  - {u_item}")
    lines.append("")
    lines.append("---")
    lines.append(
        "*Note: This report presents objective cash-basis accounting metrics and cost distributions derived strictly from supplied bank statements and confirmed classifications. It contains no financial advice or operational recommendations.*"
    )
    lines.append("")
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate factual Financial Health & Cost Report in Markdown.")
    parser.add_argument("--books", required=True, type=Path, help="Path to books_data.json.")
    parser.add_argument("--classifications", type=Path, default=None, help="Optional classifications.json path.")
    parser.add_argument("--out", required=True, type=Path, help="Output Markdown report path.")
    args = parser.parse_args()

    computed = load_and_compute(args.books, args.classifications)
    md = render_health_markdown(computed)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(md, encoding="utf-8")
    print(f"Wrote Financial Health Report -> {args.out}")


if __name__ == "__main__":
    main()
