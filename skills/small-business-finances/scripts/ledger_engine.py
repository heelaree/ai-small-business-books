#!/usr/bin/env python3
"""Core ledger aggregation, classification override, Summary P&L, and Financial Health engine.

Strictly separates factual accounting calculations from presentation layers (HTML, Markdown, Excel).
Business-agnostic: works for service studios, retail stores, agencies, clinics, and hospitality.
"""

from __future__ import annotations

import calendar
from collections import defaultdict
import json
from pathlib import Path
import re
from typing import Any

from _contract import UNVERIFIED_PROPERTIES

CURRENCY_SYMBOLS: dict[str, str] = {
    "NGN": "₦",
    "USD": "$",
    "GBP": "£",
    "EUR": "€",
    "KES": "KSh ",
    "GHS": "GH₵",
    "ZAR": "R ",
    "CAD": "CA$",
    "AUD": "A$",
}

PNL_SECTIONS = ("Revenue", "Direct costs", "Operating expenses", "Taxes", "Bank fees")

OPEX_BUCKETS = (
    "People & Staff",
    "Power, Fuel & Maintenance",
    "Marketing & Promotion",
    "Premises, Cleaning & Security",
    "Admin, Comms & Operations",
)

NON_PNL_CASH_SECTIONS = (
    "Unmatched own-account transfers",
    "Owner transactions",
    "Loans",
    "Asset purchases",
)


def format_money(amount: float, currency: str = "NGN", use_symbol: bool = True) -> str:
    """Format monetary values cleanly with accounting parentheses for negatives."""
    prefix = CURRENCY_SYMBOLS.get(currency.upper(), f"{currency.upper()} ") if use_symbol else ""
    abs_str = f"{abs(amount):,.2f}"
    if abs(amount) < 0.005:
        return f"{prefix}0.00"
    return f"({prefix}{abs_str})" if amount < 0 else f"{prefix}{abs_str}"


def infer_opex_bucket(category: str, overrides: dict[str, str] | None = None) -> str:
    """Map any Operating Expense category name into one of 5 intuitive universal buckets."""
    if overrides and category in overrides:
        return overrides[category]
    c = category.lower()
    if any(k in c for k in ("salar", "wage", "staff", "payroll", "welfare", "training", "intern", "recruit", "bonus")):
        return "People & Staff"
    if any(k in c for k in ("electric", "power", "fuel", "diesel", "generator", "solar", "inverter", "repair", "mainten")):
        return "Power, Fuel & Maintenance"
    if any(k in c for k in ("advertis", "ads", "marketing", "promo", "sponsor", "brand", "social media", "pr ")):
        return "Marketing & Promotion"
    if any(k in c for k in ("rent", "lease", "space", "clean", "securit", "water", "hygiene", "sanit", "waste", "premise")):
        return "Premises, Cleaning & Security"
    return "Admin, Comms & Operations"


def infer_summary_group(
    section: str,
    category: str,
    amount: float,
    direct_cost_label: str = "Cost of Providing Services",
    opex_overrides: dict[str, str] | None = None,
) -> str:
    """Assign each ledger component to its Summary P&L or Cash Bridge group."""
    if section == "Revenue":
        return "Revenue (Net of Refunds)"
    if section == "Direct costs":
        return direct_cost_label
    if section == "Operating expenses":
        return infer_opex_bucket(category, opex_overrides)
    if section == "Taxes":
        return "Taxes & Statutory Levies"
    if section == "Bank fees":
        return "Bank & Payment Fees"
    if section == "Needs Review":
        return "Unclassified Receipts" if amount >= 0 else "Unclassified Payments"
    if section == "Internal transfers":
        return "Matched Internal Transfers"
    return "Non-P&L Cash Movements"


def extract_counterparty_signature(
    entry: dict[str, Any],
    entity_name: str = "",
    own_account_digits: set[str] | None = None,
) -> tuple[str, str]:
    """Derive a grouping key and human-readable counterparty label for Needs Review clustering."""
    direction = "IN" if entry["amount"] >= 0 else "OUT"
    src_acct = str(entry.get("source_account") or "").strip()
    own_set = set(own_account_digits or set())
    entry_own = "".join(ch for ch in entry.get("account", "") if ch.isdigit())
    if entry_own:
        own_set.add(entry_own)

    is_external_acct = bool(src_acct and src_acct not in own_set and "*" not in src_acct)
    desc = str(entry.get("description") or "").strip()
    cp = str(entry.get("counterparty") or "").strip()
    entity_upper = entity_name.upper().strip()

    # Clean common generic prefixes before matching sender/recipient names
    cleaned_desc = re.sub(r"^Sent from Kuda Business\s*-\s*", "", desc, flags=re.IGNORECASE)

    label = ""
    patterns = [
        r"-\s*FROM\s+([A-Z][A-Z0-9 _&.,'-]{3,40})$",
        r"INWARD TRANSFER \([A-Z]\)\s+FROM\s+[A-Z0-9 ]+/\s*([A-Z0-9 _&.,'-]{3,40}?)(?:-MOBILE|/|$)",
        r"Transfer from\s+([A-Za-z][A-Za-z0-9 _&.,'-]{3,40})",
        r"FROM\s+([A-Z][A-Z0-9 _&.,'-]{3,40})",
        r"^([A-Z][A-Z0-9 _]{4,35}?)\s+-?\s*TO\s+",
        r"PURCHASE FOR\s+([A-Z0-9/ _-]{4,35}?)(?:\s+\d{4}|\s+United|\s+Zenith|/PUR|$)",
        r"PAYOUT FOR\s+([A-Z0-9 _-]{4,35}?)(?:/|$)",
        r"BANK\|([A-Z0-9 _-]{4,35}?)(?:\||/|$)",
    ]
    for pat in patterns:
        m = re.search(pat, cleaned_desc, re.IGNORECASE)
        if m:
            candidate = re.sub(r"\s+", " ", m.group(1)).strip(" -_/,.")
            if (
                candidate
                and candidate.upper() not in ("KUDA BUSINESS", "MMF", "USSD")
                and (not entity_upper or entity_upper not in candidate.upper())
            ):
                label = candidate.title()
                break

    if not label and cp and (not entity_upper or entity_upper not in cp.upper()) and not cp.isdigit():
        label = cp.title()

    if is_external_acct:
        key = f"{direction}:ACCT:{src_acct}"
        display = f"{label} (Acct {src_acct})" if label else f"Account {src_acct}"
        return key, display

    if label:
        norm = re.sub(r"[^A-Z0-9]", "", label.upper())[:24]
        return f"{direction}:NAME:{norm}", label

    short_desc = re.sub(r"\s+", " ", desc)[:52]
    return f"{direction}:TXN:{entry['id']}", short_desc or entry["id"]


def load_and_compute(
    books_data_path: Path,
    classifications_path: Path | None = None,
    parsed_path: Path | None = None,
) -> dict[str, Any]:
    """Load reconciled books_data.json, apply durable overrides, and compute all views & health metrics."""
    raw_books = json.loads(books_data_path.read_text(encoding="utf-8"))
    cls_cfg: dict[str, Any] = {}
    if classifications_path and classifications_path.exists():
        cls_cfg = json.loads(classifications_path.read_text(encoding="utf-8"))

    docs = raw_books.get("documents", [])
    valid_docs = [d for d in docs if d.get("start")]
    default_owner = next((d.get("owner") for d in valid_docs if d.get("owner")), "Small Business")
    default_currency = next((d.get("currency") for d in valid_docs if d.get("currency")), "NGN")

    profile = {
        "business_name": default_owner.title(),
        "legal_entity": default_owner,
        "business_type": "service",
        "industry": "Small Business",
        "currency": default_currency,
        "direct_cost_label": "Cost of Providing Services",
        "reporting_basis": "Cash basis",
        **(cls_cfg.get("business_profile") or {}),
    }
    currency = profile["currency"]
    direct_label = profile["direct_cost_label"]
    opex_map = cls_cfg.get("category_group_overrides") or {}
    txn_overrides: dict[str, dict[str, Any]] = cls_cfg.get("transaction_overrides") or {}
    cp_rules: list[dict[str, Any]] = cls_cfg.get("counterparty_rules") or []

    # Prepare ledger entries with applied overrides
    ledger: list[dict[str, Any]] = []
    resolved_ids: dict[str, dict[str, Any]] = {}

    own_digits_set = {
        "".join(ch for ch in d.get("account", "") if ch.isdigit())
        for d in valid_docs
        if any(ch.isdigit() for ch in d.get("account", ""))
    }

    # Pre-resolve group labels across shared external source accounts so recurring senders share one display name
    acct_known_names: dict[str, str] = {}
    for e in raw_books["ledger"]:
        _, disp = extract_counterparty_signature(e, profile["legal_entity"], own_digits_set)
        src = str(e.get("source_account") or "").strip()
        if (
            src
            and src not in own_digits_set
            and "*" not in src
            and "(Acct " in disp
            and src not in acct_known_names
        ):
            acct_known_names[src] = disp

    for orig in raw_books["ledger"]:
        e = dict(orig)
        e["month"] = e["date"][:7]
        src = str(e.get("source_account") or "").strip()
        cp_key, cp_disp = extract_counterparty_signature(e, profile["legal_entity"], own_digits_set)
        if src and src in acct_known_names:
            cp_disp = acct_known_names[src]
        e["counterparty_key"] = cp_key
        e["counterparty_display"] = cp_disp

        # Apply counterparty rule first, then explicit transaction override
        applied: dict[str, Any] | None = None
        for rule in cp_rules:
            if rule.get("counterparty_key") and rule["counterparty_key"] == cp_key:
                applied = rule
                break
            pat = str(rule.get("match_text") or "").strip().upper()
            if pat and pat in e.get("description", "").upper():
                applied = rule
                break
        if e["id"] in txn_overrides:
            applied = txn_overrides[e["id"]]

        if applied and applied.get("section"):
            e["section"] = applied["section"]
            e["category"] = applied.get("category") or e["category"]
            e["user_classified"] = True
            e["provenance"] = applied.get("provenance") or "owner_confirmed"
            e["user_note"] = applied.get("note", "Confirmed by business owner")
            e["question"] = ""
            e["review_type"] = ""
            resolved_ids[e["id"]] = applied
        else:
            e["user_classified"] = False
            e["provenance"] = orig.get("provenance") or ("unclassified_default" if e["section"] == "Needs Review" else "statement_narration_rule")
            e["user_note"] = ""

        e["summary_group"] = infer_summary_group(
            e["section"], e["category"], float(e["amount"]), direct_label, opex_map
        )
        e["pnl_impact"] = round(float(e["amount"]), 2) if e["section"] in PNL_SECTIONS else 0.0
        e["external_cash"] = 0.0 if e["section"] == "Internal transfers" else round(float(e["amount"]), 2)
        ledger.append(e)

    # Determine covered months and partial-month flags
    months = sorted({e["month"] for e in ledger} | {d["start"][:7] for d in valid_docs if d.get("start")})
    accounts = sorted({d["account"] for d in valid_docs if d.get("account")})

    # Account monthly roll-forwards
    doc_by_id = {d["id"]: d for d in docs}
    account_cash: list[dict[str, Any]] = []
    month_partial_flags: dict[str, bool] = {}

    for m in months:
        yr, mo = map(int, m.split("-"))
        last_day = f"{m}-{calendar.monthrange(yr, mo)[1]:02d}"
        first_day = f"{m}-01"
        all_end_early = True

        for acct in accounts:
            d = next(
                (
                    x
                    for x in valid_docs
                    if not x.get("is_consolidated", False)
                    and x.get("id") != "S17"
                    and x.get("account") == acct
                    and x["start"][:7] == m
                ),
                None,
            )
            alloc_cr = round(sum(max(e["amount"], 0.0) for e in ledger if e["month"] == m and e["account"] == acct), 2)
            alloc_dr = round(sum(max(-e["amount"], 0.0) for e in ledger if e["month"] == m and e["account"] == acct), 2)

            if d:
                cov_start, cov_end = d["start"], d["end"]
                opening = float(d["opening"])
                src_closing = float(d["closing"])
                src_id = d["id"]
                if cov_start == first_day and cov_end == last_day:
                    cov_note = "Full month"
                    all_end_early = False
                elif cov_end < last_day and m == months[-1]:
                    cov_note = f"Partial month through {cov_end}"
                else:
                    cov_note = f"Coverage {cov_start} to {cov_end} (missing days)"
                    all_end_early = False
            else:
                # Fallback to consolidated or prior month carry-forward
                cons = next(
                    (
                        x
                        for x in valid_docs
                        if (x.get("is_consolidated", False) or x.get("id") == "S17")
                        and x.get("account") == acct
                        and x["end"] >= first_day
                    ),
                    None,
                )
                prev_rows = [r for r in account_cash if r["account"] == acct]
                prev_close = prev_rows[-1]["source_closing"] if prev_rows else 0.0
                cov_start = first_day
                cov_end = min(cons["end"], last_day) if cons else last_day
                opening = prev_close
                src_closing = float(cons["closing"]) if (cons and m == cons["end"][:7]) else round(opening + alloc_cr - alloc_dr, 2)
                src_id = cons["id"] if cons else "Carry-forward"
                cov_note = f"Through {cov_end}; no transactions" if alloc_cr == 0 and alloc_dr == 0 else f"Through {cov_end}"

            calc_closing = round(opening + alloc_cr - alloc_dr, 2)
            account_cash.append(
                {
                    "month": m,
                    "account": acct,
                    "source_id": src_id,
                    "start": cov_start,
                    "end": cov_end,
                    "opening": round(opening, 2),
                    "allocated_credits": alloc_cr,
                    "allocated_debits": alloc_dr,
                    "calculated_closing": calc_closing,
                    "source_closing": round(src_closing, 2),
                    "closing_diff": round(calc_closing - src_closing, 2),
                    "coverage": cov_note,
                }
            )
        month_partial_flags[m] = all_end_early

    # Monthly & cumulative Summary P&L
    def sum_by_month(pred) -> dict[str, float]:
        res = {m: 0.0 for m in months}
        for e in ledger:
            if pred(e):
                res[e["month"]] = round(res[e["month"]] + e["amount"], 2)
        res["total"] = round(sum(res[m] for m in months), 2)
        return res

    rev_m = sum_by_month(lambda e: e["section"] == "Revenue")
    # Express costs as positive magnitudes for clean display alongside signed net formulas
    dc_m = {k: round(-v, 2) for k, v in sum_by_month(lambda e: e["section"] == "Direct costs").items()}
    gp_m = {k: round(rev_m[k] - dc_m[k], 2) for k in [*months, "total"]}

    opex_by_bucket: dict[str, dict[str, float]] = {}
    for b in OPEX_BUCKETS:
        raw_b = sum_by_month(lambda e, b=b: e["section"] == "Operating expenses" and e["summary_group"] == b)
        opex_by_bucket[b] = {k: round(-v, 2) for k, v in raw_b.items()}

    tot_opex_m = {k: round(sum(opex_by_bucket[b][k] for b in OPEX_BUCKETS), 2) for k in [*months, "total"]}
    op_profit_m = {k: round(gp_m[k] - tot_opex_m[k], 2) for k in [*months, "total"]}

    tax_m = {k: round(-v, 2) for k, v in sum_by_month(lambda e: e["section"] == "Taxes").items()}
    fee_m = {k: round(-v, 2) for k, v in sum_by_month(lambda e: e["section"] == "Bank fees").items()}
    tax_fee_m = {k: round(tax_m[k] + fee_m[k], 2) for k in [*months, "total"]}
    net_profit_m = {k: round(op_profit_m[k] - tax_fee_m[k], 2) for k in [*months, "total"]}

    unclass_in_m = sum_by_month(lambda e: e["section"] == "Needs Review" and e["amount"] > 0)
    unclass_out_m = {k: round(-v, 2) for k, v in sum_by_month(lambda e: e["section"] == "Needs Review" and e["amount"] < 0).items()}
    unclass_net_m = {k: round(unclass_in_m[k] - unclass_out_m[k], 2) for k in [*months, "total"]}
    non_pnl_net_m = sum_by_month(lambda e: e["section"] in NON_PNL_CASH_SECTIONS)
    ext_cash_change_m = {k: round(net_profit_m[k] + unclass_net_m[k] + non_pnl_net_m[k], 2) for k in [*months, "total"]}

    # Detailed P&L categories per section
    detailed_pnl: dict[str, list[dict[str, Any]]] = {}
    for sec in PNL_SECTIONS:
        cats = sorted({e["category"] for e in ledger if e["section"] == sec})
        rows_sec = []
        for cat in cats:
            vals = sum_by_month(lambda e, s=sec, c=cat: e["section"] == s and e["category"] == c)
            disp_vals = vals if sec == "Revenue" else {k: round(-v, 2) for k, v in vals.items()}
            rows_sec.append(
                {
                    "section": sec,
                    "category": cat,
                    "summary_group": infer_summary_group(sec, cat, -1.0 if sec != "Revenue" else 1.0, direct_label, opex_map),
                    "values": disp_vals,
                    "txn_count": sum(1 for e in ledger if e["section"] == sec and e["category"] == cat),
                }
            )
        rows_sec.sort(key=lambda x: abs(x["values"]["total"]), reverse=True)
        detailed_pnl[sec] = rows_sec

    # Cash Flow monthly roll-forward
    opening_by_m = {m: round(sum(r["opening"] for r in account_cash if r["month"] == m), 2) for m in months}
    closing_by_m = {m: round(sum(r["calculated_closing"] for r in account_cash if r["month"] == m), 2) for m in months}
    opening_by_m["total"] = opening_by_m[months[0]] if months else 0.0
    closing_by_m["total"] = closing_by_m[months[-1]] if months else 0.0

    # Needs Review register & Smart Counterparty Groups
    review_items = []
    for r in raw_books.get("review", []):
        item = dict(r)
        lid = item["ledger_id"]
        if lid in resolved_ids:
            ov = resolved_ids[lid]
            item["status"] = "Resolved"
            item["section"] = ov["section"]
            item["category"] = ov.get("category", item["category"])
            item["response"] = ov.get("note", "Classified by owner")
            item["resolution"] = f"{ov['section']} -> {item['category']}"
        else:
            item["status"] = "Open"
        review_items.append(item)

    # Sort review items by absolute amount descending so high-impact items surface first
    review_items.sort(key=lambda x: (x["status"] == "Resolved", -abs(float(x["amount"]))))

    # Build Smart Counterparty Groups for open Needs Review ledger entries
    cp_clusters: dict[str, dict[str, Any]] = {}
    for e in ledger:
        if orig_sec := next((o["section"] for o in raw_books["ledger"] if o["id"] == e["id"]), e["section"]):
            pass
        # Include originally Needs Review items so user can classify or edit their classification in UI
        is_originally_unknown = any(
            o["id"] == e["id"] and o["section"] == "Needs Review" for o in raw_books["ledger"]
        )
        if not is_originally_unknown:
            continue
        k = e["counterparty_key"]
        if k not in cp_clusters:
            cp_clusters[k] = {
                "counterparty_key": k,
                "display_name": e["counterparty_display"],
                "direction": "Receipt (Money In)" if e["amount"] >= 0 else "Payment (Money Out)",
                "total_amount": 0.0,
                "txn_ids": [],
                "dates": [],
                "sample_description": e["description"],
                "source_account": e.get("source_account", ""),
            }
        cp_clusters[k]["total_amount"] = round(cp_clusters[k]["total_amount"] + e["amount"], 2)
        cp_clusters[k]["txn_ids"].append(e["id"])
        cp_clusters[k]["dates"].append(e["date"])

    smart_groups = sorted(cp_clusters.values(), key=lambda g: abs(g["total_amount"]), reverse=True)

    # Objective Financial Health & Cost Concentration Metrics (Zero Suggestions)
    total_classified_costs = round(dc_m["total"] + tot_opex_m["total"] + tax_fee_m["total"], 2)

    # Cost clusters across ALL classified cost buckets
    cost_clusters = [
        {"name": direct_label, "group_type": "Direct Costs", "amount": dc_m["total"]},
        *[
            {"name": b, "group_type": "Operating Expenses", "amount": opex_by_bucket[b]["total"]}
            for b in OPEX_BUCKETS
            if abs(opex_by_bucket[b]["total"]) >= 0.01
        ],
        {"name": "Taxes & Statutory Levies", "group_type": "Taxes & Fees", "amount": tax_m["total"]},
        {"name": "Bank & Payment Fees", "group_type": "Taxes & Fees", "amount": fee_m["total"]},
    ]
    cost_clusters = [c for c in cost_clusters if c["amount"] > 0]
    cost_clusters.sort(key=lambda c: c["amount"], reverse=True)
    for c in cost_clusters:
        c["share_pct"] = round((c["amount"] / total_classified_costs * 100.0), 1) if total_classified_costs > 0 else 0.0

    # Top individual expense categories
    all_cost_cats = []
    for sec in ("Direct costs", "Operating expenses", "Taxes", "Bank fees"):
        for row in detailed_pnl[sec]:
            if row["values"]["total"] > 0:
                all_cost_cats.append(
                    {
                        "category": row["category"],
                        "section": sec,
                        "summary_group": row["summary_group"],
                        "amount": row["values"]["total"],
                        "txn_count": row["txn_count"],
                        "share_pct": round(row["values"]["total"] / total_classified_costs * 100.0, 1)
                        if total_classified_costs > 0
                        else 0.0,
                    }
                )
    all_cost_cats.sort(key=lambda x: x["amount"], reverse=True)

    # Monthly total classified costs
    monthly_total_costs = {
        m: round(dc_m[m] + tot_opex_m[m] + tax_fee_m[m], 2) for m in months
    }
    full_months = [m for m in months if not month_partial_flags.get(m, False)]
    cost_check_months = full_months or months
    peak_cost_month = max(cost_check_months, key=lambda m: monthly_total_costs[m])
    low_cost_month = min(cost_check_months, key=lambda m: monthly_total_costs[m])
    peak_cash_month = max(cost_check_months, key=lambda m: closing_by_m[m])
    low_cash_month = min(cost_check_months, key=lambda m: closing_by_m[m])

    # Completeness stats (external cash basis, excluding internal transfers)
    ext_in_total = round(sum(e["amount"] for e in ledger if e["section"] != "Internal transfers" and e["amount"] > 0), 2)
    ext_out_total = round(sum(-e["amount"] for e in ledger if e["section"] != "Internal transfers" and e["amount"] < 0), 2)
    classified_in = round(ext_in_total - unclass_in_m["total"], 2)
    classified_out = round(ext_out_total - unclass_out_m["total"], 2)

    # Reconstruct all_source_rows if parsed.json is supplied
    all_source_rows: list[dict[str, Any]] = []
    if parsed_path and parsed_path.exists():
        parsed_docs = json.loads(parsed_path.read_text(encoding="utf-8"))
        raw_lookup = {f"{r['statement']}:{r['seq']}": r for r in raw_books["raw"]}
        for d in parsed_docs:
            if "rows" not in d:
                continue
            running_prev = float(d.get("opening") or 0.0)
            is_excl_doc = d.get("is_consolidated", False) or d.get("id") == "S17"
            for orig_r in d["rows"]:
                matched = raw_lookup.get(f"{d['id']}:{orig_r['seq']}")
                if matched:
                    row_copy = dict(matched)
                    row_copy["included"] = 1
                else:
                    row_copy = {
                        **orig_r,
                        "balance_order": orig_r["seq"],
                        "chain_before": round(running_prev, 2),
                        "raw_id": "",
                        "included": 0 if is_excl_doc else 1,
                    }
                    running_prev = float(orig_r["balance"])
                all_source_rows.append(row_copy)
    else:
        for r in raw_books["raw"]:
            rc = dict(r)
            rc["included"] = 1
            all_source_rows.append(rc)

    encrypted_doc_count = sum(1 for d in docs if d.get("encrypted") or d.get("error"))
    unclassified_in_cnt = sum(1 for e in ledger if e["section"] == "Needs Review" and e["amount"] > 0)
    unclassified_out_cnt = sum(1 for e in ledger if e["section"] == "Needs Review" and e["amount"] < 0)
    open_review_cnt = sum(1 for r in review_items if r["status"] == "Open") + len(raw_books.get("issues", []))

    upstream_vdc = raw_books.get("vdc_report") or {}
    upstream_math = upstream_vdc.get("verdict", {}).get("math_integrity", "VERIFIED")
    completeness_verdict = (
        "VERIFIED"
        if (
            unclassified_in_cnt == 0
            and unclassified_out_cnt == 0
            and len(raw_books.get("issues", [])) == 0
            and encrypted_doc_count == 0
        )
        else "PROVISIONAL_REVIEW_REQUIRED"
    )
    unverified_range = sorted(
        set(UNVERIFIED_PROPERTIES["build_all"])
        | set(upstream_vdc.get("unverified_range", []))
        | {f"{iss[0]}: {iss[1]}" for iss in raw_books.get("issues", [])}
    )

    health = {
        "opening_cash": opening_by_m["total"],
        "closing_cash": closing_by_m["total"],
        "net_cash_change": ext_cash_change_m["total"],
        "mixed_cutoff": len({r["end"] for r in account_cash if r["month"] == months[-1]}) > 1 if months else False,
        "latest_account_balances": [r for r in account_cash if r["month"] == months[-1]] if months else [],
        "opening_account_balances": [r for r in account_cash if r["month"] == months[0]] if months else [],
        "external_inflows_total": ext_in_total,
        "external_outflows_total": ext_out_total,
        "inflow_outflow_ratio": round(ext_in_total / ext_out_total, 2) if ext_out_total > 0 else 0.0,
        "peak_cash_month": {"month": peak_cash_month, "balance": closing_by_m[peak_cash_month]},
        "low_cash_month": {"month": low_cash_month, "balance": closing_by_m[low_cash_month]},
        "total_classified_costs": total_classified_costs,
        "direct_cost_share_pct": round(dc_m["total"] / total_classified_costs * 100.0, 1) if total_classified_costs > 0 else 0.0,
        "opex_share_pct": round(tot_opex_m["total"] / total_classified_costs * 100.0, 1) if total_classified_costs > 0 else 0.0,
        "tax_fee_share_pct": round(tax_fee_m["total"] / total_classified_costs * 100.0, 1) if total_classified_costs > 0 else 0.0,
        "cost_clusters": cost_clusters,
        "top_cost_categories": all_cost_cats[:10],
        "monthly_total_costs": monthly_total_costs,
        "avg_full_month_cost": round(sum(monthly_total_costs[m] for m in cost_check_months) / len(cost_check_months), 2),
        "peak_cost_month": {"month": peak_cost_month, "cost": monthly_total_costs[peak_cost_month]},
        "low_cost_month": {"month": low_cost_month, "cost": monthly_total_costs[low_cost_month]},
        "embedded_fee_count": sum(1 for e in ledger if e["component"] == "Embedded fee"),
        "embedded_fee_total": round(sum(-e["amount"] for e in ledger if e["component"] == "Embedded fee"), 2),
        "bank_charge_count": sum(1 for e in ledger if e["section"] == "Bank fees" and e["component"] != "Embedded fee"),
        "bank_charge_total": round(
            sum(-e["amount"] for e in ledger if e["section"] == "Bank fees" and e["component"] != "Embedded fee"), 2
        ),
        "taxes_total": tax_m["total"],
        "inflow_classified_pct": round(classified_in / ext_in_total * 100.0, 1) if ext_in_total > 0 else 100.0,
        "outflow_classified_pct": round(classified_out / ext_out_total * 100.0, 1) if ext_out_total > 0 else 100.0,
        "unclassified_in_count": unclassified_in_cnt,
        "unclassified_out_count": unclassified_out_cnt,
        "open_review_count": open_review_cnt,
        "encrypted_document_count": encrypted_doc_count,
        "vdc_status": {
            "math_integrity": upstream_math,
            "accounting_completeness": completeness_verdict,
            "unverified_range": unverified_range,
        },
    }

    return {
        "profile": profile,
        "currency": currency,
        "currency_symbol": CURRENCY_SYMBOLS.get(currency.upper(), f"{currency} "),
        "months": months,
        "month_partial_flags": month_partial_flags,
        "accounts": accounts,
        "documents": docs,
        "valid_documents": valid_docs,
        "raw_transactions": raw_books["raw"],
        "all_source_rows": all_source_rows,
        "ledger": ledger,
        "duplicates": raw_books.get("duplicates", []),
        "transfers": raw_books.get("transfers", []),
        "review": review_items,
        "scope_issues": raw_books.get("issues", []),
        "smart_review_groups": smart_groups,
        "account_cash": account_cash,
        "vdc_report": upstream_vdc,
        "summary_pnl": {
            "revenue": rev_m,
            "direct_costs": dc_m,
            "gross_profit": gp_m,
            "opex_buckets": opex_by_bucket,
            "total_opex": tot_opex_m,
            "operating_profit": op_profit_m,
            "taxes": tax_m,
            "bank_fees": fee_m,
            "taxes_and_fees": tax_fee_m,
            "net_profit": net_profit_m,
            "unclassified_in": unclass_in_m,
            "unclassified_out": unclass_out_m,
            "unclassified_net": unclass_net_m,
            "non_pnl_net": non_pnl_net_m,
            "actual_cash_change": ext_cash_change_m,
        },
        "detailed_pnl": detailed_pnl,
        "cash_flow": {
            "opening": opening_by_m,
            "closing": closing_by_m,
            "sections": {
                sec: sum_by_month(lambda e, s=sec: e["section"] == s)
                for sec in (
                    *PNL_SECTIONS,
                    *NON_PNL_CASH_SECTIONS,
                    "Internal transfers",
                )
            },
            "unclassified_in": unclass_in_m,
            "unclassified_out": {k: round(-v, 2) for k, v in unclass_out_m.items()},
            "external_net_change": ext_cash_change_m,
        },
        "health": health,
        "classifications_config": cls_cfg,
    }
