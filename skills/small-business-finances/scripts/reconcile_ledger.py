#!/usr/bin/env python3
"""Universal bank statement deduplication, balance-chain ordering, transfer matching, and fee splitting.

Designed for the portable `small-business-finances` skill so any AI agent can pass
extracted statement JSON (`parsed_statements.json`) and receive a mathematically
reconciled `books_data.json` ready for reporting and classification.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def deduplicate_statements(documents: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Exclude overlapping consolidated statements whose rows match monthly statements."""
    valid = [d for d in documents if "rows" in d]
    monthly_docs = [d for d in valid if not d.get("is_consolidated", False) and d.get("id") != "S17"]
    consolidated_docs = [d for d in valid if d.get("is_consolidated", False) or d.get("id") == "S17"]

    raw: list[dict[str, Any]] = [r for d in monthly_docs for r in d["rows"]]

    def row_key(r: dict[str, Any]) -> tuple[Any, ...]:
        return (r["account"], r["date"], round(float(r["debit"]), 2), round(float(r["credit"]), 2), round(float(r["balance"]), 2))

    key_map: dict[tuple[Any, ...], dict[str, Any]] = {}
    for r in raw:
        k = row_key(r)
        if k not in key_map:
            key_map[k] = r

    duplicates: list[dict[str, Any]] = []
    for cdoc in consolidated_docs:
        for r in cdoc["rows"]:
            k = row_key(r)
            if k in key_map:
                kept = key_map[k]
                duplicates.append(
                    {
                        "excluded_statement": cdoc["id"],
                        "excluded_seq": r["seq"],
                        "excluded_pages": str(r.get("pages", "")),
                        "kept_statement": kept["statement"],
                        "kept_seq": kept["seq"],
                        "basis": "Unique match on account, date, debit, credit and running balance; monthly copy retained",
                    }
                )
            else:
                raw.append(r)
                key_map[k] = r

    return raw, duplicates


def chain_statement_balances(documents: list[dict[str, Any]]) -> None:
    """Restore chronological balance sequence when same-timestamp fees print out of order."""
    for doc in documents:
        if "rows" not in doc:
            continue
        current = float(doc.get("opening") or 0.0)
        ordered: list[dict[str, Any]] = []
        pending = list(doc["rows"])
        breaks: list[int] = []

        while pending:
            earliest_dt = min(r.get("datetime", r["date"]) for r in pending)
            same_dt = [r for r in pending if r.get("datetime", r["date"]) == earliest_dt]
            match = [
                r
                for r in same_dt
                if abs(round(current + float(r["credit"]) - float(r["debit"]) - float(r["balance"]), 2)) < 0.01
            ]
            if not match:
                same_day = [r for r in pending if r["date"] == earliest_dt[:10]]
                match = [
                    r
                    for r in same_day
                    if abs(round(current + float(r["credit"]) - float(r["debit"]) - float(r["balance"]), 2)) < 0.01
                ]
                candidates = match or same_dt
            else:
                candidates = match

            chosen = min(candidates, key=lambda r: int(r["seq"]))
            chosen["chain_before"] = round(current, 2)
            chosen["running_diff"] = round(current + float(chosen["credit"]) - float(chosen["debit"]) - float(chosen["balance"]), 2)
            before_val = chosen.get("before")
            chosen["source_before_diff"] = (
                None
                if before_val is None
                else round(float(before_val) + float(chosen["credit"]) - float(chosen["debit"]) - float(chosen["balance"]), 2)
            )
            chosen["balance_order"] = len(ordered) + 1
            if abs(chosen["running_diff"]) >= 0.01:
                breaks.append(int(chosen["seq"]))
            current = float(chosen["balance"])
            ordered.append(chosen)
            pending.remove(chosen)

        doc["running_breaks"] = breaks
        doc["reordered"] = sum(int(r["seq"]) != int(r["balance_order"]) for r in ordered)
        closing = float(doc.get("closing") or current)
        doc["ending_diff"] = round(current - closing, 2)


def match_internal_transfers(
    raw_rows: list[dict[str, Any]],
    entity_keywords: list[str] | None = None,
) -> list[dict[str, Any]]:
    """Pair same-day cross-account transfers between supplied business bank accounts."""
    accounts = sorted({r["account"] for r in raw_rows})
    if len(accounts) < 2:
        return []

    acct_numbers = {}
    for acct in accounts:
        digits = "".join(ch for ch in acct if ch.isdigit())
        if len(digits) >= 6:
            acct_numbers[acct] = digits

    keywords = [k.upper() for k in (entity_keywords or []) if k]
    transfers: list[dict[str, Any]] = []
    used: set[str] = set()

    debits = [r for r in raw_rows if float(r.get("debit") or 0) > 0]
    credits = [r for r in raw_rows if float(r.get("credit") or 0) > 0]

    for d_row in debits:
        if d_row["raw_id"] in used:
            continue
        d_principal = round(float(d_row["debit"]) - float(d_row.get("charge") or 0), 2)
        if d_principal <= 0:
            continue

        for c_row in credits:
            if c_row["raw_id"] in used or c_row["account"] == d_row["account"] or c_row["date"] != d_row["date"]:
                continue
            c_principal = round(float(c_row["credit"]) + float(c_row.get("charge") or 0), 2)
            if abs(d_principal - c_principal) >= 0.01:
                continue

            d_text = f"{d_row.get('description', '')} {d_row.get('counterparty', '')} {d_row.get('source_account', '')}".upper()
            c_text = f"{c_row.get('description', '')} {c_row.get('counterparty', '')} {c_row.get('source_account', '')}".upper()
            d_num = acct_numbers.get(d_row["account"], "")
            c_num = acct_numbers.get(c_row["account"], "")

            ref_d = str(d_row.get("reference") or "").strip()
            ref_c = str(c_row.get("reference") or "").strip()
            ref_match = bool(ref_d and ref_c and (ref_d in ref_c or ref_c in ref_d))
            acct_match = (d_num and d_num in c_text) or (c_num and c_num in d_text)
            entity_match = any(kw in d_text or kw in c_text for kw in keywords)

            if ref_match or acct_match or entity_match:
                tid = f"T{len(transfers) + 1:03d}"
                used.update([d_row["raw_id"], c_row["raw_id"]])
                d_row["transfer"] = tid
                c_row["transfer"] = tid
                basis = (
                    "Account identity, same date and principal amount; reference corroborates"
                    if ref_match
                    else "Account identity, same date and principal amount; bank references differ"
                )
                transfers.append(
                    {
                        "id": tid,
                        "debit_id": d_row["raw_id"],
                        "credit_id": c_row["raw_id"],
                        "date": d_row["date"],
                        "amount": d_principal,
                        "basis": basis,
                        "debit_reference": ref_d,
                        "credit_reference": ref_c,
                    }
                )
                break

    return transfers


def split_ledger_components(raw_rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Split each bank settlement into Principal (-P) and optional Embedded Fee (-F) components."""
    ledger: list[dict[str, Any]] = []
    for r in raw_rows:
        fee = round(float(r.get("charge") or 0.0), 2)
        credit = round(float(r.get("credit") or 0.0), 2)
        debit = round(float(r.get("debit") or 0.0), 2)

        section = r.get("section", "Needs Review")
        category = r.get("category", "Unidentified receipts" if credit > 0 else "Unidentified payments")
        issue = r.get("review_type", "")
        question = r.get("question", "")

        if r.get("transfer"):
            section, category, issue, question = ("Internal transfers", "Matched transfer", "", "")

        is_reversal = category == "Expense reversal" or r.get("is_expense_reversal", False)
        if is_reversal:
            amount = round(credit - fee, 2) if (credit and fee) else (credit if credit else -debit)
            fee_amt = -fee if (credit and fee) else 0.0
        else:
            amount = round((credit + fee) if credit > 0 else -(debit - fee), 2)
            fee_amt = fee

        parts = [(amount, section, category, issue, question, "Principal")]
        if abs(fee_amt) >= 0.01:
            parts.append((round(-fee_amt, 2), "Bank fees", "Embedded processing charges", "", "", "Embedded fee"))

        for idx, (amt, sec, cat, iss, quest, comp) in enumerate(parts, 1):
            suffix = "-P" if idx == 1 else "-F"
            entry = {
                **r,
                "id": f"{r['raw_id']}{suffix}",
                "amount": round(amt, 2),
                "section": sec,
                "category": cat,
                "review_type": iss,
                "question": quest,
                "component": comp,
                "transfer_id": r.get("transfer", "") if comp == "Principal" else "",
            }
            ledger.append(entry)

    net_ledger = round(sum(e["amount"] for e in ledger), 2)
    net_raw = round(sum(float(r["credit"]) - float(r["debit"]) for r in raw_rows), 2)
    if abs(net_ledger - net_raw) >= 0.01:
        raise ValueError(f"Reconciliation mismatch: ledger net ({net_ledger}) != raw bank net ({net_raw})")

    return ledger


def main() -> None:
    parser = argparse.ArgumentParser(description="Reconcile extracted bank statements into books_data.json.")
    parser.add_argument("--parsed", required=True, type=Path, help="Path to parsed statements JSON.")
    parser.add_argument("--out", required=True, type=Path, help="Output path for books_data.json.")
    parser.add_argument("--entity-keyword", action="append", default=[], help="Business name keyword(s) for transfer matching.")
    args = parser.parse_args()

    docs = json.loads(args.parsed.read_text(encoding="utf-8"))
    chain_statement_balances(docs)
    raw, duplicates = deduplicate_statements(docs)
    raw.sort(key=lambda r: (r["date"], r["account"], r.get("balance_order", r["seq"])))
    for i, r in enumerate(raw, 1):
        r["raw_id"] = f"B{i:04d}"

    owner_kw = args.entity_keyword or [d.get("owner", "") for d in docs if d.get("owner")]
    transfers = match_internal_transfers(raw, entity_keywords=owner_kw)
    ledger = split_ledger_components(raw)

    review = []
    for e in ledger:
        if e.get("question") or e.get("section") == "Needs Review":
            review.append(
                {
                    "id": f"R{len(review) + 1:03d}",
                    "ledger_id": e["id"],
                    "raw_id": e["raw_id"],
                    "date": e["date"],
                    "account": e["account"],
                    "amount": e["amount"],
                    "section": e["section"],
                    "category": e["category"],
                    "issue": e.get("review_type") or ("Receipt nature" if e["amount"] > 0 else "Payment nature"),
                    "question": e.get("question") or "Confirm business nature and supporting document.",
                    "description": e.get("description", ""),
                    "file": e.get("file", ""),
                    "pages": e.get("pages", ""),
                    "response": "",
                    "resolution": "",
                }
            )

    clean_docs = [{k: v for k, v in d.items() if k not in ("pages", "rows")} for d in docs]
    out_data = {
        "documents": clean_docs,
        "raw": raw,
        "ledger": ledger,
        "duplicates": duplicates,
        "transfers": transfers,
        "review": review,
        "issues": [],
    }
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(out_data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Reconciled {len(raw)} bank rows -> {len(ledger)} ledger components ({len(transfers)} transfer pairs).")


if __name__ == "__main__":
    main()
