#!/usr/bin/env python3
"""Universal bank statement deduplication, balance-chain ordering, transfer matching, and fee splitting.

Implements Layer 1 Verification-Driven Coding (Kajiyama, 2026):
  - Verifies every readable statement against its header controls (credits, debits, closing) and running balance chain.
  - Enforces classification provenance so ungrounded P&L tags are downgraded to Needs Review.
  - Auto-populates statement-level unknown/coverage issues (encrypted backups, date gaps, mixed cutoffs).
  - Emits a structured VDC report (deltas, unverified_range, verdict, attribution) inside books_data.json.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import re
import sys
from typing import Any

from _contract import BALANCE_TOLERANCE, UNVERIFIED_PROPERTIES, make_delta, make_vdc_report
from probe_statements import detect_coverage_gaps

NARRATION_RULES: dict[str, tuple[str, ...]] = {
    "Revenue": (
        "pos", "sale", "sales", "service", "services", "client", "settlement", "invoice",
        "lashes", "leg wash", "refund", "commission", "booking", "customer", "retail", "bridal",
        "facial", "peel", "brow", "treatment", "skincare", "menkiti",
    ),
    "Direct costs": (
        "suppl", "contractor", "technician", "tech", "packag", "packing", "deliver", "cogs",
        "wholesale", "lash", "brow", "pedicure", "nail", "bail tech", "inventory", "stock",
    ),
    "Operating expenses": (
        "rent", "lease", "space", "power", "electric", "nepa", "fuel", "diesel", "generator",
        "ads", "meta", "facebook", "instagram", "promo", "sponsor", "marketing", "salary",
        "salaries", "salty", "payroll", "staff", "welfare", "clean", "detergent", "toiletries",
        "repair", "mainten", "sumo", "internet", "telecom", "fiber", "phone", "airtime",
        "water", "security", "software", "refreshment", "juice", "transport", "t.fare",
        "printing", "training", "intern", "servicing", "hvac",
    ),
    "Taxes": (
        "vat", "value added tax", "stamp duty", "tax", "levy", "license", "premises", "statutory",
    ),
    "Bank fees": (
        "fee", "charge", "commission", "sms", "maintenance", "misc.", "processing",
    ),
    "Asset purchases": (
        "equipment", "machine", "meter", "asset", "laptop", "furniture", "device", "vehicle",
    ),
    "Loans": (
        "loan", "repayment", "credit facility",
    ),
    "Unmatched own-account transfers": (
        "zenith", "own-account", "own account", "remz aesthetics", "transfer", "trf",
    ),
    "Owner transactions": (
        "owner", "draw", "capital", "equity", "personal",
    ),
}


def has_deterministic_narration_support(row: dict[str, Any], section: str) -> bool:
    """Return True if row narration/counterparty contains deterministic keywords supporting `section`."""
    keywords = NARRATION_RULES.get(section, ())
    if not keywords:
        return False
    haystack = f"{row.get('description', '')} {row.get('counterparty', '')}".lower()
    for kw in keywords:
        if len(kw) <= 3:
            if re.search(rf"\b{re.escape(kw)}\b", haystack):
                return True
        elif kw in haystack:
            return True
    return False


def deduplicate_statements(documents: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]], dict[str, int]]:
    """Exclude overlapping consolidated statements whose rows match monthly statements."""
    valid = [d for d in documents if "rows" in d and not d.get("error")]
    monthly_docs = [d for d in valid if not d.get("is_consolidated", False) and d.get("id") != "S17"]
    consolidated_docs = [d for d in valid if d.get("is_consolidated", False) or d.get("id") == "S17"]

    raw: list[dict[str, Any]] = [r for d in monthly_docs for r in d["rows"]]

    def row_key(r: dict[str, Any]) -> tuple[Any, ...]:
        return (
            r["account"],
            r["date"],
            round(float(r["debit"]), 2),
            round(float(r["credit"]), 2),
            round(float(r["balance"]), 2),
        )

    key_map: dict[tuple[Any, ...], dict[str, Any]] = {}
    for r in raw:
        k = row_key(r)
        if k not in key_map:
            key_map[k] = r

    duplicates: list[dict[str, Any]] = []
    consolidated_matched = 0
    consolidated_unmatched = 0
    for cdoc in consolidated_docs:
        for r in cdoc["rows"]:
            k = row_key(r)
            if k in key_map:
                kept = key_map[k]
                consolidated_matched += 1
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
                consolidated_unmatched += 1
                raw.append(r)
                key_map[k] = r

    stats = {
        "consolidated_matched": consolidated_matched,
        "consolidated_unmatched": consolidated_unmatched,
    }
    return raw, duplicates, stats


def chain_statement_balances(documents: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Restore chronological balance sequence and verify statement header controls."""
    deltas: list[dict[str, Any]] = []

    for doc in documents:
        if "rows" not in doc or doc.get("error"):
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
            chosen["running_diff"] = round(
                current + float(chosen["credit"]) - float(chosen["debit"]) - float(chosen["balance"]), 2
            )
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
        closing = doc.get("closing")
        doc["ending_diff"] = round(current - float(closing), 2) if closing is not None else 0.0

        row_credits = round(sum(float(r["credit"]) for r in ordered), 2)
        row_debits = round(sum(float(r["debit"]) for r in ordered), 2)
        doc_id = doc["id"]

        deltas.append(
            make_delta(
                f"stmt.{doc_id}.running_breaks",
                f"{doc_id} running balance breaks",
                0,
                len(breaks),
                provenance="input_premeasurement",
                tolerance=0.0,
                notes=f"Broken seqs: {breaks}" if breaks else "",
            )
        )
        deltas.append(
            make_delta(
                f"stmt.{doc_id}.ending_diff",
                f"{doc_id} chained closing vs header closing",
                closing,
                round(current, 2),
                provenance="statement_header_control",
                tolerance=BALANCE_TOLERANCE,
            )
        )
        deltas.append(
            make_delta(
                f"stmt.{doc_id}.credits_diff",
                f"{doc_id} extracted credits vs header credits",
                doc.get("credits"),
                row_credits,
                provenance="statement_header_control",
                tolerance=BALANCE_TOLERANCE,
            )
        )
        deltas.append(
            make_delta(
                f"stmt.{doc_id}.debits_diff",
                f"{doc_id} extracted debits vs header debits",
                doc.get("debits"),
                row_debits,
                provenance="statement_header_control",
                tolerance=BALANCE_TOLERANCE,
            )
        )

    return deltas


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


def split_ledger_components(raw_rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], int]:
    """Split each bank settlement into Principal (-P) and optional Embedded Fee (-F) components.

    Enforces classification provenance: any non-Needs-Review tag without explicit provenance or
    deterministic narration keyword support is downgraded to Needs Review.
    """
    ledger: list[dict[str, Any]] = []
    downgraded_count = 0

    for r in raw_rows:
        fee = round(float(r.get("charge") or 0.0), 2)
        credit = round(float(r.get("credit") or 0.0), 2)
        debit = round(float(r.get("debit") or 0.0), 2)

        section = r.get("section", "Needs Review")
        category = r.get("category", "Unidentified receipts" if credit > 0 else "Unidentified payments")
        issue = r.get("review_type", "")
        question = r.get("question", "")
        provenance = r.get("provenance", "")

        if r.get("transfer"):
            section, category, issue, question = ("Internal transfers", "Matched transfer", "", "")
            provenance = "matched_transfer_pair"
        elif section == "Needs Review":
            provenance = provenance or "unclassified_default"
        else:
            if provenance in ("owner_confirmed", "statement_narration_rule", "matched_transfer_pair", "synthetic_fixture"):
                pass
            elif has_deterministic_narration_support(r, section):
                provenance = "statement_narration_rule"
            else:
                # Downgrade ungrounded agent claim to Needs Review
                section = "Needs Review"
                category = "Unidentified receipts" if credit > 0 else "Unidentified payments"
                issue = "Unverified classification"
                question = (
                    "Automated classification lacked narration support or owner provenance. "
                    "Confirm business purpose and supporting document."
                )
                provenance = "downgraded_unverified_claim"
                downgraded_count += 1

        is_reversal = category == "Expense reversal" or r.get("is_expense_reversal", False)
        if is_reversal:
            amount = round(credit - fee, 2) if (credit and fee) else (credit if credit else -debit)
            fee_amt = -fee if (credit and fee) else 0.0
        else:
            amount = round((credit + fee) if credit > 0 else -(debit - fee), 2)
            fee_amt = fee

        parts = [(amount, section, category, issue, question, "Principal", provenance)]
        if abs(fee_amt) >= 0.01:
            parts.append((round(-fee_amt, 2), "Bank fees", "Embedded processing charges", "", "", "Embedded fee", "embedded_fee_split"))

        for idx, (amt, sec, cat, iss, quest, comp, prov) in enumerate(parts, 1):
            suffix = "-P" if idx == 1 else "-F"
            entry = {
                **r,
                "id": f"{r['raw_id']}{suffix}",
                "amount": round(amt, 2),
                "section": sec,
                "category": cat,
                "provenance": prov,
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

    return ledger, downgraded_count


def build_scope_issues(docs: list[dict[str, Any]], ledger: list[dict[str, Any]]) -> list[list[str]]:
    """Auto-populate statement-level scope and unknown coverage issues."""
    coverage = detect_coverage_gaps(docs)
    issues: list[list[str]] = []

    if coverage["encrypted_files"]:
        names = ", ".join(Path(x["file"]).name for x in coverage["encrypted_files"][:4])
        extra = f" (+{len(coverage['encrypted_files']) - 4} more)" if len(coverage["encrypted_files"]) > 4 else ""
        issues.append(
            [
                "Backups / Unreadable",
                f"{len(coverage['encrypted_files'])} password-protected or unreadable file(s) could not be verified: {names}{extra}. State is unknown_encrypted.",
                "Supply decrypted copies or passwords if content verification against monthly statements is required.",
            ]
        )

    for gap in coverage["missing_day_gaps"]:
        issues.append(
            [
                "Coverage Gap",
                f"{gap['account']} statement {gap['statement_id']} omits {gap['gap_start']} to {gap['gap_end']} ({gap['missing_days']} missing day(s)).",
                f"Supply statement coverage for {gap['account']} covering {gap['gap_start']} to {gap['gap_end']}.",
            ]
        )

    if coverage["mixed_cutoff"]:
        cut_desc = "; ".join(f"{acct} through {dt}" for acct, dt in sorted(coverage["latest_account_cutoffs"].items()))
        issues.append(
            [
                "Coverage Cutoff",
                f"Accounts end on different dates in the final month ({cut_desc}); combined closing cash is not a common-date position.",
                "Supply statements aligned to a common period-end date.",
            ]
        )

    unmatched_own = [e for e in ledger if e["section"] == "Unmatched own-account transfers"]
    if unmatched_own:
        issues.append(
            [
                "Scope / Missing Accounts",
                f"{len(unmatched_own)} transaction(s) reference own-account transfers to/from bank accounts not supplied in this package.",
                "Supply statements for all referenced business accounts to eliminate cross-account transfers.",
            ]
        )

    return issues


def reconcile_documents(
    docs: list[dict[str, Any]],
    entity_keywords: list[str] | None = None,
    allow_unresolved: bool = False,
) -> dict[str, Any]:
    """Run full reconciliation and return books_data dictionary with VDC report."""
    stmt_deltas = chain_statement_balances(docs)
    raw, duplicates, dedup_stats = deduplicate_statements(docs)
    raw.sort(key=lambda r: (r["date"], r["account"], r.get("balance_order", r["seq"])))
    for i, r in enumerate(raw, 1):
        r["raw_id"] = f"B{i:04d}"

    owner_kw = entity_keywords or [d.get("owner", "") for d in docs if d.get("owner")]
    transfers = match_internal_transfers(raw, entity_keywords=owner_kw)
    ledger, downgraded_count = split_ledger_components(raw)

    net_raw = round(sum(float(r["credit"]) - float(r["debit"]) for r in raw), 2)
    net_ledger = round(sum(float(e["amount"]) for e in ledger), 2)
    stmt_deltas.append(
        make_delta(
            "reconcile.ledger_vs_raw_net",
            "split ledger net vs raw bank net",
            net_raw,
            net_ledger,
            provenance="input_premeasurement",
            tolerance=BALANCE_TOLERANCE,
        )
    )
    for t in transfers:
        pair_net = round(sum(float(e["amount"]) for e in ledger if e.get("transfer_id") == t["id"]), 2)
        stmt_deltas.append(
            make_delta(
                f"reconcile.transfer.{t['id']}",
                f"transfer {t['id']} net pair",
                0.0,
                pair_net,
                provenance="matched_transfer_pair",
                tolerance=BALANCE_TOLERANCE,
            )
        )
    stmt_deltas.append(
        make_delta(
            "reconcile.downgraded_unverified_claims",
            "ungrounded classifications downgraded to Needs Review",
            0,
            downgraded_count,
            provenance="specification",
            tolerance=0.0,
            notes="Non-zero means agent-proposed P&L tags lacked narration support or owner provenance and were safely quarantined.",
        )
    )

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

    issues = build_scope_issues(docs, ledger)

    # Check for hard reconciliation failures (exclude informational downgraded count from math failure)
    math_failures = [
        d for d in stmt_deltas if d["status"] == "fail" and d["check_id"] != "reconcile.downgraded_unverified_claims"
    ]
    unverified = list(UNVERIFIED_PROPERTIES["reconcile_ledger"])
    for iss in issues:
        unverified.append(f"{iss[0]}: {iss[1]}")
    if dedup_stats["consolidated_unmatched"]:
        unverified.append(f"consolidated_unmatched_rows_retained:{dedup_stats['consolidated_unmatched']}")

    has_unknowns = bool(review or issues or any(d["status"] == "unknown" for d in stmt_deltas))
    math_verdict = "UNRESOLVED" if math_failures else "VERIFIED"
    comp_verdict = "PROVISIONAL_REVIEW_REQUIRED" if has_unknowns else "VERIFIED"

    vdc_report = make_vdc_report(
        deltas=stmt_deltas,
        unverified_range=unverified,
        math_integrity=math_verdict,
        accounting_completeness=comp_verdict,
        attribution_state="unresolved" if math_failures else "none",
        attribution_grounds=(
            f"{len(math_failures)} statement/ledger arithmetic check(s) exceeded tolerance."
            if math_failures
            else "All readable statement running balances, header controls, transfers, and fee splits reconciled."
        ),
        next_measurement=(
            "Compare extracted statement rows against source PDF pages for failing statement IDs."
            if math_failures
            else ""
        ),
    )

    clean_docs = [{k: v for k, v in d.items() if k not in ("pages", "rows")} for d in docs]
    out_data = {
        "documents": clean_docs,
        "raw": raw,
        "ledger": ledger,
        "duplicates": duplicates,
        "transfers": transfers,
        "review": review,
        "issues": issues,
        "vdc_report": vdc_report,
    }

    if math_failures and not allow_unresolved:
        raise ValueError(
            "Reconciliation failed VDC gate: "
            + "; ".join(f"{f['check_id']} (delta={f['delta']})" for f in math_failures[:5])
        )

    return out_data


def main() -> None:
    parser = argparse.ArgumentParser(description="Reconcile extracted bank statements into books_data.json.")
    parser.add_argument("--parsed", required=True, type=Path, help="Path to parsed statements JSON.")
    parser.add_argument("--out", required=True, type=Path, help="Output path for books_data.json.")
    parser.add_argument("--entity-keyword", action="append", default=[], help="Business name keyword(s) for transfer matching.")
    parser.add_argument(
        "--allow-unresolved",
        action="store_true",
        help="Write books_data.json even when statement header or running-balance checks fail (marks verdict UNRESOLVED).",
    )
    parser.add_argument("--json", action="store_true", help="Print structured VDC reconciliation report as JSON.")
    args = parser.parse_args()

    docs = json.loads(args.parsed.read_text(encoding="utf-8"))
    try:
        out_data = reconcile_documents(docs, entity_keywords=args.entity_keyword, allow_unresolved=args.allow_unresolved)
    except ValueError as exc:
        err_payload = {
            "status": "FAILED",
            "error": str(exc),
            "attribution": "unresolved",
            "next_measurement": "Re-run with --allow-unresolved or inspect extracted rows against PDF header controls.",
        }
        if args.json:
            print(json.dumps(err_payload, indent=2), file=sys.stderr)
        else:
            print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(json.dumps(out_data, ensure_ascii=False, indent=2), encoding="utf-8")

    if args.json:
        print(
            json.dumps(
                {
                    "status": "PASSED",
                    "raw_rows": len(out_data["raw"]),
                    "ledger_rows": len(out_data["ledger"]),
                    "transfers": len(out_data["transfers"]),
                    "review_items": len(out_data["review"]),
                    "scope_issues": len(out_data["issues"]),
                    "vdc_report": out_data["vdc_report"],
                },
                indent=2,
            )
        )
    else:
        print(
            f"Reconciled {len(out_data['raw'])} bank rows -> {len(out_data['ledger'])} ledger components "
            f"({len(out_data['transfers'])} transfer pairs, {len(out_data['issues'])} scope issue(s))."
        )


if __name__ == "__main__":
    main()
