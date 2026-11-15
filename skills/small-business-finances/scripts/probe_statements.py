#!/usr/bin/env python3
"""Pre-measurement probe for bank statement files and parsed_statements.json.

Implements Criterion 2 & 3a of Verification-Driven Coding (Kajiyama, 2026):
  - Measures statement inventory, SHA-256 hashes, page counts, and encryption status before acting.
  - Explicitly classifies password-protected or unreadable files as `unknown_encrypted` (never `missing`).
  - Detects coverage date gaps and mixed account cutoff dates across parsed statements.
"""

from __future__ import annotations

import argparse
import calendar
from datetime import date, datetime, timedelta
import hashlib
import json
from pathlib import Path
from typing import Any

from _contract import BALANCE_TOLERANCE, UNVERIFIED_PROPERTIES, make_delta, make_vdc_report


def _parse_iso(d: str | None) -> date | None:
    if not d:
        return None
    return datetime.strptime(d[:10], "%Y-%m-%d").date()


def detect_coverage_gaps(docs: list[dict[str, Any]]) -> dict[str, Any]:
    """Detect encrypted files, intra-period missing days, and mixed final-month cutoffs."""
    encrypted_docs = [d for d in docs if d.get("encrypted") or d.get("error")]
    readable_docs = [d for d in docs if "rows" in d and not d.get("error")]
    monthly_docs = [d for d in readable_docs if not d.get("is_consolidated", False) and d.get("id") != "S17"]

    by_account: dict[str, list[dict[str, Any]]] = {}
    for d in monthly_docs:
        acct = d.get("account") or "Unknown Account"
        if d.get("start") and d.get("end"):
            by_account.setdefault(acct, []).append(d)

    missing_day_gaps: list[dict[str, Any]] = []
    latest_ends: dict[str, str] = {}

    for acct, acct_docs in sorted(by_account.items()):
        acct_docs.sort(key=lambda x: (x["start"], x["end"]))
        latest_ends[acct] = acct_docs[-1]["end"]
        for idx, cur in enumerate(acct_docs):
            end_dt = _parse_iso(cur["end"])
            if not end_dt:
                continue
            # Check if statement ends before calendar month end while a later statement exists for this account
            if idx + 1 < len(acct_docs):
                next_start = _parse_iso(acct_docs[idx + 1]["start"])
                month_last_day = calendar.monthrange(end_dt.year, end_dt.month)[1]
                expected_month_end = date(end_dt.year, end_dt.month, month_last_day)
                if next_start and (next_start - end_dt).days > 1:
                    gap_start = end_dt + timedelta(days=1)
                    gap_end = next_start - timedelta(days=1)
                    missing_day_gaps.append(
                        {
                            "account": acct,
                            "statement_id": cur["id"],
                            "gap_start": gap_start.isoformat(),
                            "gap_end": gap_end.isoformat(),
                            "missing_days": (gap_end - gap_start).days + 1,
                        }
                    )
                elif end_dt < expected_month_end and next_start and next_start.month != end_dt.month:
                    missing_day_gaps.append(
                        {
                            "account": acct,
                            "statement_id": cur["id"],
                            "gap_start": (end_dt + timedelta(days=1)).isoformat(),
                            "gap_end": expected_month_end.isoformat(),
                            "missing_days": (expected_month_end - end_dt).days,
                        }
                    )

    unique_cutoffs = sorted(set(latest_ends.values()))
    mixed_cutoff = len(unique_cutoffs) > 1

    return {
        "encrypted_files": [
            {"id": d.get("id", ""), "file": d.get("file", ""), "state": "unknown_encrypted", "error": d.get("error", "")}
            for d in encrypted_docs
        ],
        "missing_day_gaps": missing_day_gaps,
        "latest_account_cutoffs": latest_ends,
        "mixed_cutoff": mixed_cutoff,
    }


def probe_directory(stmt_dir: Path) -> dict[str, Any]:
    """Pre-measure raw statement files in a directory before extraction."""
    try:
        from pypdf import PdfReader
    except ImportError:
        PdfReader = None

    files = sorted(
        p for p in stmt_dir.rglob("*") if p.is_file() and p.suffix.lower() in {".pdf", ".csv", ".xlsx"}
    )
    entries: list[dict[str, Any]] = []
    encrypted_count = 0
    for idx, p in enumerate(files, 1):
        rel = p.as_posix()
        sha = hashlib.sha256(p.read_bytes()).hexdigest()
        item: dict[str, Any] = {
            "probe_id": f"P{idx:02d}",
            "file": rel,
            "suffix": p.suffix.lower(),
            "bytes": p.stat().st_size,
            "sha256": sha,
            "state": "readable",
            "page_count": None,
        }
        if p.suffix.lower() == ".pdf":
            if PdfReader is None:
                item["state"] = "unknown_no_pypdf"
            else:
                try:
                    reader = PdfReader(p)
                    if reader.is_encrypted and not reader.decrypt(""):
                        item["state"] = "unknown_encrypted"
                        item["error"] = "Password required; contents unverified"
                        encrypted_count += 1
                    else:
                        item["page_count"] = len(reader.pages)
                except Exception as exc:
                    item["state"] = "unknown_unreadable"
                    item["error"] = str(exc)
                    encrypted_count += 1
        entries.append(item)

    unverified = list(UNVERIFIED_PROPERTIES["probe_statements"])
    if encrypted_count:
        unverified.append(f"encrypted_or_unreadable_files_count:{encrypted_count}")

    deltas = [
        make_delta(
            "probe.readable_files",
            "readable_file_count",
            len(entries),
            sum(1 for e in entries if e["state"] == "readable"),
            provenance="input_premeasurement",
            tolerance=0.0,
            notes="Encrypted or unreadable files are explicitly marked unknown_encrypted rather than missing.",
        )
    ]
    completeness = "VERIFIED" if encrypted_count == 0 and entries else "PROVISIONAL_REVIEW_REQUIRED"
    return {
        "mode": "directory",
        "target": str(stmt_dir),
        "total_files": len(entries),
        "readable_count": sum(1 for e in entries if e["state"] == "readable"),
        "unknown_count": sum(1 for e in entries if e["state"].startswith("unknown")),
        "files": entries,
        "vdc_report": make_vdc_report(
            deltas=deltas,
            unverified_range=unverified,
            math_integrity="VERIFIED" if entries else "UNRESOLVED",
            accounting_completeness=completeness,
            attribution_state="none" if entries else "unresolved",
            attribution_grounds="Pre-measurement file inventory completed." if entries else "No statement files found.",
        ),
    }


def probe_parsed(parsed_path: Path) -> dict[str, Any]:
    """Pre-measure parsed_statements.json before reconciliation."""
    docs = json.loads(parsed_path.read_text(encoding="utf-8"))
    coverage = detect_coverage_gaps(docs)
    readable_docs = [d for d in docs if "rows" in d and not d.get("error")]

    deltas: list[dict[str, Any]] = []
    missing_header_controls = 0
    for d in readable_docs:
        rows = d.get("rows", [])
        row_credits = round(sum(float(r.get("credit") or 0.0) for r in rows), 2)
        row_debits = round(sum(float(r.get("debit") or 0.0) for r in rows), 2)
        hdr_cr = d.get("credits")
        hdr_dr = d.get("debits")
        if hdr_cr is None or hdr_dr is None:
            missing_header_controls += 1
        deltas.append(
            make_delta(
                f"probe.{d['id']}.credits",
                f"{d['id']} extracted credits vs statement header",
                hdr_cr,
                row_credits,
                provenance="statement_header_control",
                tolerance=BALANCE_TOLERANCE,
            )
        )
        deltas.append(
            make_delta(
                f"probe.{d['id']}.debits",
                f"{d['id']} extracted debits vs statement header",
                hdr_dr,
                row_debits,
                provenance="statement_header_control",
                tolerance=BALANCE_TOLERANCE,
            )
        )

    failures = [x for x in deltas if x["status"] == "fail"]
    unverified = list(UNVERIFIED_PROPERTIES["probe_statements"])
    if coverage["encrypted_files"]:
        unverified.append(f"unknown_encrypted_files:{len(coverage['encrypted_files'])}")
    if coverage["missing_day_gaps"]:
        unverified.append(f"missing_calendar_day_gaps:{len(coverage['missing_day_gaps'])}")
    if coverage["mixed_cutoff"]:
        unverified.append("mixed_final_month_account_cutoffs")
    if missing_header_controls:
        unverified.append(f"statements_without_header_totals:{missing_header_controls}")

    has_unknowns = bool(
        coverage["encrypted_files"] or coverage["missing_day_gaps"] or coverage["mixed_cutoff"] or missing_header_controls
    )
    return {
        "mode": "parsed",
        "target": str(parsed_path),
        "total_documents": len(docs),
        "readable_documents": len(readable_docs),
        "total_extracted_rows": sum(len(d.get("rows", [])) for d in readable_docs),
        "coverage": coverage,
        "vdc_report": make_vdc_report(
            deltas=deltas,
            unverified_range=unverified,
            math_integrity="FAILED" if failures else "VERIFIED",
            accounting_completeness="PROVISIONAL_REVIEW_REQUIRED" if has_unknowns else "VERIFIED",
            attribution_state="unresolved" if failures else "none",
            attribution_grounds=(
                f"{len(failures)} statement header tie-out mismatch(es) detected before reconciliation."
                if failures
                else "All readable statement rows match reported statement header controls."
            ),
            next_measurement="Inspect PDF extraction rows vs statement summary headers for failing statements." if failures else "",
        ),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Pre-measure bank statements or parsed_statements.json.")
    parser.add_argument("--dir", type=Path, default=None, help="Directory of raw bank statement files.")
    parser.add_argument("--parsed", type=Path, default=None, help="Path to parsed_statements.json.")
    parser.add_argument("--out", type=Path, default=None, help="Optional JSON output path.")
    parser.add_argument("--json", action="store_true", help="Print JSON probe report to stdout.")
    args = parser.parse_args()

    if not args.dir and not args.parsed:
        parser.error("Provide --dir or --parsed.")

    result = probe_parsed(args.parsed) if args.parsed else probe_directory(args.dir)
    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(payload + "\n", encoding="utf-8")
    print(payload)


if __name__ == "__main__":
    main()
