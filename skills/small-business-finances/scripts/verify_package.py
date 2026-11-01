#!/usr/bin/env python3
"""Comprehensive verification suite for the Small-Business Financial Insights package.

Checks:
1. Exact arithmetic & cash-bridge tie-out across ledger, Summary P&L, Detailed P&L, and Cash Flow.
2. Excel (.xlsx) 14-sheet priority structure, balanced formula syntax, and cross-sheet references.
3. Standalone HTML dashboard JSON payload integrity and round-trip override invariance.
4. Strict Non-Advisory Guardrail scan on the Financial Health Report (zero suggestions/advice).
"""

from __future__ import annotations

import argparse
import copy
import csv
import json
from pathlib import Path
import re
import tempfile
from typing import Any
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from ledger_engine import OPEX_BUCKETS, load_and_compute

EXPECTED_SHEETS = [
    "Executive Overview",
    "Summary P&L",
    "Financial Health",
    "Detailed P&L",
    "Cash Flow",
    "Master Ledger",
    "Needs Review",
    "Account Cash",
    "Transfer Matches",
    "Statement Inventory",
    "Source Rows",
    "Duplicate Exclusions",
    "Reconciliations",
    "ReadMe",
]

FORBIDDEN_ADVISORY_PATTERNS = [
    r"\byou should\b",
    r"\bwe recommend\b",
    r"\brecommended action\b",
    r"\bconsider cutting\b",
    r"\bconsider reducing\b",
    r"\btry to reduce\b",
    r"\bsuggestions?:\b",
    r"\bour advice\b",
]


def verify_math_and_bridge(computed: dict[str, Any]) -> dict[str, Any]:
    sp = computed["summary_pnl"]
    h = computed["health"]
    ledger = computed["ledger"]
    raw = computed["raw_transactions"]

    raw_net = round(sum(float(r["credit"]) - float(r["debit"]) for r in raw), 2)
    ledger_net = round(sum(float(e["amount"]) for e in ledger), 2)
    assert abs(raw_net - ledger_net) < 0.01, f"Ledger net {ledger_net} != raw net {raw_net}"

    # Summary P&L equations
    gp_calc = round(sp["revenue"]["total"] - sp["direct_costs"]["total"], 2)
    assert abs(gp_calc - sp["gross_profit"]["total"]) < 0.01, "Gross profit mismatch"

    opex_sum = round(sum(sp["opex_buckets"][b]["total"] for b in OPEX_BUCKETS), 2)
    assert abs(opex_sum - sp["total_opex"]["total"]) < 0.01, "OpEx bucket sum mismatch"

    op_calc = round(sp["gross_profit"]["total"] - sp["total_opex"]["total"], 2)
    assert abs(op_calc - sp["operating_profit"]["total"]) < 0.01, "Operating profit mismatch"

    net_calc = round(sp["operating_profit"]["total"] - sp["taxes_and_fees"]["total"], 2)
    assert abs(net_calc - sp["net_profit"]["total"]) < 0.01, "Net profit mismatch"

    # Cash Bridge equation
    bridge_cash = round(
        sp["net_profit"]["total"]
        + sp["unclassified_in"]["total"]
        - sp["unclassified_out"]["total"]
        + sp["non_pnl_net"]["total"],
        2,
    )
    assert abs(bridge_cash - sp["actual_cash_change"]["total"]) < 0.01, (
        f"Cash bridge {bridge_cash} != actual cash change {sp['actual_cash_change']['total']}"
    )
    assert abs(bridge_cash - ledger_net) < 0.01, f"Cash bridge {bridge_cash} != ledger net {ledger_net}"

    expected_close = round(h["opening_cash"] + bridge_cash, 2)
    assert abs(expected_close - h["closing_cash"]) < 0.01, (
        f"Opening + net change ({expected_close}) != closing cash ({h['closing_cash']})"
    )

    # Transfer pair nets
    for t in computed["transfers"]:
        pair_net = round(sum(e["amount"] for e in ledger if e.get("transfer_id") == t["id"]), 2)
        assert abs(pair_net) < 0.01, f"Transfer {t['id']} net {pair_net} != 0"

    return {
        "ledger_rows": len(ledger),
        "raw_transactions": len(raw),
        "revenue": sp["revenue"]["total"],
        "direct_costs": sp["direct_costs"]["total"],
        "gross_profit": sp["gross_profit"]["total"],
        "operating_expenses": sp["total_opex"]["total"],
        "operating_profit": sp["operating_profit"]["total"],
        "taxes_and_fees": sp["taxes_and_fees"]["total"],
        "net_profit": sp["net_profit"]["total"],
        "unclassified_in": sp["unclassified_in"]["total"],
        "unclassified_out": sp["unclassified_out"]["total"],
        "actual_cash_change": sp["actual_cash_change"]["total"],
        "opening_cash": h["opening_cash"],
        "closing_cash": h["closing_cash"],
    }


def verify_xlsx(xlsx_path: Path) -> dict[str, Any]:
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with ZipFile(xlsx_path) as z:
        assert z.testzip() is None, "Corrupt XLSX zip archive"
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = [s.attrib["name"] for s in wb_xml.find("s:sheets", ns)]
        assert sheets == EXPECTED_SHEETS, f"Sheet order mismatch: {sheets} != {EXPECTED_SHEETS}"

        sheet_set = set(sheets)
        formula_count = 0
        for idx, s_name in enumerate(sheets, 1):
            tree = ET.fromstring(z.read(f"xl/worksheets/sheet{idx}.xml"))
            for cell in tree.findall(".//s:sheetData/s:row/s:c", ns):
                coord = cell.attrib.get("r", "")
                assert cell.attrib.get("t") != "e", f"Excel error cell at {s_name}!{coord}"
                f_node = cell.find("s:f", ns)
                if f_node is not None and f_node.text:
                    formula_count += 1
                    ftxt = f_node.text
                    # Strip quoted strings before checking parentheses & sheet refs
                    stripped = re.sub(r'"[^"]*"', '""', ftxt)
                    assert stripped.count("(") == stripped.count(")"), (
                        f"Unbalanced parentheses in {s_name}!{coord}: {ftxt}"
                    )
                    for ref_sheet in re.findall(r"'([^']+)'!", stripped):
                        assert ref_sheet in sheet_set, (
                            f"Unknown sheet reference '{ref_sheet}' in {s_name}!{coord}"
                        )

    return {"sheets": len(sheets), "formula_cells": formula_count}


def verify_dashboard_html(html_path: Path, computed: dict[str, Any]) -> dict[str, Any]:
    html = html_path.read_text(encoding="utf-8")
    m = re.search(r'<script id="books-payload" type="application/json">(.*?)</script>', html, re.DOTALL)
    assert m is not None, "Missing embedded books-payload JSON in dashboard HTML"
    payload = json.loads(m.group(1).replace("<\\/", "</"))
    assert len(payload["ledger"]) == len(computed["ledger"]), "Dashboard ledger count mismatch"
    assert abs(payload["summary_pnl"]["net_profit"]["total"] - computed["summary_pnl"]["net_profit"]["total"]) < 0.01
    assert abs(payload["health"]["closing_cash"] - computed["health"]["closing_cash"]) < 0.01
    return {"html_kb": round(len(html) / 1024, 1), "smart_groups": len(payload["smart_review_groups"])}


def verify_health_report_guardrails(md_path: Path) -> dict[str, Any]:
    text = md_path.read_text(encoding="utf-8")
    for pat in FORBIDDEN_ADVISORY_PATTERNS:
        m = re.search(pat, text, re.IGNORECASE)
        assert m is None, f"Forbidden prescriptive language '{m.group(0)}' found in Health Report!"
    assert "## 1. Summary Profit & Loss" in text
    assert "## 3. Where Costs Are" in text
    return {"markdown_chars": len(text), "advisory_violations": 0}


def verify_override_roundtrip(books_path: Path, classifications_path: Path | None) -> bool:
    base = load_and_compute(books_path, classifications_path)
    unknown_in = next((e for e in base["ledger"] if e["section"] == "Needs Review" and e["amount"] > 0), None)
    if not unknown_in:
        return True

    base_cls = copy.deepcopy(base.get("classifications_config") or {})
    base_cls.setdefault("transaction_overrides", {})[unknown_in["id"]] = {
        "section": "Revenue",
        "category": "Owner verified service sale",
        "note": "Automated roundtrip verification",
    }
    with tempfile.NamedTemporaryFile("w", suffix=".json", encoding="utf-8", delete=False) as tf:
        json.dump(base_cls, tf)
        tmp_path = Path(tf.name)

    try:
        updated = load_and_compute(books_path, tmp_path)
        delta_profit = round(
            updated["summary_pnl"]["net_profit"]["total"] - base["summary_pnl"]["net_profit"]["total"], 2
        )
        assert abs(delta_profit - unknown_in["amount"]) < 0.01, (
            f"Override profit delta {delta_profit} != {unknown_in['amount']}"
        )
        assert abs(
            updated["summary_pnl"]["actual_cash_change"]["total"]
            - base["summary_pnl"]["actual_cash_change"]["total"]
        ) < 0.01, "Classification override must never alter actual bank cash change!"
    finally:
        tmp_path.unlink(missing_ok=True)
    return True


def verify_all(
    books_path: Path,
    xlsx_path: Path,
    html_path: Path,
    md_path: Path,
    csv_path: Path | None = None,
    classifications_path: Path | None = None,
    parsed_path: Path | None = None,
) -> dict[str, Any]:
    computed = load_and_compute(books_path, classifications_path, parsed_path)
    math_res = verify_math_and_bridge(computed)
    xlsx_res = verify_xlsx(xlsx_path)
    html_res = verify_dashboard_html(html_path, computed)
    md_res = verify_health_report_guardrails(md_path)
    roundtrip_ok = verify_override_roundtrip(books_path, classifications_path)

    if csv_path and csv_path.exists():
        with csv_path.open(encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))
        assert len(rows) == len(computed["ledger"]), "CSV row count mismatch"
        csv_sum = round(sum(float(r["amount"]) for r in rows), 2)
        assert abs(csv_sum - math_res["actual_cash_change"]) < 0.01, "CSV net sum mismatch"

    return {
        "status": "PASSED",
        "math": math_res,
        "xlsx": xlsx_res,
        "dashboard": html_res,
        "health_report": md_res,
        "override_roundtrip": roundtrip_ok,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Verify generated Small-Business Financial Insights package.")
    parser.add_argument("--books", required=True, type=Path)
    parser.add_argument("--xlsx", required=True, type=Path)
    parser.add_argument("--html", required=True, type=Path)
    parser.add_argument("--md", required=True, type=Path)
    parser.add_argument("--csv", type=Path, default=None)
    parser.add_argument("--classifications", type=Path, default=None)
    parser.add_argument("--parsed", type=Path, default=None)
    args = parser.parse_args()

    summary = verify_all(
        args.books, args.xlsx, args.html, args.md, args.csv, args.classifications, args.parsed
    )
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
