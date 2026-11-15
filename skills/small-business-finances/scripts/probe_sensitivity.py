#!/usr/bin/env python3
"""Layer 2 Development-Time Evaluation: Seeded-Defect Sensitivity Probing.

Implements Criterion 5 of Verification-Driven Coding (Kajiyama, 2026):
  - Seeds 10 distinct classes of financial and artifact defects into isolated temp copies.
  - Runs an unseeded control to measure false alarms.
  - Reports detections, misses, false alarms, sensitivity, and attribution (known_by_design).
"""

from __future__ import annotations

import argparse
import copy
import json
from pathlib import Path
import shutil
import tempfile
from typing import Any
from zipfile import ZipFile

from build_all import build_package
from reconcile_ledger import reconcile_documents
from verify_package import verify_all


def _corrupt_xlsx_formula(xlsx_path: Path, old_sub: str, new_sub: str) -> None:
    """Rewrite a worksheet XML inside an .xlsx zip while keeping valid OOXML syntax."""
    tmp_zip = xlsx_path.with_suffix(".tmp.xlsx")
    replaced = False
    with ZipFile(xlsx_path, "r") as zin, ZipFile(tmp_zip, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/worksheets/sheet2.xml":
                txt = data.decode("utf-8")
                if old_sub in txt:
                    txt = txt.replace(old_sub, new_sub, 1)
                    replaced = True
                data = txt.encode("utf-8")
            zout.writestr(item, data)
    if not replaced:
        raise RuntimeError(f"Could not locate '{old_sub}' in {xlsx_path} sheet2.xml")
    shutil.move(str(tmp_zip), str(xlsx_path))


def run_sensitivity_probe(
    parsed_path: Path,
    classifications_path: Path | None = None,
    entity_keywords: list[str] | None = None,
) -> dict[str, Any]:
    """Run 1 unseeded control + 10 seeded-defect probes in an isolated temporary directory."""
    base_parsed = json.loads(parsed_path.read_text(encoding="utf-8"))
    results: list[dict[str, Any]] = []

    with tempfile.TemporaryDirectory() as td:
        work = Path(td)
        base_books = reconcile_documents(copy.deepcopy(base_parsed), entity_keywords=entity_keywords)
        base_books_path = work / "base_books.json"
        base_books_path.write_text(json.dumps(base_books, indent=2), encoding="utf-8")

        out_dir = work / "base_out"
        pkg = build_package(
            books_path=base_books_path,
            out_dir=out_dir,
            slug="Sensitivity_Base",
            classifications_path=classifications_path,
            parsed_path=parsed_path,
        )

        # Unseeded control
        ctrl_ok = pkg["verification"]["status"] == "PASSED"
        downgrade_delta = next(
            (d["measured"] for d in base_books["vdc_report"]["deltas"] if d["check_id"] == "reconcile.downgraded_unverified_claims"),
            0,
        )
        false_alarm = not ctrl_ok or downgrade_delta != 0
        results.append(
            {
                "case_id": "C00_unseeded_control",
                "kind": "control",
                "expected_outcome": "pass",
                "observed_outcome": "pass" if not false_alarm else "false_alarm",
                "false_alarm": false_alarm,
                "attribution": "none",
            }
        )

        def run_verify_expect_fail(case_id: str, defect_class: str, books_obj: dict[str, Any], mutate_artifacts=None) -> None:
            case_dir = work / case_id
            case_dir.mkdir(parents=True, exist_ok=True)
            b_path = case_dir / "books.json"
            b_path.write_text(json.dumps(books_obj), encoding="utf-8")

            xlsx_p = case_dir / "books.xlsx"
            html_p = case_dir / "dash.html"
            md_p = case_dir / "health.md"
            csv_p = case_dir / "ledger.csv"
            shutil.copy2(pkg["outputs"]["workbook_xlsx"], xlsx_p)
            shutil.copy2(pkg["outputs"]["dashboard_html"], html_p)
            shutil.copy2(pkg["outputs"]["health_report_md"], md_p)
            shutil.copy2(pkg["outputs"]["ledger_csv"], csv_p)

            if mutate_artifacts:
                mutate_artifacts(xlsx_p, html_p, md_p, csv_p)

            detected = False
            field_caught = ""
            try:
                verify_all(b_path, xlsx_p, html_p, md_p, csv_p, classifications_path, parsed_path)
            except Exception as exc:
                detected = True
                field_caught = str(exc)

            results.append(
                {
                    "case_id": case_id,
                    "kind": "seeded_defect",
                    "defect_class": defect_class,
                    "detected": detected,
                    "reported_in": field_caught[:160] if detected else "MISSED",
                    "attribution": "known_by_design",
                }
            )

        # D01: Running balance break (+0.05)
        p1 = copy.deepcopy(base_parsed)
        first_valid = next(d for d in p1 if d.get("rows"))
        first_valid["rows"][0]["balance"] = round(first_valid["rows"][0]["balance"] + 0.05, 2)
        d01_det = False
        d01_msg = ""
        try:
            reconcile_documents(p1, entity_keywords=entity_keywords)
        except ValueError as exc:
            d01_det = True
            d01_msg = str(exc)
        results.append(
            {
                "case_id": "D01_balance_break",
                "kind": "seeded_defect",
                "defect_class": "running_balance_break",
                "detected": d01_det,
                "reported_in": d01_msg[:160] if d01_det else "MISSED",
                "attribution": "known_by_design",
            }
        )

        # D02: Header credit total mismatch (+1.00, simulates dropped PDF row)
        p2 = copy.deepcopy(base_parsed)
        first_valid2 = next(d for d in p2 if d.get("rows") and d.get("credits") is not None)
        first_valid2["credits"] = round(float(first_valid2["credits"]) + 1.00, 2)
        d02_det = False
        d02_msg = ""
        try:
            reconcile_documents(p2, entity_keywords=entity_keywords)
        except ValueError as exc:
            d02_det = True
            d02_msg = str(exc)
        results.append(
            {
                "case_id": "D02_header_credit_mismatch",
                "kind": "seeded_defect",
                "defect_class": "statement_header_control_mismatch",
                "detected": d02_det,
                "reported_in": d02_msg[:160] if d02_det else "MISSED",
                "attribution": "known_by_design",
            }
        )

        # D03: Undetected duplicate row injected into books_data raw & ledger
        b3 = copy.deepcopy(base_books)
        dup_raw = copy.deepcopy(b3["raw"][0])
        dup_raw["raw_id"] = "B9999"
        b3["raw"].append(dup_raw)
        dup_led = copy.deepcopy(b3["ledger"][0])
        dup_led["id"] = "B9999-P"
        dup_led["raw_id"] = "B9999"
        b3["ledger"].append(dup_led)
        run_verify_expect_fail("D03_undetected_duplicate", "duplicate_row_leak", b3)

        # D04: Asymmetric transfer pair
        b4 = copy.deepcopy(base_books)
        if b4["transfers"]:
            tid = b4["transfers"][0]["id"]
            leg = next(e for e in b4["ledger"] if e.get("transfer_id") == tid)
            leg["amount"] = round(leg["amount"] + 10.0, 2)
        run_verify_expect_fail("D04_asymmetric_transfer", "asymmetric_transfer_pair", b4)

        # D05: Bad embedded fee split
        b5 = copy.deepcopy(base_books)
        fee_entry = next(e for e in b5["ledger"] if e["component"] == "Embedded fee")
        fee_entry["amount"] = round(fee_entry["amount"] - 0.50, 2)
        run_verify_expect_fail("D05_bad_fee_split", "embedded_fee_split_mismatch", b5)

        # D06: Hallucinated Revenue classification without provenance or narration keywords
        p6 = copy.deepcopy(base_parsed)
        target_row = next(
            r
            for d in p6
            if d.get("rows")
            for r in d["rows"]
            if r.get("section") == "Needs Review" and float(r.get("credit") or 0) > 0
        )
        target_row["description"] = "ZELLE TRANSFER FROM UNKNOWN PARTY 8891"
        target_row["counterparty"] = "UNKNOWN PARTY"
        target_row["section"] = "Revenue"
        target_row["category"] = "Studio services"
        target_row.pop("provenance", None)
        reconciled_6 = reconcile_documents(p6, entity_keywords=entity_keywords)
        downgraded_6 = next(
            (d["measured"] for d in reconciled_6["vdc_report"]["deltas"] if d["check_id"] == "reconcile.downgraded_unverified_claims"),
            0,
        )
        matched_entry_6 = next(
            e for e in reconciled_6["ledger"] if e["description"] == "ZELLE TRANSFER FROM UNKNOWN PARTY 8891"
        )
        d06_det = downgraded_6 >= 1 and matched_entry_6["section"] == "Needs Review"
        results.append(
            {
                "case_id": "D06_hallucinated_revenue_no_provenance",
                "kind": "seeded_defect",
                "defect_class": "ungrounded_classification_claim",
                "detected": d06_det,
                "reported_in": f"downgraded={downgraded_6}, section={matched_entry_6['section']}" if d06_det else "MISSED",
                "attribution": "known_by_design",
            }
        )

        # D07: Ledger component dropped (cash-altering ledger defect)
        b7 = copy.deepcopy(base_books)
        b7["ledger"] = [e for e in b7["ledger"] if e["id"] != b7["ledger"][0]["id"]]
        run_verify_expect_fail("D07_cash_altering_ledger_drop", "ledger_cash_conservation_violation", b7)

        # D08: Corrupted Excel SUMIFS formula (balanced parens, wrong section string)
        b8 = copy.deepcopy(base_books)
        run_verify_expect_fail(
            "D08_corrupted_xlsx_sumifs",
            "excel_formula_semantic_corruption",
            b8,
            mutate_artifacts=lambda x, h, m, c: _corrupt_xlsx_formula(x, '"Revenue"', '"WrongSection"'),
        )

        # D09: Dashboard HTML payload drift (+100.00 net profit)
        b9 = copy.deepcopy(base_books)

        def mutate_html(x: Path, h: Path, m: Path, c: Path) -> None:
            txt = h.read_text(encoding="utf-8")
            marker = '<script id="books-payload" type="application/json">'
            i1 = txt.index(marker) + len(marker)
            i2 = txt.index("</script>", i1)
            payload = json.loads(txt[i1:i2].replace("<\\/", "</"))
            payload["summary_pnl"]["net_profit"]["total"] = round(payload["summary_pnl"]["net_profit"]["total"] + 100.0, 2)
            new_json = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")
            h.write_text(txt[:i1] + new_json + txt[i2:], encoding="utf-8")

        run_verify_expect_fail("D09_dashboard_payload_drift", "html_payload_drift", b9, mutate_artifacts=mutate_html)

        # D10: Forbidden advisory phrase injected into Health Report
        b10 = copy.deepcopy(base_books)
        run_verify_expect_fail(
            "D10_advisory_phrase_injected",
            "prescriptive_advisory_violation",
            b10,
            mutate_artifacts=lambda x, h, m, c: m.write_text(
                m.read_text(encoding="utf-8") + "\nYou should reduce rent expenses immediately.\n",
                encoding="utf-8",
            ),
        )

    seeded = [r for r in results if r["kind"] == "seeded_defect"]
    detected_cnt = sum(1 for r in seeded if r["detected"])
    missed_cnt = len(seeded) - detected_cnt
    false_alarms = sum(1 for r in results if r["kind"] == "control" and r["false_alarm"])

    return {
        "status": "PASSED" if missed_cnt == 0 and false_alarms == 0 else "FAILED",
        "total_seeded": len(seeded),
        "detected": detected_cnt,
        "missed": missed_cnt,
        "false_alarms": false_alarms,
        "sensitivity": round(detected_cnt / len(seeded), 4) if seeded else 0.0,
        "cases": results,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Run Layer 2 seeded-defect sensitivity probe against VDC verifiers.")
    parser.add_argument("--parsed", type=Path, default=None, help="Path to parsed_statements.json (defaults to examples/demo_work/parsed_statements.json).")
    parser.add_argument("--classifications", type=Path, default=None, help="Optional classifications.json path.")
    parser.add_argument("--entity-keyword", action="append", default=[], help="Business name keyword(s) for transfer matching.")
    parser.add_argument("--json", action="store_true", help="Print JSON report.")
    args = parser.parse_args()

    repo_root = Path(__file__).resolve().parent.parent.parent.parent
    parsed_p = args.parsed or (repo_root / "examples" / "demo_work" / "parsed_statements.json")
    cls_p = args.classifications or (
        repo_root / "examples" / "demo_work" / "classifications.json" if args.parsed is None else None
    )
    kw = args.entity_keyword or (["AURA STUDIO"] if args.parsed is None else [])

    report = run_sensitivity_probe(parsed_p, cls_p, entity_keywords=kw)
    print(json.dumps(report, indent=2))
    if report["status"] != "PASSED":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
