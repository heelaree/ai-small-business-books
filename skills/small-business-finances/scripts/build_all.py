#!/usr/bin/env python3
"""Single-command orchestrator to generate and verify the complete Small-Business Finance package.

Outputs:
1. Interactive HTML Financial Overview & Classification Dashboard (.html)
2. Objective Financial Health & Cost Report (.md)
3. Accountant & Owner Multi-Tab Cash Books (.xlsx)
4. Categorized Cash Ledger (.csv)
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from generate_dashboard import render_dashboard_html
from generate_health_report import render_health_markdown
from generate_workbook import build_workbook, write_csv_ledger
from ledger_engine import load_and_compute
from verify_package import verify_all


def safe_save_workbook(wb, target_path: Path) -> Path:
    """Save workbook, gracefully handling Windows Excel file locks if an older copy is open."""
    target_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(target_path)
        return target_path
    except PermissionError:
        alt = target_path.with_name(f"{target_path.stem}_updated{target_path.suffix}")
        wb.save(alt)
        print(f"Note: '{target_path.name}' is currently open in Excel; saved to '{alt.name}' instead.")
        return alt


def build_package(
    books_path: Path,
    out_dir: Path,
    slug: str,
    classifications_path: Path | None = None,
    parsed_path: Path | None = None,
    html_out: Path | None = None,
) -> dict:
    computed = load_and_compute(books_path, classifications_path, parsed_path)
    out_dir.mkdir(parents=True, exist_ok=True)

    html_path = html_out or (out_dir / f"{slug}_Financial_Dashboard.html")
    md_path = out_dir / f"{slug}_Financial_Health_Report.md"
    xlsx_path = out_dir / f"{slug}_Cash_Books.xlsx"
    csv_path = out_dir / f"{slug}_Cash_Ledger.csv"

    html_path.parent.mkdir(parents=True, exist_ok=True)
    html_path.write_text(render_dashboard_html(computed), encoding="utf-8")
    md_path.write_text(render_health_markdown(computed), encoding="utf-8")

    wb = build_workbook(computed)
    actual_xlsx = safe_save_workbook(wb, xlsx_path)
    write_csv_ledger(computed, csv_path)

    verification = verify_all(
        books_path=books_path,
        xlsx_path=actual_xlsx,
        html_path=html_path,
        md_path=md_path,
        csv_path=csv_path,
        classifications_path=classifications_path,
        parsed_path=parsed_path,
    )
    return {
        "outputs": {
            "dashboard_html": str(html_path),
            "health_report_md": str(md_path),
            "workbook_xlsx": str(actual_xlsx),
            "ledger_csv": str(csv_path),
        },
        "verification": verification,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Build and verify full small-business finance package.")
    parser.add_argument("--books", required=True, type=Path, help="Path to books_data.json.")
    parser.add_argument("--out-dir", required=True, type=Path, help="Output folder.")
    parser.add_argument("--slug", required=True, help="Filename prefix (e.g., Aura_Studio_2026).")
    parser.add_argument("--classifications", type=Path, default=None, help="Optional classifications.json.")
    parser.add_argument("--parsed", type=Path, default=None, help="Optional parsed.json for full source rows.")
    parser.add_argument("--html-out", type=Path, default=None, help="Optional custom HTML path (e.g., docs/index.html).")
    args = parser.parse_args()

    result = build_package(
        books_path=args.books,
        out_dir=args.out_dir,
        slug=args.slug,
        classifications_path=args.classifications,
        parsed_path=args.parsed,
        html_out=args.html_out,
    )
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
