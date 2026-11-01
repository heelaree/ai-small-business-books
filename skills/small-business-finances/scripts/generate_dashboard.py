#!/usr/bin/env python3
"""Standalone Interactive HTML Dashboard generator.

Embeds reconciled ledger, Summary P&L, Financial Health metrics, and Smart Counterparty Groups
into a single zero-server HTML file that opens offline in any browser or hosts on GitHub Pages.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from ledger_engine import load_and_compute


def render_dashboard_html(computed: dict[str, Any], template_path: Path | None = None) -> str:
    if template_path is None:
        template_path = Path(__file__).resolve().parent.parent / "templates" / "dashboard.html"
    tpl = template_path.read_text(encoding="utf-8")

    # Trim payload to only what the browser needs so HTML stays fast and lightweight
    slim_payload = {
        "profile": computed["profile"],
        "currency": computed["currency"],
        "currency_symbol": computed["currency_symbol"],
        "months": computed["months"],
        "month_partial_flags": computed["month_partial_flags"],
        "accounts": computed["accounts"],
        "transfers": computed["transfers"],
        "smart_review_groups": computed["smart_review_groups"],
        "summary_pnl": computed["summary_pnl"],
        "cash_flow": computed["cash_flow"],
        "health": computed["health"],
        "classifications_config": computed["classifications_config"],
        "ledger": [
            {
                "id": e["id"],
                "raw_id": e["raw_id"],
                "date": e["date"],
                "month": e["month"],
                "account": e["account"],
                "description": e["description"],
                "component": e["component"],
                "amount": e["amount"],
                "section": e["section"],
                "category": e["category"],
                "summary_group": e["summary_group"],
                "counterparty_key": e["counterparty_key"],
                "counterparty_display": e["counterparty_display"],
                "question": e.get("question", ""),
                "statement": e.get("statement", ""),
                "pages": e.get("pages", ""),
                "reference": e.get("reference", ""),
            }
            for e in computed["ledger"]
        ],
    }

    # Escape </script> sequences safely inside JSON script tag
    json_str = json.dumps(slim_payload, ensure_ascii=False).replace("</", "<\\/")
    title = f"{computed['profile']['business_name']} — Interactive Financial Overview"
    return tpl.replace("{{PAGE_TITLE}}", title).replace("{{BOOKS_JSON_PAYLOAD}}", json_str)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate standalone interactive HTML financial dashboard.")
    parser.add_argument("--books", required=True, type=Path, help="Path to books_data.json.")
    parser.add_argument("--classifications", type=Path, default=None, help="Optional classifications.json path.")
    parser.add_argument("--out", required=True, type=Path, help="Output HTML file path.")
    args = parser.parse_args()

    computed = load_and_compute(args.books, args.classifications)
    html = render_dashboard_html(computed)
    args.out.parent.mkdir(parents=True, exist_ok=True)
    args.out.write_text(html, encoding="utf-8")
    print(f"Wrote Interactive Dashboard -> {args.out} ({len(html) / 1024:.1f} KB)")


if __name__ == "__main__":
    main()
