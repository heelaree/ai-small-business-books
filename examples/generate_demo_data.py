#!/usr/bin/env python3
"""Generate a realistic, 100% synthetic small-business dataset ("Aura Studio & Spa")
and run the full `small-business-finances` pipeline to produce:
  - docs/index.html (GitHub Pages live interactive portfolio demo)
  - examples/output/Aura_Studio_2026_Financial_Health_Report.md
  - examples/output/Aura_Studio_2026_Cash_Books.xlsx
  - examples/output/Aura_Studio_2026_Cash_Ledger.csv
"""

from __future__ import annotations

import json
from pathlib import Path
import subprocess
import sys

ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = ROOT / "skills" / "small-business-finances" / "scripts"
DEMO_WORK = ROOT / "examples" / "demo_work"
DEMO_OUT = ROOT / "examples" / "output"
DOCS_DIR = ROOT / "docs"


def build_synthetic_parsed_statements() -> list[dict]:
    """Create 6 months (Jan-Jun 2026) of realistic two-account studio bank statements."""
    acct_main = "Metro Business Checking 884102"
    acct_pos = "Horizon Merchant Wallet 440918"

    # Monthly transaction scripts: (day, acct, desc, debit, credit, charge, cp, src_acct, sec, cat, rev_type, question)
    plan = [
        # JANUARY 2026
        ("2026-01-03", acct_pos, "POS CLIENT SETTLEMENT - FACIAL & BROW SERVICES", 0.0, 2450.0, 35.0, "CLIENT CARD BATCH", "", "Revenue", "Studio services", "", ""),
        ("2026-01-05", acct_main, "ACH DEBIT - LUMINA SKIN & LASH WHOLESALE SUPPLIES", 680.0, 0.0, 0.0, "LUMINA SUPPLIES", "", "Direct costs", "Treatment supplies & consumables", "", ""),
        ("2026-01-08", acct_main, "LEASE PAYMENT - SUITE 204 STUDIO RENT JAN", 1450.0, 0.0, 0.0, "OAKWOOD PROPERTIES", "", "Operating expenses", "Studio rent & lease", "", ""),
        ("2026-01-12", acct_pos, "POS CLIENT SETTLEMENT - AESTHETIC TREATMENTS", 0.0, 1980.0, 28.0, "CLIENT CARD BATCH", "", "Revenue", "Studio services", "", ""),
        ("2026-01-15", acct_pos, "SAME-DAY PAYOUT TRANSFER TO AURA STUDIO 884102", 2505.0, 0.0, 5.0, "AURA STUDIO LLC", "884102", "Internal transfers", "Matched transfer", "", ""),
        ("2026-01-15", acct_main, "INCOMING SETTLEMENT FROM HORIZON WALLET 440918 AURA STUDIO", 0.0, 2500.0, 0.0, "AURA STUDIO LLC", "440918", "Internal transfers", "Matched transfer", "", ""),
        ("2026-01-19", acct_main, "CONTRACTOR PAYOUT - NAIL & BROW SPECIALIST", 520.0, 0.0, 0.0, "MAYA LIN", "", "Direct costs", "Service contractors & technicians", "", ""),
        ("2026-01-22", acct_main, "CITY POWER & ELECTRIC UTILITY BILL", 215.0, 0.0, 0.0, "CITY POWER", "", "Operating expenses", "Electricity & utilities", "", ""),
        ("2026-01-25", acct_main, "META PLATFORMS ADS - INSTAGRAM LOCAL CAMPAIGN", 310.0, 0.0, 0.0, "META ADS", "", "Operating expenses", "Advertising & social media", "", ""),
        ("2026-01-28", acct_main, "PAYROLL ACH - FRONT DESK & STUDIO ASSISTANT", 1350.0, 0.0, 0.0, "STAFF PAYROLL", "", "Operating expenses", "Staff salaries & wages", "", ""),
        ("2026-01-30", acct_main, "INBOUND ZELLE TRANSFER - FROM VELVET LOUNGE CO", 0.0, 650.0, 0.0, "VELVET LOUNGE CO", "9920144", "Needs Review", "Unidentified receipts", "Receipt nature", "What was this receipt from Velvet Lounge Co (shared studio suite fee, event collab, or client package)?"),
        ("2026-01-31", acct_main, "MONTHLY ACCOUNT MAINTENANCE FEE", 15.0, 0.0, 0.0, "METRO BANK", "", "Bank fees", "Bank maintenance charges", "", ""),

        # FEBRUARY 2026
        ("2026-02-04", acct_pos, "POS CLIENT SETTLEMENT - VALENTINE SKINCARE PACKAGES", 0.0, 3120.0, 44.0, "CLIENT CARD BATCH", "", "Revenue", "Studio services", "", ""),
        ("2026-02-07", acct_main, "STUDIO RENT SUITE 204 - FEB", 1450.0, 0.0, 0.0, "OAKWOOD PROPERTIES", "", "Operating expenses", "Studio rent & lease", "", ""),
        ("2026-02-10", acct_main, "LUMINA SKIN & LASH WHOLESALE ORDER", 840.0, 0.0, 0.0, "LUMINA SUPPLIES", "", "Direct costs", "Treatment supplies & consumables", "", ""),
        ("2026-02-14", acct_pos, "POS CLIENT SETTLEMENT - LASH & BROW STUDIO", 0.0, 2640.0, 37.0, "CLIENT CARD BATCH", "", "Revenue", "Studio services", "", ""),
        ("2026-02-16", acct_pos, "TRANSFER TO METRO CHECKING 884102 AURA STUDIO", 4005.0, 0.0, 5.0, "AURA STUDIO LLC", "884102", "Internal transfers", "Matched transfer", "", ""),
        ("2026-02-16", acct_main, "TRANSFER FROM HORIZON WALLET 440918 AURA STUDIO", 0.0, 4000.0, 0.0, "AURA STUDIO LLC", "440918", "Internal transfers", "Matched transfer", "", ""),
        ("2026-02-18", acct_main, "CUSTOM CLIENT GIFT BAGS & PACKAGING BOXES", 195.0, 0.0, 0.0, "PACKCRAFT CO", "", "Direct costs", "Client packaging", "", ""),
        ("2026-02-20", acct_main, "SPARKLE COMMERCIAL CLEANING & SANITATION", 180.0, 0.0, 0.0, "SPARKLE CLEAN", "", "Operating expenses", "Cleaning & hygiene", "", ""),
        ("2026-02-23", acct_main, "CITY POWER & WATER BILL", 230.0, 0.0, 0.0, "CITY POWER", "", "Operating expenses", "Electricity & utilities", "", ""),
        ("2026-02-26", acct_main, "PAYROLL ACH - STUDIO STAFF", 1350.0, 0.0, 0.0, "STAFF PAYROLL", "", "Operating expenses", "Staff salaries & wages", "", ""),
        ("2026-02-27", acct_main, "CLIENT REFUND - RESCHEDULED APPOINTMENT", 95.0, 0.0, 0.0, "CLIENT REFUND", "", "Revenue", "Customer refunds", "", ""),
        ("2026-02-28", acct_main, "MONTHLY ACCOUNT MAINTENANCE FEE", 15.0, 0.0, 0.0, "METRO BANK", "", "Bank fees", "Bank maintenance charges", "", ""),

        # MARCH 2026
        ("2026-03-04", acct_pos, "POS CLIENT SETTLEMENT - FACIAL & PEEL SERVICES", 0.0, 2890.0, 41.0, "CLIENT CARD BATCH", "", "Revenue", "Studio services", "", ""),
        ("2026-03-06", acct_main, "STUDIO RENT SUITE 204 - MAR", 1450.0, 0.0, 0.0, "OAKWOOD PROPERTIES", "", "Operating expenses", "Studio rent & lease", "", ""),
        ("2026-03-11", acct_main, "CONTRACTOR PAYOUT - NAIL & BROW SPECIALIST", 640.0, 0.0, 0.0, "MAYA LIN", "", "Direct costs", "Service contractors & technicians", "", ""),
        ("2026-03-15", acct_pos, "POS CLIENT SETTLEMENT - SPRING PACKAGES", 0.0, 2750.0, 39.0, "CLIENT CARD BATCH", "", "Revenue", "Studio services", "", ""),
        ("2026-03-17", acct_pos, "TRANSFER TO METRO CHECKING 884102 AURA STUDIO", 5005.0, 0.0, 5.0, "AURA STUDIO LLC", "884102", "Internal transfers", "Matched transfer", "", ""),
        ("2026-03-17", acct_main, "TRANSFER FROM HORIZON WALLET 440918 AURA STUDIO", 0.0, 5000.0, 0.0, "AURA STUDIO LLC", "440918", "Internal transfers", "Matched transfer", "", ""),
        ("2026-03-20", acct_main, "META PLATFORMS ADS - SPRING PROMO", 420.0, 0.0, 0.0, "META ADS", "", "Operating expenses", "Advertising & social media", "", ""),
        ("2026-03-22", acct_main, "STEAMER & TREATMENT CHAIR SERVICING REPAIR", 340.0, 0.0, 0.0, "MEDSPA TECH REPAIR", "", "Operating expenses", "Equipment repairs & maintenance", "", ""),
        ("2026-03-25", acct_main, "INBOUND ZELLE TRANSFER - FROM VELVET LOUNGE CO", 0.0, 650.0, 0.0, "VELVET LOUNGE CO", "9920144", "Needs Review", "Unidentified receipts", "Receipt nature", "Recurring receipt from Velvet Lounge Co — confirm if sublease reimbursement or service revenue."),
        ("2026-03-27", acct_main, "CITY MUNICIPAL BUSINESS LICENSE LEVY", 240.0, 0.0, 0.0, "CITY REVENUE DEPT", "", "Taxes", "Municipal business license & levies", "", ""),
        ("2026-03-29", acct_main, "PAYROLL ACH - STUDIO STAFF", 1400.0, 0.0, 0.0, "STAFF PAYROLL", "", "Operating expenses", "Staff salaries & wages", "", ""),
        ("2026-03-31", acct_main, "MONTHLY ACCOUNT MAINTENANCE FEE", 15.0, 0.0, 0.0, "METRO BANK", "", "Bank fees", "Bank maintenance charges", "", ""),

        # APRIL 2026
        ("2026-04-03", acct_pos, "POS CLIENT SETTLEMENT - STUDIO TREATMENTS", 0.0, 3340.0, 47.0, "CLIENT CARD BATCH", "", "Revenue", "Studio services", "", ""),
        ("2026-04-06", acct_main, "STUDIO RENT SUITE 204 - APR", 1450.0, 0.0, 0.0, "OAKWOOD PROPERTIES", "", "Operating expenses", "Studio rent & lease", "", ""),
        ("2026-04-09", acct_main, "LUMINA SKIN & LASH WHOLESALE ORDER", 910.0, 0.0, 0.0, "LUMINA SUPPLIES", "", "Direct costs", "Treatment supplies & consumables", "", ""),
        ("2026-04-14", acct_main, "DERMALUX LED THERAPY MACHINE PURCHASE", 1850.0, 0.0, 0.0, "DERMALUX EQUIP", "", "Asset purchases", "Clinical studio equipment", "", ""),
        ("2026-04-16", acct_pos, "POS CLIENT SETTLEMENT - BRIDAL & EVENT BOOKINGS", 0.0, 2910.0, 42.0, "CLIENT CARD BATCH", "", "Revenue", "Studio services", "", ""),
        ("2026-04-18", acct_pos, "TRANSFER TO METRO CHECKING 884102 AURA STUDIO", 5505.0, 0.0, 5.0, "AURA STUDIO LLC", "884102", "Internal transfers", "Matched transfer", "", ""),
        ("2026-04-18", acct_main, "TRANSFER FROM HORIZON WALLET 440918 AURA STUDIO", 0.0, 5500.0, 0.0, "AURA STUDIO LLC", "440918", "Internal transfers", "Matched transfer", "", ""),
        ("2026-04-21", acct_main, "CITY POWER & ELECTRIC UTILITY BILL", 225.0, 0.0, 0.0, "CITY POWER", "", "Operating expenses", "Electricity & utilities", "", ""),
        ("2026-04-23", acct_main, "FIBER INTERNET & STUDIO BOOKING PHONE", 115.0, 0.0, 0.0, "FIBERNET TELECOM", "", "Operating expenses", "Internet, phone & booking software", "", ""),
        ("2026-04-25", acct_main, "OUTBOUND WIRE - JORDAN MERCER MEMO: APRIL TRANSFER", 480.0, 0.0, 0.0, "JORDAN MERCER", "7738120", "Needs Review", "Unidentified payments", "Payment nature", "Confirm purpose of $480 transfer to Jordan Mercer (contractor, owner draw, or vendor?)."),
        ("2026-04-28", acct_main, "PAYROLL ACH - STUDIO STAFF", 1400.0, 0.0, 0.0, "STAFF PAYROLL", "", "Operating expenses", "Staff salaries & wages", "", ""),
        ("2026-04-30", acct_main, "MONTHLY ACCOUNT MAINTENANCE FEE", 15.0, 0.0, 0.0, "METRO BANK", "", "Bank fees", "Bank maintenance charges", "", ""),

        # MAY 2026
        ("2026-05-04", acct_pos, "POS CLIENT SETTLEMENT - FACIAL & BROW SERVICES", 0.0, 3580.0, 51.0, "CLIENT CARD BATCH", "", "Revenue", "Studio services", "", ""),
        ("2026-05-06", acct_main, "STUDIO RENT SUITE 204 - MAY", 1450.0, 0.0, 0.0, "OAKWOOD PROPERTIES", "", "Operating expenses", "Studio rent & lease", "", ""),
        ("2026-05-11", acct_main, "CONTRACTOR PAYOUT - NAIL & BROW SPECIALIST", 710.0, 0.0, 0.0, "MAYA LIN", "", "Direct costs", "Service contractors & technicians", "", ""),
        ("2026-05-15", acct_pos, "POS CLIENT SETTLEMENT - SKINCARE RETAIL & TREATMENTS", 0.0, 3190.0, 45.0, "CLIENT CARD BATCH", "", "Revenue", "Studio services", "", ""),
        ("2026-05-18", acct_pos, "TRANSFER TO METRO CHECKING 884102 AURA STUDIO", 6505.0, 0.0, 5.0, "AURA STUDIO LLC", "884102", "Internal transfers", "Matched transfer", "", ""),
        ("2026-05-18", acct_main, "TRANSFER FROM HORIZON WALLET 440918 AURA STUDIO", 0.0, 6500.0, 0.0, "AURA STUDIO LLC", "440918", "Internal transfers", "Matched transfer", "", ""),
        ("2026-05-20", acct_main, "META PLATFORMS ADS & LOCAL BRIDAL EXPO SPONSORSHIP", 590.0, 0.0, 0.0, "META ADS", "", "Operating expenses", "Advertising & social media", "", ""),
        ("2026-05-22", acct_main, "SPARKLE COMMERCIAL CLEANING & SANITATION", 180.0, 0.0, 0.0, "SPARKLE CLEAN", "", "Operating expenses", "Cleaning & hygiene", "", ""),
        ("2026-05-25", acct_main, "INBOUND ZELLE TRANSFER - FROM VELVET LOUNGE CO", 0.0, 650.0, 0.0, "VELVET LOUNGE CO", "9920144", "Needs Review", "Unidentified receipts", "Receipt nature", "Recurring receipt from Velvet Lounge Co — confirm classification."),
        ("2026-05-27", acct_main, "INBOUND ACH - CLAIRE BENNETT DEPOSIT", 0.0, 420.0, 0.0, "CLAIRE BENNETT", "5501923", "Needs Review", "Unidentified receipts", "Receipt nature", "Is $420 from Claire Bennett a bridal group booking deposit or completed service sale?"),
        ("2026-05-29", acct_main, "PAYROLL ACH - STUDIO STAFF", 1450.0, 0.0, 0.0, "STAFF PAYROLL", "", "Operating expenses", "Staff salaries & wages", "", ""),
        ("2026-05-31", acct_main, "MONTHLY ACCOUNT MAINTENANCE FEE", 15.0, 0.0, 0.0, "METRO BANK", "", "Bank fees", "Bank maintenance charges", "", ""),

        # JUNE 2026
        ("2026-06-03", acct_pos, "POS CLIENT SETTLEMENT - SUMMER GLOW PACKAGES", 0.0, 3760.0, 53.0, "CLIENT CARD BATCH", "", "Revenue", "Studio services", "", ""),
        ("2026-06-06", acct_main, "STUDIO RENT SUITE 204 - JUN", 1450.0, 0.0, 0.0, "OAKWOOD PROPERTIES", "", "Operating expenses", "Studio rent & lease", "", ""),
        ("2026-06-10", acct_main, "LUMINA SKIN & LASH WHOLESALE ORDER", 960.0, 0.0, 0.0, "LUMINA SUPPLIES", "", "Direct costs", "Treatment supplies & consumables", "", ""),
        ("2026-06-14", acct_pos, "POS CLIENT SETTLEMENT - STUDIO TREATMENTS", 0.0, 3410.0, 48.0, "CLIENT CARD BATCH", "", "Revenue", "Studio services", "", ""),
        ("2026-06-17", acct_pos, "TRANSFER TO METRO CHECKING 884102 AURA STUDIO", 7005.0, 0.0, 5.0, "AURA STUDIO LLC", "884102", "Internal transfers", "Matched transfer", "", ""),
        ("2026-06-17", acct_main, "TRANSFER FROM HORIZON WALLET 440918 AURA STUDIO", 0.0, 7000.0, 0.0, "AURA STUDIO LLC", "440918", "Internal transfers", "Matched transfer", "", ""),
        ("2026-06-20", acct_main, "CITY POWER & HVAC FILTER MAINTENANCE", 310.0, 0.0, 0.0, "CITY POWER", "", "Operating expenses", "Electricity & utilities", "", ""),
        ("2026-06-22", acct_main, "FIBER INTERNET & STUDIO BOOKING SOFTWARE", 115.0, 0.0, 0.0, "FIBERNET TELECOM", "", "Operating expenses", "Internet, phone & booking software", "", ""),
        ("2026-06-24", acct_main, "OUTBOUND ZELLE - APEX STUDIO SUPPLY CO", 360.0, 0.0, 0.0, "APEX STUDIO SUPPLY", "6610982", "Needs Review", "Unidentified payments", "Payment nature", "Confirm whether Apex Studio Supply purchase was consumable treatment supplies or capital equipment."),
        ("2026-06-28", acct_main, "PAYROLL ACH - STUDIO STAFF", 1450.0, 0.0, 0.0, "STAFF PAYROLL", "", "Operating expenses", "Staff salaries & wages", "", ""),
        ("2026-06-30", acct_main, "MONTHLY ACCOUNT MAINTENANCE FEE", 15.0, 0.0, 0.0, "METRO BANK", "", "Bank fees", "Bank maintenance charges", "", ""),
    ]

    months = [
        ("2026-01", "2026-01-01", "2026-01-31"),
        ("2026-02", "2026-02-01", "2026-02-28"),
        ("2026-03", "2026-03-01", "2026-03-31"),
        ("2026-04", "2026-04-01", "2026-04-30"),
        ("2026-05", "2026-05-01", "2026-05-31"),
        ("2026-06", "2026-06-01", "2026-06-30"),
    ]

    running_bal = {
        acct_main: 4250.00,
        acct_pos: 680.00,
    }
    docs = []
    doc_idx = 1

    for ym, start_d, end_d in months:
        for acct in (acct_main, acct_pos):
            m_items = [x for x in plan if x[0].startswith(ym) and x[1] == acct]
            opening = round(running_bal[acct], 2)
            rows = []
            for seq, item in enumerate(m_items, 1):
                dt, _, desc, dr, cr, chg, cp, src_a, sec, cat, rtype, quest = item
                # For POS credits with withheld fee, net credit settled is cr - chg; for debits with fee, dr includes chg
                settled_cr = round(cr - chg, 2) if cr > 0 else 0.0
                settled_dr = round(dr, 2)
                before = round(running_bal[acct], 2)
                after = round(before + settled_cr - settled_dr, 2)
                running_bal[acct] = after
                rows.append(
                    {
                        "seq": seq,
                        "date": dt,
                        "datetime": f"{dt}T12:00:00",
                        "reference": f"REF-{ym.replace('-', '')}-{doc_idx:02d}-{seq:03d}",
                        "description": desc,
                        "debit": settled_dr,
                        "credit": settled_cr,
                        "before": before,
                        "balance": after,
                        "charge": chg,
                        "principal": cr if cr > 0 else round(dr - chg, 2),
                        "counterparty": cp,
                        "source_account": src_a,
                        "pages": "1",
                        "statement": f"D{doc_idx:02d}",
                        "account": acct,
                        "file": f"demo_statements/{acct.replace(' ', '_')}_{ym}.pdf",
                        "section": sec,
                        "category": cat,
                        "review_type": rtype,
                        "question": quest,
                    }
                )
            docs.append(
                {
                    "id": f"D{doc_idx:02d}",
                    "file": f"demo_statements/{acct.replace(' ', '_')}_{ym}.pdf",
                    "sha256": f"synthetic-demo-sha256-{doc_idx:04d}",
                    "encrypted": False,
                    "page_count": 2,
                    "account": acct,
                    "currency": "USD",
                    "owner": "AURA STUDIO LLC",
                    "opening": opening,
                    "closing": round(running_bal[acct], 2),
                    "credits": round(sum(r["credit"] for r in rows), 2),
                    "debits": round(sum(r["debit"] for r in rows), 2),
                    "start": start_d,
                    "end": end_d,
                    "rows": rows,
                }
            )
            doc_idx += 1

    return docs


def main() -> None:
    DEMO_WORK.mkdir(parents=True, exist_ok=True)
    DEMO_OUT.mkdir(parents=True, exist_ok=True)
    DOCS_DIR.mkdir(parents=True, exist_ok=True)

    parsed_path = DEMO_WORK / "parsed_statements.json"
    books_path = DEMO_WORK / "books_data.json"
    cls_path = DEMO_WORK / "classifications.json"

    parsed_docs = build_synthetic_parsed_statements()
    parsed_path.write_text(json.dumps(parsed_docs, indent=2), encoding="utf-8")

    cls_cfg = {
        "schema_version": "1.0",
        "business_profile": {
            "business_name": "Aura Studio & Spa (Demo)",
            "legal_entity": "AURA STUDIO LLC",
            "business_type": "service",
            "industry": "Beauty & Wellness Studio",
            "currency": "USD",
            "direct_cost_label": "Cost of Providing Services",
            "reporting_basis": "Cash basis",
        },
        "category_group_overrides": {},
        "counterparty_rules": [],
        "transaction_overrides": {},
    }
    cls_path.write_text(json.dumps(cls_cfg, indent=2), encoding="utf-8")

    subprocess.run(
        [
            sys.executable,
            str(SCRIPTS_DIR / "reconcile_ledger.py"),
            "--parsed", str(parsed_path),
            "--out", str(books_path),
            "--entity-keyword", "AURA STUDIO",
        ],
        check=True,
    )

    subprocess.run(
        [
            sys.executable,
            str(SCRIPTS_DIR / "build_all.py"),
            "--books", str(books_path),
            "--classifications", str(cls_path),
            "--parsed", str(parsed_path),
            "--out-dir", str(DEMO_OUT),
            "--slug", "Aura_Studio_2026",
            "--html-out", str(DOCS_DIR / "index.html"),
        ],
        check=True,
    )
    print(f"Generated GitHub Pages live demo -> {DOCS_DIR / 'index.html'}")


if __name__ == "__main__":
    main()
