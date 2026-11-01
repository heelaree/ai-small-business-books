#!/usr/bin/env python3
"""Portable multi-sheet Excel Cash Books & CSV Ledger generator using standard `openpyxl`.

Reorganizes tabs in small-business-first priority order:
  1. Executive Overview  2. Summary P&L  3. Financial Health
  4. Detailed P&L  5. Cash Flow  6. Master Ledger  7. Needs Review
  8. Account Cash  9. Transfer Matches  10. Statement Inventory
  11. Source Rows  12. Duplicate Exclusions  13. Reconciliations  14. ReadMe
"""

from __future__ import annotations

import argparse
import csv
from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from ledger_engine import OPEX_BUCKETS, PNL_SECTIONS, load_and_compute

MONEY_FMT = '#,##0.00;(#,##0.00);"-"'
DIFF_FMT = "0.00;(0.00);0.00"
PCT_FMT = "0.0%"
DATE_FMT = "dd-mmm-yyyy"
MONTH_FMT = "mmm-yy"

FONT_FAMILY = "Calibri"
FILL_HEADER = PatternFill("solid", fgColor="1E293B")
FILL_SUBTOTAL = PatternFill("solid", fgColor="E2E8F0")
FILL_HIGHLIGHT = PatternFill("solid", fgColor="DBEAFE")
FILL_AMBER = PatternFill("solid", fgColor="FEF3C7")
FILL_ZEBRA = PatternFill("solid", fgColor="F8FAFC")
FILL_RED_ERR = PatternFill("solid", fgColor="FEE2E2")

FONT_TITLE = Font(name=FONT_FAMILY, size=15, bold=True, color="0F172A")
FONT_SUBTITLE = Font(name=FONT_FAMILY, size=10, italic=True, color="475569")
FONT_HEADER = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
FONT_BODY = Font(name=FONT_FAMILY, size=10, color="0F172A")
FONT_BOLD = Font(name=FONT_FAMILY, size=10, bold=True, color="0F172A")
FONT_BLUE_INPUT = Font(name=FONT_FAMILY, size=10, color="1D4ED8")
FONT_RED_ERR = Font(name=FONT_FAMILY, size=10, bold=True, color="B91C1C")

THIN_BORDER = Border(
    bottom=Side(style="thin", color="E2E8F0"),
)
SUBTOTAL_BORDER = Border(
    top=Side(style="thin", color="94A3B8"),
    bottom=Side(style="double", color="64748B"),
)


def parse_iso_date(s: str | None) -> date | None:
    if not s:
        return None
    return datetime.strptime(s[:10], "%Y-%m-%d").date()


def setup_sheet(ws, title: str, subtitle: str, tab_color: str) -> None:
    ws.views.sheetView[0].showGridLines = False
    ws.sheet_properties.tabColor = tab_color
    ws["A2"] = title
    ws["A2"].font = FONT_TITLE
    ws.row_dimensions[2].height = 24
    ws["A3"] = subtitle
    ws["A3"].font = FONT_SUBTITLE
    ws.row_dimensions[3].height = 18


def style_header_row(ws, row_idx: int, max_col: int, height: int = 30) -> None:
    ws.row_dimensions[row_idx].height = height
    for c in range(1, max_col + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def set_col_widths(ws, widths: dict[int | str, float], default_width: float | None = None, max_col: int = 0) -> None:
    if default_width and max_col:
        for c in range(1, max_col + 1):
            ws.column_dimensions[get_column_letter(c)].width = default_width
    for k, w in widths.items():
        col_letter = get_column_letter(k) if isinstance(k, int) else k
        ws.column_dimensions[col_letter].width = w


def add_nonzero_red_rule(ws, cell_range: str) -> None:
    ws.conditional_formatting.add(
        cell_range,
        CellIsRule(operator="notEqual", formula=["0"], fill=FILL_RED_ERR, font=FONT_RED_ERR),
    )


def build_summary_group_formula(row_idx: int, direct_label: str) -> str:
    dl = direct_label.replace('"', '""')
    m = f"M{row_idx}"
    n = f"N{row_idx}"
    i = f"I{row_idx}"
    return (
        f'=IF({m}="Revenue","Revenue (Net of Refunds)",'
        f'IF({m}="Direct costs","{dl}",'
        f'IF({m}="Taxes","Taxes & Statutory Levies",'
        f'IF({m}="Bank fees","Bank & Payment Fees",'
        f'IF({m}="Needs Review",IF({i}>=0,"Unclassified Receipts","Unclassified Payments"),'
        f'IF({m}="Internal transfers","Matched Internal Transfers",'
        f'IF({m}="Operating expenses",'
        f'IF(OR(ISNUMBER(SEARCH("salar",{n})),ISNUMBER(SEARCH("wage",{n})),ISNUMBER(SEARCH("staff",{n})),ISNUMBER(SEARCH("welfare",{n})),ISNUMBER(SEARCH("training",{n})),ISNUMBER(SEARCH("intern",{n}))),"People & Staff",'
        f'IF(OR(ISNUMBER(SEARCH("electric",{n})),ISNUMBER(SEARCH("power",{n})),ISNUMBER(SEARCH("fuel",{n})),ISNUMBER(SEARCH("diesel",{n})),ISNUMBER(SEARCH("generator",{n})),ISNUMBER(SEARCH("repair",{n})),ISNUMBER(SEARCH("mainten",{n}))),"Power, Fuel & Maintenance",'
        f'IF(OR(ISNUMBER(SEARCH("advertis",{n})),ISNUMBER(SEARCH("ads",{n})),ISNUMBER(SEARCH("marketing",{n})),ISNUMBER(SEARCH("promo",{n})),ISNUMBER(SEARCH("sponsor",{n})),ISNUMBER(SEARCH("brand",{n}))),"Marketing & Promotion",'
        f'IF(OR(ISNUMBER(SEARCH("rent",{n})),ISNUMBER(SEARCH("lease",{n})),ISNUMBER(SEARCH("space",{n})),ISNUMBER(SEARCH("clean",{n})),ISNUMBER(SEARCH("securit",{n})),ISNUMBER(SEARCH("water",{n})),ISNUMBER(SEARCH("hygiene",{n}))),"Premises, Cleaning & Security",'
        f'"Admin, Comms & Operations")))),"Non-P&L Cash Movements")))))))'
    )


def build_workbook(computed: dict[str, Any]) -> openpyxl.Workbook:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    p = computed["profile"]
    cur = computed["currency"]
    direct_label = p["direct_cost_label"]
    months = computed["months"]
    pflags = computed["month_partial_flags"]
    ledger = computed["ledger"]
    all_src = computed["all_source_rows"]
    valid_docs = computed["valid_documents"]
    docs = computed["documents"]
    transfers = computed["transfers"]
    duplicates = computed["duplicates"]
    review_items = computed["review"]
    scope_issues = computed["scope_issues"]
    acct_cash = computed["account_cash"]

    sheet_specs = [
        ("Executive Overview", "1E40AF"),
        ("Summary P&L", "059669"),
        ("Financial Health", "0284C7"),
        ("Detailed P&L", "334155"),
        ("Cash Flow", "334155"),
        ("Master Ledger", "334155"),
        ("Needs Review", "D97706"),
        ("Account Cash", "64748B"),
        ("Transfer Matches", "64748B"),
        ("Statement Inventory", "64748B"),
        ("Source Rows", "64748B"),
        ("Duplicate Exclusions", "64748B"),
        ("Reconciliations", "64748B"),
        ("ReadMe", "94A3B8"),
    ]
    ws_map = {name: wb.create_sheet(title=name) for name, _ in sheet_specs}
    tab_colors = dict(sheet_specs)

    L_END = len(ledger) + 5
    R_END = len(all_src) + 5
    AC_END = len(acct_cash) + 5
    NR_END = len(scope_issues) + len(review_items) + 5
    M_COUNT = len(months)
    month_dates = [parse_iso_date(f"{m}-01") for m in months]

    def lr(col: str) -> str:
        return f"'Master Ledger'!${col}$6:${col}${L_END}"

    def sr(col: str) -> str:
        return f"'Source Rows'!${col}$6:${col}${R_END}"

    # -------------------------------------------------------------------------
    # 6. MASTER LEDGER (populate core ledger ranges first)
    # -------------------------------------------------------------------------
    ws = ws_map["Master Ledger"]
    setup_sheet(
        ws,
        f"{p['business_name']} | Categorized Master Cash Ledger",
        f"{cur}. Positive = receipt; negative = payment. Edit Section/Category (yellow cells) to update all P&L & Health tabs.",
        tab_colors["Master Ledger"],
    )
    ml_headers = [
        "Transaction ID", "Bank Txn ID", "Date", "Month", "Account",
        "Original Description", "Component", f"Amount ({cur})", f"Signed ({cur})",
        f"Credit ({cur})", f"Debit ({cur})", "Direction", "Section", "Category",
        "Summary Group", "Transfer ID", "Review ID", "Classification Basis",
        f"P&L Impact ({cur})", f"External Cash ({cur})", "Statement ID",
        "Source Filename", "Page(s)", "Bank Reference", "Question / Owner Note",
    ]
    ws.append([])  # row 4 spacer
    for c_idx, h_text in enumerate(ml_headers, 1):
        ws.cell(row=5, column=c_idx, value=h_text)
    style_header_row(ws, 5, len(ml_headers))

    review_id_map = {r["ledger_id"]: r["id"] for r in review_items}
    for idx, e in enumerate(ledger, 6):
        basis = (
            "Owner confirmed"
            if e.get("user_classified")
            else ("Unclassified" if e["section"] == "Needs Review" else ("Tentative - review" if e.get("question") else "Narration supported"))
        )
        note_or_q = e.get("user_note") or e.get("question", "")
        row_vals = [
            e["id"], e["raw_id"], parse_iso_date(e["date"]), f"=DATE(YEAR(C{idx}),MONTH(C{idx}),1)",
            e["account"], e["description"], e["component"], f"=ABS(I{idx})", e["amount"],
            f"=MAX(I{idx},0)", f"=MAX(-I{idx},0)", f'=IF(I{idx}>0,"Credit","Debit")',
            e["section"], e["category"], build_summary_group_formula(idx, direct_label),
            e.get("transfer_id", ""), review_id_map.get(e["id"], ""), basis,
            f'=IF(OR(M{idx}="Revenue",M{idx}="Direct costs",M{idx}="Operating expenses",M{idx}="Taxes",M{idx}="Bank fees"),I{idx},0)',
            f'=IF(M{idx}="Internal transfers",0,I{idx})',
            e.get("statement", ""), e.get("file", ""), e.get("pages", ""), e.get("reference", ""), note_or_q,
        ]
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row=idx, column=c_idx, value=val)
            cell.font = FONT_BLUE_INPUT if c_idx == 9 else FONT_BODY
            cell.border = THIN_BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=(c_idx in (6, 25)))
            if c_idx in (3, 4):
                cell.number_format = DATE_FMT if c_idx == 3 else MONTH_FMT
            elif c_idx in (8, 9, 10, 11, 19, 20):
                cell.number_format = MONEY_FMT
            if c_idx in (13, 14):
                cell.fill = FILL_AMBER
        ws.row_dimensions[idx].height = 38

    dv_sec = DataValidation(
        type="list",
        formula1='"Revenue,Direct costs,Operating expenses,Taxes,Bank fees,Internal transfers,Owner transactions,Loans,Asset purchases,Unmatched own-account transfers,Needs Review"',
        allow_blank=False,
    )
    ws.add_data_validation(dv_sec)
    dv_sec.add(f"M6:M{L_END}")
    ws.freeze_panes = "C6"
    set_col_widths(
        ws,
        {"A": 16, "B": 14, "C": 14, "D": 12, "E": 24, "F": 58, "G": 15, "H": 16, "I": 16, "J": 16, "K": 16, "L": 12, "M": 26, "N": 28, "O": 28, "P": 13, "Q": 12, "R": 20, "S": 16, "T": 16, "U": 13, "V": 46, "W": 10, "X": 34, "Y": 52},
    )

    # -------------------------------------------------------------------------
    # 2. SUMMARY P&L (SHORT P&L + CASH BRIDGE)
    # -------------------------------------------------------------------------
    ws = ws_map["Summary P&L"]
    setup_sheet(
        ws,
        f"{p['business_name']} | Summary Profit & Loss (Short P&L)",
        f"{cur}, {p['reporting_basis'].lower()}. Concise owner summary with Gross Profit, Operating Profit, and Cash Bridge.",
        tab_colors["Summary P&L"],
    )
    sp_last_col = M_COUNT + 2  # Col A = Line item, Cols B.. = Months, Last Col = Cumulative
    sp_cum_col = get_column_letter(sp_last_col)
    sp_last_m_col = get_column_letter(sp_last_col - 1)

    ws.cell(row=5, column=1, value="Summary P&L Line Item")
    for m_i, md in enumerate(month_dates, 2):
        c = ws.cell(row=5, column=m_i, value=md)
        c.number_format = 'mmm-yy" (partial)"' if pflags.get(months[m_i - 2]) else MONTH_FMT
    ws.cell(row=5, column=sp_last_col, value="Cumulative")
    style_header_row(ws, 5, sp_last_col)

    sp_rows: dict[str, int] = {}
    r_ptr = 6

    def add_sp_formula_row(key: str, label: str, m_formula_fn, style: str = "normal") -> int:
        nonlocal r_ptr
        sp_rows[key] = r_ptr
        row_font = FONT_BOLD if style in ("subtotal", "highlight", "bridge_total") else FONT_BODY
        lbl_cell = ws.cell(row=r_ptr, column=1, value=label)
        lbl_cell.font = row_font
        for m_i in range(2, sp_last_col):
            col_l = get_column_letter(m_i)
            cell = ws.cell(row=r_ptr, column=m_i, value=m_formula_fn(col_l))
            cell.number_format = MONEY_FMT
            cell.font = row_font
        cum_cell = ws.cell(row=r_ptr, column=sp_last_col, value=f"=SUM(B{r_ptr}:{sp_last_m_col}{r_ptr})")
        cum_cell.number_format = MONEY_FMT
        cum_cell.font = row_font

        for c_i in range(1, sp_last_col + 1):
            c = ws.cell(row=r_ptr, column=c_i)
            if style == "subtotal":
                c.fill = FILL_SUBTOTAL
                c.border = SUBTOTAL_BORDER
            elif style == "highlight":
                c.fill = FILL_HIGHLIGHT
                c.border = SUBTOTAL_BORDER
            elif style == "bridge":
                c.fill = FILL_AMBER
            elif style == "bridge_total":
                c.fill = FILL_SUBTOTAL
                c.border = SUBTOTAL_BORDER
            else:
                c.border = THIN_BORDER
        ws.row_dimensions[r_ptr].height = 24
        r_ptr += 1
        return r_ptr - 1

    add_sp_formula_row("rev", "Revenue (Net of Refunds)", lambda c: f'=SUMIFS({lr("I")},{lr("D")},{c}$5,{lr("M")},"Revenue")')
    add_sp_formula_row("dc", direct_label, lambda c: f'=-SUMIFS({lr("I")},{lr("D")},{c}$5,{lr("M")},"Direct costs")')
    add_sp_formula_row("gp", "GROSS PROFIT / (LOSS)", lambda c: f"={c}{sp_rows['rev']}-{c}{sp_rows['dc']}", "subtotal")
    r_ptr += 1

    opex_start = r_ptr
    for b in OPEX_BUCKETS:
        add_sp_formula_row(
            f"opex_{b}",
            f"Operating Expenses — {b}",
            lambda c, b=b: f'=-SUMIFS({lr("I")},{lr("D")},{c}$5,{lr("M")},"Operating expenses",{lr("O")},"{b}")',
        )
    opex_end = r_ptr - 1
    add_sp_formula_row("tot_opex", "Total Operating Expenses", lambda c: f"=SUM({c}{opex_start}:{c}{opex_end})", "subtotal")
    add_sp_formula_row("op_profit", "OPERATING PROFIT / (LOSS)", lambda c: f"={c}{sp_rows['gp']}-{c}{sp_rows['tot_opex']}", "subtotal")
    r_ptr += 1

    add_sp_formula_row("tax", "Taxes & Statutory Levies", lambda c: f'=-SUMIFS({lr("I")},{lr("D")},{c}$5,{lr("M")},"Taxes")')
    add_sp_formula_row("fee", "Bank & Payment Fees", lambda c: f'=-SUMIFS({lr("I")},{lr("D")},{c}$5,{lr("M")},"Bank fees")')
    add_sp_formula_row("tax_fee", "Total Taxes & Bank Fees", lambda c: f"={c}{sp_rows['tax']}+{c}{sp_rows['fee']}", "subtotal")
    r_ptr += 1

    add_sp_formula_row(
        "net_profit",
        "PROVISIONAL NET PROFIT / (LOSS)",
        lambda c: f"={c}{sp_rows['op_profit']}-{c}{sp_rows['tax_fee']}",
        "highlight",
    )
    r_ptr += 1

    # Cash Reconciliation Bridge
    add_sp_formula_row(
        "unclass_in",
        "Cash Bridge: Unclassified Receipts (Excluded from P&L)",
        lambda c: f'=SUMIFS({lr("J")},{lr("D")},{c}$5,{lr("M")},"Needs Review")',
        "bridge",
    )
    add_sp_formula_row(
        "unclass_out",
        "Cash Bridge: Unclassified Payments (Excluded from P&L)",
        lambda c: f'=-SUMIFS({lr("K")},{lr("D")},{c}$5,{lr("M")},"Needs Review")',
        "bridge",
    )
    add_sp_formula_row(
        "non_pnl",
        "Cash Bridge: Non-P&L Cash Movements (Loans, Assets, Owner & Other Accounts)",
        lambda c: f'=SUMIFS({lr("I")},{lr("D")},{c}$5,{lr("O")},"Non-P&L Cash Movements")',
        "bridge",
    )
    add_sp_formula_row(
        "cash_change",
        "ACTUAL NET BANK CASH CHANGE",
        lambda c: f"={c}{sp_rows['net_profit']}+{c}{sp_rows['unclass_in']}+{c}{sp_rows['unclass_out']}+{c}{sp_rows['non_pnl']}",
        "bridge_total",
    )

    ws.freeze_panes = "B6"
    set_col_widths(ws, {1: 54}, default_width=18, max_col=sp_last_col)

    # -------------------------------------------------------------------------
    # 4. DETAILED P&L
    # -------------------------------------------------------------------------
    ws = ws_map["Detailed P&L"]
    setup_sheet(
        ws,
        f"{p['business_name']} | Detailed Monthly Profit & Loss",
        f"{cur}, {p['reporting_basis'].lower()}. Full category schedule with Gross Profit, Operating Profit, and dynamic catch-all rows.",
        tab_colors["Detailed P&L"],
    )
    dp_last_col = M_COUNT + 3  # Col A = Category, Col B = Section, Cols C.. = Months, Last = Cumulative
    dp_cum_col = get_column_letter(dp_last_col)
    dp_last_m_col = get_column_letter(dp_last_col - 1)

    ws.cell(row=5, column=1, value="Category")
    ws.cell(row=5, column=2, value="Section")
    for m_i, md in enumerate(month_dates, 3):
        c = ws.cell(row=5, column=m_i, value=md)
        c.number_format = 'mmm-yy" (partial)"' if pflags.get(months[m_i - 3]) else MONTH_FMT
    ws.cell(row=5, column=dp_last_col, value="Cumulative")
    style_header_row(ws, 5, dp_last_col)

    dp_totals: dict[str, int] = {}
    dr = 6
    for sec in PNL_SECTIONS:
        sec_start = dr
        sign_prefix = "" if sec == "Revenue" else "-"
        cats = [x["category"] for x in computed["detailed_pnl"][sec]]
        for cat in cats:
            ws.cell(row=dr, column=1, value=cat).font = FONT_BODY
            ws.cell(row=dr, column=2, value=sec).font = FONT_BODY
            for m_i in range(3, dp_last_col):
                cl = get_column_letter(m_i)
                cell = ws.cell(
                    row=dr, column=m_i,
                    value=f'={sign_prefix}SUMIFS({lr("I")},{lr("D")},{cl}$5,{lr("M")},$B{dr},{lr("N")},$A{dr})',
                )
                cell.number_format = MONEY_FMT
                cell.font = FONT_BODY
            cc = ws.cell(row=dr, column=dp_last_col, value=f"=SUM(C{dr}:{dp_last_m_col}{dr})")
            cc.number_format = MONEY_FMT
            cc.font = FONT_BODY
            ws.row_dimensions[dr].height = 22
            dr += 1

        # Dynamic catch-all row for newly typed category names
        ws.cell(row=dr, column=1, value="Other categories (auto-captured)").font = FONT_BODY
        ws.cell(row=dr, column=2, value=sec).font = FONT_BODY
        for m_i in range(3, dp_last_col):
            cl = get_column_letter(m_i)
            prev_sum = f"SUM({cl}{sec_start}:{cl}{dr - 1})" if dr > sec_start else "0"
            cell = ws.cell(
                row=dr, column=m_i,
                value=f'=ROUND({sign_prefix}SUMIFS({lr("I")},{lr("D")},{cl}$5,{lr("M")},$B{dr})-{prev_sum},2)',
            )
            cell.number_format = MONEY_FMT
            cell.font = FONT_BODY
        cc = ws.cell(row=dr, column=dp_last_col, value=f"=SUM(C{dr}:{dp_last_m_col}{dr})")
        cc.number_format = MONEY_FMT
        cc.font = FONT_BODY
        ws.row_dimensions[dr].height = 22
        dr += 1

        dp_totals[sec] = dr
        ws.cell(row=dr, column=1, value=f"Total {sec.lower()}")
        for c_i in range(1, dp_last_col + 1):
            cell = ws.cell(row=dr, column=c_i)
            if c_i >= 3:
                cl = get_column_letter(c_i)
                cell.value = f"=SUM({cl}{sec_start}:{cl}{dr - 1})"
                cell.number_format = MONEY_FMT
            cell.font = FONT_BOLD
            cell.fill = FILL_SUBTOTAL
            cell.border = SUBTOTAL_BORDER
        ws.row_dimensions[dr].height = 24
        dr += 1

        if sec == "Direct costs":
            dp_totals["Gross profit"] = dr
            ws.cell(row=dr, column=1, value="GROSS PROFIT / (LOSS)")
            for c_i in range(1, dp_last_col + 1):
                cell = ws.cell(row=dr, column=c_i)
                if c_i >= 3:
                    cl = get_column_letter(c_i)
                    cell.value = f"={cl}{dp_totals['Revenue']}-{cl}{dp_totals['Direct costs']}"
                    cell.number_format = MONEY_FMT
                cell.font = FONT_BOLD
                cell.fill = FILL_HIGHLIGHT
                cell.border = SUBTOTAL_BORDER
            ws.row_dimensions[dr].height = 24
            dr += 1

        if sec == "Operating expenses":
            dp_totals["Operating profit"] = dr
            ws.cell(row=dr, column=1, value="OPERATING PROFIT / (LOSS)")
            for c_i in range(1, dp_last_col + 1):
                cell = ws.cell(row=dr, column=c_i)
                if c_i >= 3:
                    cl = get_column_letter(c_i)
                    cell.value = f"={cl}{dp_totals['Gross profit']}-{cl}{dp_totals['Operating expenses']}"
                    cell.number_format = MONEY_FMT
                cell.font = FONT_BOLD
                cell.fill = FILL_HIGHLIGHT
                cell.border = SUBTOTAL_BORDER
            ws.row_dimensions[dr].height = 24
            dr += 1

        dr += 1

    dp_profit_row = dr
    ws.cell(row=dr, column=1, value="PROVISIONAL NET PROFIT / (LOSS)")
    for c_i in range(1, dp_last_col + 1):
        cell = ws.cell(row=dr, column=c_i)
        if c_i >= 3:
            cl = get_column_letter(c_i)
            cell.value = f"={cl}{dp_totals['Operating profit']}-{cl}{dp_totals['Taxes']}-{cl}{dp_totals['Bank fees']}"
            cell.number_format = MONEY_FMT
        cell.font = FONT_BOLD
        cell.fill = FILL_HIGHLIGHT
        cell.border = SUBTOTAL_BORDER
    ws.row_dimensions[dr].height = 26
    dr += 2

    dp_un_in_row = dr
    dp_un_out_row = dr + 1
    for offset, (lbl, lcol) in enumerate([("Unclassified receipts excluded", "J"), ("Unclassified payments excluded", "K")]):
        rr = dr + offset
        ws.cell(row=rr, column=1, value=lbl).font = FONT_BODY
        ws.cell(row=rr, column=1).fill = FILL_AMBER
        ws.cell(row=rr, column=2).fill = FILL_AMBER
        for m_i in range(3, dp_last_col):
            cl = get_column_letter(m_i)
            cell = ws.cell(row=rr, column=m_i, value=f'=SUMIFS({lr(lcol)},{lr("D")},{cl}$5,{lr("M")},"Needs Review")')
            cell.number_format = MONEY_FMT
            cell.fill = FILL_AMBER
        cc = ws.cell(row=rr, column=dp_last_col, value=f"=SUM(C{rr}:{dp_last_m_col}{rr})")
        cc.number_format = MONEY_FMT
        cc.fill = FILL_AMBER
        ws.row_dimensions[rr].height = 23

    ws.freeze_panes = "C6"
    set_col_widths(ws, {1: 38, 2: 24}, default_width=18, max_col=dp_last_col)

    # -------------------------------------------------------------------------
    # 8. ACCOUNT CASH (needed for Cash Flow & Overview opening/closing links)
    # -------------------------------------------------------------------------
    ws = ws_map["Account Cash"]
    setup_sheet(
        ws,
        "Monthly Account Cash Roll-Forwards",
        f"{cur}. Independent statement opening/closing balances reconciled to allocated ledger credits and debits.",
        tab_colors["Account Cash"],
    )
    ac_headers = [
        "Month", "Account", "Balance Source", "Coverage Start", "Coverage End",
        f"Opening ({cur})", f"Allocated Credits ({cur})", f"Allocated Debits ({cur})",
        f"Calculated Closing ({cur})", f"Source Closing ({cur})", "Closing Difference", "Coverage Note",
    ]
    for c_i, h_t in enumerate(ac_headers, 1):
        ws.cell(row=5, column=c_i, value=h_t)
    style_header_row(ws, 5, len(ac_headers))

    for idx, r in enumerate(acct_cash, 6):
        vals = [
            parse_iso_date(f"{r['month']}-01"), r["account"], r["source_id"],
            parse_iso_date(r["start"]), parse_iso_date(r["end"]), r["opening"],
            f'=SUMIFS({lr("J")},{lr("D")},A{idx},{lr("E")},B{idx})',
            f'=SUMIFS({lr("K")},{lr("D")},A{idx},{lr("E")},B{idx})',
            f"=F{idx}+G{idx}-H{idx}", r["source_closing"], f"=ROUND(I{idx}-J{idx},2)", r["coverage"],
        ]
        for c_i, v in enumerate(vals, 1):
            c = ws.cell(row=idx, column=c_i, value=v)
            c.font = FONT_BODY
            c.border = THIN_BORDER
            if c_i == 1:
                c.number_format = MONTH_FMT
            elif c_i in (4, 5):
                c.number_format = DATE_FMT
            elif c_i in (6, 7, 8, 9, 10):
                c.number_format = MONEY_FMT
            elif c_i == 11:
                c.number_format = DIFF_FMT
        ws.row_dimensions[idx].height = 24
    add_nonzero_red_rule(ws, f"K6:K{AC_END}")
    ws.freeze_panes = "C6"
    set_col_widths(ws, {1: 14, 2: 26, 3: 15, 4: 15, 5: 15, 6: 18, 7: 20, 8: 20, 9: 20, 10: 20, 11: 18, 12: 38})

    # -------------------------------------------------------------------------
    # 5. CASH FLOW
    # -------------------------------------------------------------------------
    ws = ws_map["Cash Flow"]
    setup_sheet(
        ws,
        f"{p['business_name']} | Monthly External Cash Flow",
        f"{cur}. Matched internal transfers excluded from external cash. Cumulative balances never sum monthly stocks.",
        tab_colors["Cash Flow"],
    )
    cf_last_col = M_COUNT + 2
    cf_cum_col = get_column_letter(cf_last_col)
    cf_last_m_col = get_column_letter(cf_last_col - 1)
    ws.cell(row=5, column=1, value="Cash Flow Line Item")
    for m_i, md in enumerate(month_dates, 2):
        c = ws.cell(row=5, column=m_i, value=md)
        c.number_format = 'mmm-yy" (partial)"' if pflags.get(months[m_i - 2]) else MONTH_FMT
    ws.cell(row=5, column=cf_last_col, value="Cumulative")
    style_header_row(ws, 5, cf_last_col)

    cf_labels = {
        6: "Opening bank balances",
        8: "Revenue receipts net of refunds",
        9: f"{direct_label} paid, net",
        10: "Operating expenses paid, net",
        11: "Taxes paid, net",
        12: "Bank charges paid, net",
        13: "Known operating cash flow - provisional",
        15: "Asset purchases / recoveries",
        16: "Owner contributions / drawings",
        17: "Loan proceeds / repayments",
        18: "Own-account transfers outside supplied scope",
        19: "Unclassified receipts",
        20: "Unclassified payments",
        21: "Combined external net cash change",
        23: "Last available bank balances",
        24: "Gross matched transfer receipts excluded",
        25: "Gross matched transfer payments excluded",
    }
    for acct_idx, acct_name in enumerate(computed["accounts"]):
        cf_labels[27 + acct_idx] = f"{acct_name} closing balance"

    for r_num, lbl in cf_labels.items():
        ws.cell(row=r_num, column=1, value=lbl).font = FONT_BOLD if r_num in (6, 13, 21, 23) else FONT_BODY
        ws.row_dimensions[r_num].height = 24

    for m_i in range(2, cf_last_col):
        cl = get_column_letter(m_i)
        ws[f"{cl}6"] = f"=SUMIFS('Account Cash'!$F$6:$F${AC_END},'Account Cash'!$A$6:$A${AC_END},{cl}$5)"
        for r_num, sec in [
            (8, "Revenue"), (9, "Direct costs"), (10, "Operating expenses"), (11, "Taxes"), (12, "Bank fees"),
            (15, "Asset purchases"), (16, "Owner transactions"), (17, "Loans"), (18, "Unmatched own-account transfers"),
        ]:
            ws[f"{cl}{r_num}"] = f'=SUMIFS({lr("I")},{lr("D")},{cl}$5,{lr("M")},"{sec}")'
        ws[f"{cl}13"] = f"=SUM({cl}8:{cl}12)"
        ws[f"{cl}19"] = f'=SUMIFS({lr("J")},{lr("D")},{cl}$5,{lr("M")},"Needs Review")'
        ws[f"{cl}20"] = f'=-SUMIFS({lr("K")},{lr("D")},{cl}$5,{lr("M")},"Needs Review")'
        ws[f"{cl}21"] = f"=SUM({cl}13,{cl}15:{cl}20)"
        ws[f"{cl}23"] = f"={cl}6+{cl}21"
        ws[f"{cl}24"] = f'=SUMIFS({lr("J")},{lr("D")},{cl}$5,{lr("M")},"Internal transfers")'
        ws[f"{cl}25"] = f'=-SUMIFS({lr("K")},{lr("D")},{cl}$5,{lr("M")},"Internal transfers")'
        for acct_idx, acct_name in enumerate(computed["accounts"]):
            ws[f"{cl}{27 + acct_idx}"] = (
                f"=SUMIFS('Account Cash'!$I$6:$I${AC_END},'Account Cash'!$A$6:$A${AC_END},{cl}$5,'Account Cash'!$B$6:$B${AC_END},\"{acct_name}\")"
            )

    for r_num in cf_labels:
        if r_num == 6:
            ws[f"{cf_cum_col}{r_num}"] = "=B6"
        elif r_num == 23 or r_num >= 27:
            ws[f"{cf_cum_col}{r_num}"] = f"={cf_last_m_col}{r_num}"
        else:
            ws[f"{cf_cum_col}{r_num}"] = f"=SUM(B{r_num}:{cf_last_m_col}{r_num})"

        for c_i in range(1, cf_last_col + 1):
            c = ws.cell(row=r_num, column=c_i)
            if c_i >= 2:
                c.number_format = MONEY_FMT
            if r_num in (6, 13, 21, 23):
                c.fill = FILL_SUBTOTAL
                c.font = FONT_BOLD
                c.border = SUBTOTAL_BORDER

    ws.freeze_panes = "B6"
    set_col_widths(ws, {1: 46}, default_width=19, max_col=cf_last_col)

    # -------------------------------------------------------------------------
    # 3. FINANCIAL HEALTH (STRICTLY FACTUAL — ZERO PRESCRIPTIVE ADVICE)
    # -------------------------------------------------------------------------
    ws = ws_map["Financial Health"]
    setup_sheet(
        ws,
        f"{p['business_name']} | Financial Health & Cost Structure Report",
        "Objective cash, cost concentration, and completeness diagnostics. Contains no financial advice or recommendations.",
        tab_colors["Financial Health"],
    )
    # Table A: Liquidity & Completeness Snapshot
    ws["A5"], ws["B5"], ws["C5"] = "Liquidity & Coverage Metric", f"Value ({cur} / %)", "Factual Basis"
    style_header_row(ws, 5, 3)
    tot_cost_formula = f"='Summary P&L'!{sp_cum_col}{sp_rows['dc']}+'Summary P&L'!{sp_cum_col}{sp_rows['tot_opex']}+'Summary P&L'!{sp_cum_col}{sp_rows['tax_fee']}"
    ext_in_formula = f"SUMIFS({lr('J')},{lr('M')},\"<>Internal transfers\")"
    ext_out_formula = f"SUMIFS({lr('K')},{lr('M')},\"<>Internal transfers\")"

    fh_metrics = [
        ("Opening Combined Bank Cash", f"='Cash Flow'!{cf_cum_col}6", MONEY_FMT, "Combined opening balance at start of period"),
        ("Latest Supplied Bank Cash", f"='Cash Flow'!{cf_cum_col}23", MONEY_FMT, "Latest supplied closing balance across accounts"),
        ("Actual Net Bank Cash Change", f"='Summary P&L'!{sp_cum_col}{sp_rows['cash_change']}", MONEY_FMT, "Total external cash inflows minus outflows"),
        ("Total External Cash Inflows", f"={ext_in_formula}", MONEY_FMT, "Excludes matched internal transfers"),
        ("Total External Cash Outflows", f"={ext_out_formula}", MONEY_FMT, "Excludes matched internal transfers"),
        ("Total Classified P&L Costs", tot_cost_formula, MONEY_FMT, f"{direct_label} + Operating Expenses + Taxes & Fees"),
        ("External Inflows Classified (%)", f"=IF({ext_in_formula}>0,({ext_in_formula}-'Summary P&L'!{sp_cum_col}{sp_rows['unclass_in']})/{ext_in_formula},1)", PCT_FMT, "Share of external inflow value confirmed vs unclassified"),
        ("External Outflows Classified (%)", f"=IF({ext_out_formula}>0,({ext_out_formula}+'Summary P&L'!{sp_cum_col}{sp_rows['unclass_out']})/{ext_out_formula},1)", PCT_FMT, "Share of external outflow value confirmed vs unclassified"),
    ]
    for idx, (m_lbl, f_str, n_fmt, ctx) in enumerate(fh_metrics, 6):
        ws.cell(row=idx, column=1, value=m_lbl).font = FONT_BOLD if idx in (8, 11) else FONT_BODY
        v_cell = ws.cell(row=idx, column=2, value=f_str)
        v_cell.font = FONT_BOLD if idx in (8, 11) else FONT_BODY
        v_cell.number_format = n_fmt
        ws.cell(row=idx, column=3, value=ctx).font = FONT_BODY
        for c_i in range(1, 4):
            ws.cell(row=idx, column=c_i).border = THIN_BORDER
        ws.row_dimensions[idx].height = 23

    # Table B: Where Costs Are (Cost Clusters by Share of Total Classified Costs)
    cb_hdr = 16
    ws.cell(row=cb_hdr, column=1, value="Cost Cluster (Where Classified Costs Go)")
    ws.cell(row=cb_hdr, column=2, value=f"Cumulative Spend ({cur})")
    ws.cell(row=cb_hdr, column=3, value="% of Classified Costs")
    ws.cell(row=cb_hdr, column=4, value="Cost Classification Type")
    style_header_row(ws, cb_hdr, 4)

    cluster_defs = [
        (direct_label, f"='Summary P&L'!{sp_cum_col}{sp_rows['dc']}", "Direct Service / Goods Costs"),
        *[(b, f"='Summary P&L'!{sp_cum_col}{sp_rows['opex_' + b]}", "Operating Overhead") for b in OPEX_BUCKETS],
        ("Taxes & Statutory Levies", f"='Summary P&L'!{sp_cum_col}{sp_rows['tax']}", "Statutory Taxes & Levies"),
        ("Bank & Payment Fees", f"='Summary P&L'!{sp_cum_col}{sp_rows['fee']}", "Banking & Payment Friction"),
    ]
    for idx, (c_name, c_form, c_type) in enumerate(cluster_defs, cb_hdr + 1):
        ws.cell(row=idx, column=1, value=c_name).font = FONT_BODY
        amt_c = ws.cell(row=idx, column=2, value=c_form)
        amt_c.number_format = MONEY_FMT
        amt_c.font = FONT_BODY
        pct_c = ws.cell(row=idx, column=3, value=f"=IF($B$11>0,B{idx}/$B$11,0)")
        pct_c.number_format = PCT_FMT
        pct_c.font = FONT_BOLD
        ws.cell(row=idx, column=4, value=c_type).font = FONT_BODY
        for c_i in range(1, 5):
            ws.cell(row=idx, column=c_i).border = THIN_BORDER
        ws.row_dimensions[idx].height = 23

    tot_cl_row = cb_hdr + 1 + len(cluster_defs)
    ws.cell(row=tot_cl_row, column=1, value="Total Classified P&L Costs").font = FONT_BOLD
    ws.cell(row=tot_cl_row, column=2, value=f"=SUM(B{cb_hdr + 1}:B{tot_cl_row - 1})").number_format = MONEY_FMT
    ws.cell(row=tot_cl_row, column=2).font = FONT_BOLD
    ws.cell(row=tot_cl_row, column=3, value=f"=SUM(C{cb_hdr + 1}:C{tot_cl_row - 1})").number_format = PCT_FMT
    ws.cell(row=tot_cl_row, column=3).font = FONT_BOLD
    for c_i in range(1, 5):
        ws.cell(row=tot_cl_row, column=c_i).fill = FILL_SUBTOTAL
        ws.cell(row=tot_cl_row, column=c_i).border = SUBTOTAL_BORDER

    # Table C: Top 10 Individual Cost Categories
    tc_hdr = tot_cl_row + 3
    ws.cell(row=tc_hdr, column=1, value="Top Individual Cost Categories")
    ws.cell(row=tc_hdr, column=2, value="Section")
    ws.cell(row=tc_hdr, column=3, value=f"Spend ({cur})")
    ws.cell(row=tc_hdr, column=4, value="% of Total Costs")
    ws.cell(row=tc_hdr, column=5, value="Transactions")
    style_header_row(ws, tc_hdr, 5)

    for idx, cat_item in enumerate(computed["health"]["top_cost_categories"], tc_hdr + 1):
        ws.cell(row=idx, column=1, value=cat_item["category"]).font = FONT_BODY
        ws.cell(row=idx, column=2, value=cat_item["section"]).font = FONT_BODY
        sc = ws.cell(row=idx, column=3, value=f'=-SUMIFS({lr("I")},{lr("M")},B{idx},{lr("N")},A{idx})')
        sc.number_format = MONEY_FMT
        sc.font = FONT_BODY
        pc = ws.cell(row=idx, column=4, value=f"=IF($B$11>0,C{idx}/$B$11,0)")
        pc.number_format = PCT_FMT
        pc.font = FONT_BOLD
        cnt = ws.cell(row=idx, column=5, value=f'=COUNTIFS({lr("M")},B{idx},{lr("N")},A{idx})')
        cnt.number_format = "0"
        cnt.font = FONT_BODY
        for c_i in range(1, 6):
            ws.cell(row=idx, column=c_i).border = THIN_BORDER
        ws.row_dimensions[idx].height = 22

    set_col_widths(ws, {1: 42, 2: 24, 3: 24, 4: 34, 5: 16})

    # -------------------------------------------------------------------------
    # 7. NEEDS REVIEW
    # -------------------------------------------------------------------------
    ws = ws_map["Needs Review"]
    setup_sheet(
        ws,
        "Needs Review Register (Sorted by Impact)",
        "Largest open items appear first. Record owner response, update Master Ledger Section/Category, and mark Resolved.",
        tab_colors["Needs Review"],
    )
    nr_headers = [
        "Review ID", "Level", "Ledger Txn ID", "Date", "Account", f"Signed ({cur})",
        "Issue", "Question / Evidence Needed", "Original Description / Finding",
        "Source File", "Page(s)", "Current Category", "Status", "Owner Response", "Resolution",
    ]
    for c_i, h_t in enumerate(nr_headers, 1):
        ws.cell(row=5, column=c_i, value=h_t)
    style_header_row(ws, 5, len(nr_headers))

    nr_row = 6
    for i_idx, iss in enumerate(scope_issues, 1):
        vals = [f"G{i_idx:02d}", "Scope / completeness", "", None, "", None, iss[0], iss[2], iss[1], "", "", "", "Open", "", ""]
        for c_i, v in enumerate(vals, 1):
            c = ws.cell(row=nr_row, column=c_i, value=v)
            c.font = FONT_BODY
            c.border = THIN_BORDER
            c.alignment = Alignment(vertical="center", wrap_text=(c_i in (8, 9, 14, 15)))
            if c_i in (13, 14, 15):
                c.fill = FILL_AMBER
        ws.row_dimensions[nr_row].height = 46
        nr_row += 1

    for r in review_items:
        vals = [
            r["id"], "Transaction", r["ledger_id"], parse_iso_date(r["date"]), r["account"],
            f'=INDEX({lr("I")},MATCH(C{nr_row},{lr("A")},0))', r["issue"], r["question"],
            r["description"], r["file"], r["pages"],
            f'=INDEX({lr("N")},MATCH(C{nr_row},{lr("A")},0))',
            r.get("status", "Open"), r.get("response", ""), r.get("resolution", ""),
        ]
        for c_i, v in enumerate(vals, 1):
            c = ws.cell(row=nr_row, column=c_i, value=v)
            c.font = FONT_BODY
            c.border = THIN_BORDER
            c.alignment = Alignment(vertical="center", wrap_text=(c_i in (8, 9, 14, 15)))
            if c_i == 4:
                c.number_format = DATE_FMT
            elif c_i == 6:
                c.number_format = MONEY_FMT
            if c_i in (13, 14, 15):
                c.fill = FILL_AMBER
        ws.row_dimensions[nr_row].height = 46
        nr_row += 1

    dv_stat = DataValidation(type="list", formula1='"Open,Resolved"', allow_blank=False)
    ws.add_data_validation(dv_stat)
    if NR_END >= 6:
        dv_stat.add(f"M6:M{NR_END}")
    ws.freeze_panes = "D6"
    set_col_widths(ws, {1: 12, 2: 20, 3: 16, 4: 14, 5: 24, 6: 16, 7: 22, 8: 52, 9: 50, 10: 44, 11: 10, 12: 28, 13: 14, 14: 36, 15: 36})

    # -------------------------------------------------------------------------
    # 1. EXECUTIVE OVERVIEW (FRONT LANDING SHEET)
    # -------------------------------------------------------------------------
    ws = ws_map["Executive Overview"]
    setup_sheet(
        ws,
        f"{p['business_name']} | Executive Financial Overview",
        f"{cur} • {p['reporting_basis']} • Start here for headline results, cash bridge, and month-by-month performance.",
        tab_colors["Executive Overview"],
    )
    # Block 1: Headline Business Summary & Cash Bridge
    ws["A5"], ws["B5"], ws["C5"] = "Headline Business Summary", f"Cumulative ({cur})", "What This Means"
    style_header_row(ws, 5, 3)
    ov_summary = [
        ("Money Made (Confirmed Revenue)", f"='Summary P&L'!{sp_cum_col}{sp_rows['rev']}", "Operating revenue confirmed from bank narrations / owner tags", False),
        (direct_label, f"=-'Summary P&L'!{sp_cum_col}{sp_rows['dc']}", "Direct costs of delivering services or goods", False),
        ("Gross Profit / (Loss)", f"='Summary P&L'!{sp_cum_col}{sp_rows['gp']}", "Margin after direct service/product costs", True),
        ("Operating Expenses & Taxes/Fees", f"=-('Summary P&L'!{sp_cum_col}{sp_rows['tot_opex']}+'Summary P&L'!{sp_cum_col}{sp_rows['tax_fee']})", "Staff, power/fuel, marketing, premises, admin, taxes & bank fees", False),
        ("Provisional Net Profit / (Loss)", f"='Summary P&L'!{sp_cum_col}{sp_rows['net_profit']}", "Confirmed P&L bottom line (excludes unclassified cash below)", True),
        ("Unclassified Receipts Pending Label", f"='Summary P&L'!{sp_cum_col}{sp_rows['unclass_in']}", "Bank inflows awaiting owner confirmation in Needs Review", "amber"),
        ("Unclassified Payments Pending Label", f"='Summary P&L'!{sp_cum_col}{sp_rows['unclass_out']}", "Bank outflows awaiting owner confirmation in Needs Review", "amber"),
        ("Non-P&L Cash (Loans, Assets, Other Accts)", f"='Summary P&L'!{sp_cum_col}{sp_rows['non_pnl']}", "Cash movements outside P&L (loans, equipment, transfers)", False),
        ("Actual Net Bank Cash Change", f"='Summary P&L'!{sp_cum_col}{sp_rows['cash_change']}", "Exact net cash change across supplied bank statements", True),
        ("Open Review Questions", f'=COUNTIFS(\'Needs Review\'!$M$6:$M${NR_END},"Open")', "Remaining open items in Needs Review", "count"),
    ]
    for idx, (lbl, form, desc_txt, st) in enumerate(ov_summary, 6):
        is_b = st is True
        ws.cell(row=idx, column=1, value=lbl).font = FONT_BOLD if is_b else FONT_BODY
        vc = ws.cell(row=idx, column=2, value=form)
        vc.font = FONT_BOLD if is_b else FONT_BODY
        vc.number_format = "0" if st == "count" else MONEY_FMT
        ws.cell(row=idx, column=3, value=desc_txt).font = FONT_BODY
        for c_i in range(1, 4):
            c = ws.cell(row=idx, column=c_i)
            if is_b:
                c.fill = FILL_HIGHLIGHT if "Net Profit" in lbl else FILL_SUBTOTAL
                c.border = SUBTOTAL_BORDER
            elif st == "amber":
                c.fill = FILL_AMBER
                c.border = THIN_BORDER
            else:
                c.border = THIN_BORDER
        ws.row_dimensions[idx].height = 24

    # Block 2: Bank Account Balances
    ab_hdr = 18
    ws.cell(row=ab_hdr, column=1, value="Supplied Bank Accounts")
    ws.cell(row=ab_hdr, column=2, value=f"Opening ({cur})")
    ws.cell(row=ab_hdr, column=3, value=f"Latest Closing ({cur})")
    style_header_row(ws, ab_hdr, 3)

    num_accts = len(computed["accounts"])
    for a_i, acct_name in enumerate(computed["accounts"]):
        rr = ab_hdr + 1 + a_i
        first_ac_row = 6 + a_i
        last_ac_row = AC_END - (num_accts - 1 - a_i)
        ws.cell(row=rr, column=1, value=acct_name).font = FONT_BODY
        c1 = ws.cell(row=rr, column=2, value=f"='Account Cash'!F{first_ac_row}")
        c2 = ws.cell(row=rr, column=3, value=f"='Account Cash'!J{last_ac_row}")
        c1.number_format = c2.number_format = MONEY_FMT
        c1.font = c2.font = FONT_BODY
        ws.row_dimensions[rr].height = 23

    tot_ab_row = ab_hdr + 1 + num_accts
    ws.cell(row=tot_ab_row, column=1, value="Combined Supplied Bank Cash").font = FONT_BOLD
    ws.cell(row=tot_ab_row, column=2, value=f"=SUM(B{ab_hdr + 1}:B{tot_ab_row - 1})").number_format = MONEY_FMT
    ws.cell(row=tot_ab_row, column=3, value=f"=SUM(C{ab_hdr + 1}:C{tot_ab_row - 1})").number_format = MONEY_FMT
    for c_i in range(1, 4):
        ws.cell(row=tot_ab_row, column=c_i).font = FONT_BOLD
        ws.cell(row=tot_ab_row, column=c_i).fill = FILL_SUBTOTAL
        ws.cell(row=tot_ab_row, column=c_i).border = SUBTOTAL_BORDER
    ws.row_dimensions[tot_ab_row].height = 24

    # Block 3: Month-by-Month Headline Comparison
    mc_hdr = tot_ab_row + 3
    ov_last_col = M_COUNT + 2
    ws.cell(row=mc_hdr, column=1, value="Month-by-Month Headline Comparison")
    for m_i, md in enumerate(month_dates, 2):
        c = ws.cell(row=mc_hdr, column=m_i, value=md)
        c.number_format = 'mmm-yy" (partial)"' if pflags.get(months[m_i - 2]) else MONTH_FMT
    ws.cell(row=mc_hdr, column=ov_last_col, value="Cumulative")
    style_header_row(ws, mc_hdr, ov_last_col)

    mc_lines = [
        ("Confirmed Revenue", "Summary P&L", sp_rows["rev"]),
        (direct_label, "Summary P&L", sp_rows["dc"]),
        ("Gross Profit / (Loss)", "Summary P&L", sp_rows["gp"]),
        ("Total Operating Expenses", "Summary P&L", sp_rows["tot_opex"]),
        ("Provisional Net Profit / (Loss)", "Summary P&L", sp_rows["net_profit"]),
        ("Unclassified Receipts", "Summary P&L", sp_rows["unclass_in"]),
        ("Unclassified Payments", "Summary P&L", sp_rows["unclass_out"]),
        ("Closing Bank Cash", "Cash Flow", 23),
    ]
    for offset, (lbl, src_sheet, src_row) in enumerate(mc_lines, 1):
        rr = mc_hdr + offset
        is_b = "Profit" in lbl or "Closing" in lbl
        ws.cell(row=rr, column=1, value=lbl).font = FONT_BOLD if is_b else FONT_BODY
        for c_i in range(2, ov_last_col + 1):
            cl = get_column_letter(c_i)
            cell = ws.cell(row=rr, column=c_i, value=f"='{src_sheet}'!{cl}{src_row}")
            cell.number_format = MONEY_FMT
            cell.font = FONT_BOLD if is_b else FONT_BODY
            if is_b:
                cell.fill = FILL_SUBTOTAL
        ws.row_dimensions[rr].height = 23

    set_col_widths(ws, {1: 42, 2: 20, 3: 48}, default_width=18, max_col=ov_last_col)

    # -------------------------------------------------------------------------
    # 9–14. AUDIT APPENDIX TABS
    # -------------------------------------------------------------------------
    # 9. Transfer Matches
    ws = ws_map["Transfer Matches"]
    setup_sheet(ws, "Matched Internal Transfers", "Cross-account principal pairs net to zero in external cash and P&L.", tab_colors["Transfer Matches"])
    tm_headers = ["Transfer ID", "Date", "Debit Bank ID", "Credit Bank ID", f"Principal ({cur})", f"Ledger Debit ({cur})", f"Ledger Credit ({cur})", "Net Pair Diff", "Match Basis", "Debit Ref", "Credit Ref"]
    for c_i, h_t in enumerate(tm_headers, 1):
        ws.cell(row=5, column=c_i, value=h_t)
    style_header_row(ws, 5, len(tm_headers))
    for idx, t in enumerate(transfers, 6):
        d_ref = t.get("debit_reference") or t.get("p_reference", "")
        c_ref = t.get("credit_reference") or t.get("m_reference", "")
        vals = [
            t["id"], parse_iso_date(t["date"]), t["debit_id"], t["credit_id"], t["amount"],
            f'=SUMIFS({lr("K")},{lr("P")},A{idx})', f'=SUMIFS({lr("J")},{lr("P")},A{idx})',
            f"=G{idx}-F{idx}", t["basis"], d_ref, c_ref,
        ]
        for c_i, v in enumerate(vals, 1):
            c = ws.cell(row=idx, column=c_i, value=v)
            c.font = FONT_BODY
            c.border = THIN_BORDER
            if c_i == 2:
                c.number_format = DATE_FMT
            elif c_i in (5, 6, 7, 8):
                c.number_format = DIFF_FMT if c_i == 8 else MONEY_FMT
    if transfers:
        add_nonzero_red_rule(ws, f"H6:H{len(transfers) + 5}")
    set_col_widths(ws, {1: 14, 2: 14, 3: 15, 4: 15, 5: 18, 6: 18, 7: 18, 8: 16, 9: 48, 10: 34, 11: 34})

    # 10. Statement Inventory
    ws = ws_map["Statement Inventory"]
    setup_sheet(ws, "Statement Inventory", f"{len(docs)} source files inventoried with SHA-256 hashes and coverage checks.", tab_colors["Statement Inventory"])
    si_headers = ["Source ID", "Account", "Start Date", "End Date", f"Opening ({cur})", f"Credits ({cur})", f"Debits ({cur})", f"Closing ({cur})", "Pages", "Source Type", "Ledger Use", "Coverage / Exclusion Note", "Filename", "SHA-256"]
    for c_i, h_t in enumerate(si_headers, 1):
        ws.cell(row=5, column=c_i, value=h_t)
    style_header_row(ws, 5, len(si_headers))
    inv_row_by_id = {}
    for idx, d in enumerate(docs, 6):
        inv_row_by_id[d["id"]] = idx
        backup = bool(d.get("error"))
        is_cons = d.get("is_consolidated", False) or d.get("id") == "S17"
        vals = [
            d["id"], d.get("account", ""), parse_iso_date(d.get("start")), parse_iso_date(d.get("end")),
            d.get("opening"), d.get("credits"), d.get("debits"), d.get("closing"), d.get("page_count"),
            "Encrypted backup" if backup else ("Consolidated overlap" if is_cons else "Monthly statement"),
            "Excluded" if (backup or is_cons) else "Included",
            d.get("error") or ("Consolidated duplicate overlap" if is_cons else "Primary monthly statement"),
            d.get("file", ""), d.get("sha256", ""),
        ]
        for c_i, v in enumerate(vals, 1):
            c = ws.cell(row=idx, column=c_i, value=v)
            c.font = FONT_BODY
            c.border = THIN_BORDER
            if c_i in (3, 4) and v:
                c.number_format = DATE_FMT
            elif c_i in (5, 6, 7, 8) and v is not None:
                c.number_format = MONEY_FMT
    set_col_widths(ws, {1: 12, 2: 24, 3: 14, 4: 14, 5: 16, 6: 16, 7: 16, 8: 16, 9: 8, 10: 20, 11: 14, 12: 38, 13: 54, 14: 36})

    # 11. Source Rows
    ws = ws_map["Source Rows"]
    setup_sheet(ws, "Extracted Bank Statement Rows", "Raw statement settlements and running-balance chain audit.", tab_colors["Source Rows"])
    sr_headers = ["Source Row ID", "Statement ID", "Printed Seq", "Balance Seq", "Bank Txn ID", "Included (1/0)", "Date", f"Debit ({cur})", f"Credit ({cur})", "Source Before", "Chain Before", "Source After", "Running Diff", "Source Before Diff", "Seq Diff", "Reference", "Description", "Filename", "Page(s)"]
    for c_i, h_t in enumerate(sr_headers, 1):
        ws.cell(row=5, column=c_i, value=h_t)
    style_header_row(ws, 5, len(sr_headers))
    for idx, r in enumerate(all_src, 6):
        vals = [
            f"{r['statement']}-{r['seq']}", r["statement"], r["seq"], r.get("balance_order", r["seq"]),
            r.get("raw_id", ""), r.get("included", 1), parse_iso_date(r["date"]),
            r["debit"], r["credit"], r.get("before"), r.get("chain_before"), r["balance"],
            f"=ROUND(K{idx}+I{idx}-H{idx}-L{idx},2)",
            f'=IF(ISNUMBER(J{idx}),ROUND(J{idx}+I{idx}-H{idx}-L{idx},2),"n.a.")',
            f"=D{idx}-C{idx}", r.get("reference", ""), r.get("description", ""), r.get("file", ""), r.get("pages", ""),
        ]
        for c_i, v in enumerate(vals, 1):
            c = ws.cell(row=idx, column=c_i, value=v)
            c.font = FONT_BODY
            c.border = THIN_BORDER
            if c_i == 7:
                c.number_format = DATE_FMT
            elif c_i in (8, 9, 10, 11, 12, 13):
                c.number_format = DIFF_FMT if c_i == 13 else MONEY_FMT
    if all_src:
        add_nonzero_red_rule(ws, f"M6:M{R_END}")
    ws.freeze_panes = "C6"
    set_col_widths(ws, {1: 15, 2: 13, 3: 12, 4: 12, 5: 14, 6: 13, 7: 14, 8: 16, 9: 16, 10: 16, 11: 16, 12: 16, 13: 14, 14: 16, 15: 10, 16: 34, 17: 54, 18: 48, 19: 10})

    # 12. Duplicate Exclusions
    ws = ws_map["Duplicate Exclusions"]
    setup_sheet(ws, "Duplicate Exclusions", "Overlapping consolidated rows excluded in favor of primary monthly statements.", tab_colors["Duplicate Exclusions"])
    de_headers = ["Excluded Statement", "Excluded Seq", "Excluded Page", "Retained Statement", "Retained Seq", "Matching Basis"]
    for c_i, h_t in enumerate(de_headers, 1):
        ws.cell(row=5, column=c_i, value=h_t)
    style_header_row(ws, 5, len(de_headers))
    for idx, d in enumerate(duplicates, 6):
        for c_i, v in enumerate([d["excluded_statement"], d["excluded_seq"], d["excluded_pages"], d["kept_statement"], d["kept_seq"], d["basis"]], 1):
            ws.cell(row=idx, column=c_i, value=v).font = FONT_BODY
    set_col_widths(ws, {1: 18, 2: 14, 3: 14, 4: 18, 5: 14, 6: 68})

    # 13. Reconciliations
    ws = ws_map["Reconciliations"]
    setup_sheet(ws, "Statement & Ledger Reconciliations", "All difference cells must equal 0.00. Verifies source statements, ledger splits, Cash Flow, Detailed P&L, and Summary P&L.", tab_colors["Reconciliations"])
    rc_headers = ["Source ID", "Account", "Opening", "Reported Credits", "Reported Debits", "Reported Closing", "Statement Eq Diff", "Extracted Credits", "Extracted Debits", "Credit Diff", "Debit Diff", "Closing Diff", "Running Breaks", "Reordered Rows"]
    for c_i, h_t in enumerate(rc_headers, 1):
        ws.cell(row=5, column=c_i, value=h_t)
    style_header_row(ws, 5, len(rc_headers))
    for idx, d in enumerate(valid_docs, 6):
        vals = [
            d["id"], d["account"], d["opening"], d["credits"], d["debits"], d["closing"],
            f"=ROUND(C{idx}+D{idx}-E{idx}-F{idx},2)",
            f'=SUMIFS({sr("I")},{sr("B")},A{idx})', f'=SUMIFS({sr("H")},{sr("B")},A{idx})',
            f"=ROUND(H{idx}-D{idx},2)", f"=ROUND(I{idx}-E{idx},2)", f"=ROUND(C{idx}+H{idx}-I{idx}-F{idx},2)",
            f'=COUNTIFS({sr("B")},A{idx},{sr("M")},">0.005")+COUNTIFS({sr("B")},A{idx},{sr("M")},"<-0.005")',
            d.get("reordered", 0),
        ]
        for c_i, v in enumerate(vals, 1):
            c = ws.cell(row=idx, column=c_i, value=v)
            c.font = FONT_BODY
            if c_i in (3, 4, 5, 6, 8, 9):
                c.number_format = MONEY_FMT
            elif c_i in (7, 10, 11, 12):
                c.number_format = DIFF_FMT
    rc_end = len(valid_docs) + 5
    for col_l in ("G", "J", "K", "L", "M"):
        add_nonzero_red_rule(ws, f"{col_l}6:{col_l}{rc_end}")

    audit_hdr = rc_end + 3
    ws.cell(row=audit_hdr, column=1, value="Cross-Workbook Control Check")
    ws.cell(row=audit_hdr, column=2, value=f"Difference ({cur})")
    ws.cell(row=audit_hdr, column=3, value="Control Meaning")
    style_header_row(ws, audit_hdr, 3)

    tm_end = max(len(transfers) + 5, 6)
    audits = [
        ("Allocation net vs Source Rows", f'=ROUND(SUM({lr("I")})-SUMIFS({sr("I")},{sr("F")},1)+SUMIFS({sr("H")},{sr("F")},1),2)', "Master Ledger net equals included bank statement rows"),
        ("External cash vs All ledger cash", f'=ROUND(SUM({lr("T")})-SUM({lr("I")}),2)', "Internal transfer pairs eliminate to zero"),
        ("Cash Flow closing vs Account Cash", f"=ROUND('Cash Flow'!{cf_cum_col}23-'Executive Overview'!C{tot_ab_row},2)", "Cash Flow ending balance matches Account Cash closings"),
        ("Detailed P&L vs Master Ledger P&L impact", f"=ROUND('Detailed P&L'!{dp_cum_col}{dp_profit_row}-SUM({lr('S')}),2)", "Detailed P&L net profit ties to Master Ledger"),
        ("Summary P&L vs Detailed P&L net profit", f"=ROUND('Summary P&L'!{sp_cum_col}{sp_rows['net_profit']}-'Detailed P&L'!{dp_cum_col}{dp_profit_row},2)", "Short P&L ties identically to Detailed P&L"),
        ("Summary P&L Cash Bridge vs Cash Flow net", f"=ROUND('Summary P&L'!{sp_cum_col}{sp_rows['cash_change']}-'Cash Flow'!{cf_cum_col}21,2)", "Short P&L Cash Bridge ties to external bank cash change"),
        ("Matched transfers net", f"=ROUND(SUM('Transfer Matches'!H6:H{tm_end}),2)" if transfers else "=0", "Every matched transfer pair offsets to zero"),
    ]
    for k_i, (a_lbl, a_form, a_desc) in enumerate(audits, 1):
        rr = audit_hdr + k_i
        ws.cell(row=rr, column=1, value=a_lbl).font = FONT_BOLD
        bc = ws.cell(row=rr, column=2, value=a_form)
        bc.font = FONT_BOLD
        bc.number_format = DIFF_FMT
        ws.cell(row=rr, column=3, value=a_desc).font = FONT_BODY
    add_nonzero_red_rule(ws, f"B{audit_hdr + 1}:B{audit_hdr + len(audits)}")
    set_col_widths(ws, {1: 38, 2: 24, 3: 52}, default_width=18, max_col=14)

    # 14. ReadMe
    ws = ws_map["ReadMe"]
    setup_sheet(ws, "Basis, Guide & Accounting Policies", "How to read these books, update classifications, and verify reconciliations.", tab_colors["ReadMe"])
    ws["A5"], ws["B5"] = "Topic", "Explanation & Policy"
    style_header_row(ws, 5, 2)
    readme_notes = [
        ("How to read this workbook", "Start with 'Executive Overview', 'Summary P&L', and 'Financial Health' for a concise small-business view. Accounting detail and statement-level audit trails follow in tabs 4–13."),
        ("Accounting basis", f"{p['reporting_basis']} in {cur}. Transactions are recognized when cash settles on supplied bank statements. No accruals, inventory counts, or depreciation schedules are inferred without supporting records."),
        ("Conservative uncertainty policy", "Ambiguous inflows and outflows are placed in 'Needs Review' rather than guessed into Revenue or Expenses. Their cash impact is reconciled transparently in the Summary P&L Cash Bridge."),
        ("Updating classifications", "Either label unknown transactions in the companion Interactive HTML Dashboard and export classifications.json, or edit Section/Category directly in 'Master Ledger' (yellow columns M–N). All P&L and Health formulas update automatically."),
        ("Internal transfers & embedded fees", "Same-day transfers between supplied business accounts are paired and excluded from P&L and external cash flow. Embedded processing fees are separated into Bank Fees so principal and fee sum to exact bank settlements."),
        ("Non-advisory Financial Health policy", "The 'Financial Health' sheet reports strictly factual liquidity, cost concentration, and completeness ratios. It intentionally excludes unsolicited business advice or recommendations."),
    ]
    for idx, (t_lbl, t_body) in enumerate(readme_notes, 6):
        ws.cell(row=idx, column=1, value=t_lbl).font = FONT_BOLD
        bc = ws.cell(row=idx, column=2, value=t_body)
        bc.font = FONT_BODY
        bc.alignment = Alignment(vertical="center", wrap_text=True)
        ws.row_dimensions[idx].height = 46
    set_col_widths(ws, {1: 34, 2: 110})

    return wb


def write_csv_ledger(computed: dict[str, Any], csv_path: Path) -> None:
    cols = [
        "id", "raw_id", "date", "datetime", "description", "amount", "direction",
        "account", "section", "summary_group", "category", "component",
        "transfer_id", "review_type", "question", "reference", "statement", "file", "pages",
    ]
    csv_path.parent.mkdir(parents=True, exist_ok=True)
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        writer.writeheader()
        for e in computed["ledger"]:
            writer.writerow({**e, "direction": "Credit" if e["amount"] > 0 else "Debit"})


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate portable openpyxl Cash Books & CSV ledger.")
    parser.add_argument("--books", required=True, type=Path, help="Path to books_data.json.")
    parser.add_argument("--classifications", type=Path, default=None, help="Optional classifications.json path.")
    parser.add_argument("--parsed", type=Path, default=None, help="Optional parsed.json for full source rows.")
    parser.add_argument("--out-xlsx", required=True, type=Path, help="Output .xlsx path.")
    parser.add_argument("--out-csv", type=Path, default=None, help="Optional output .csv ledger path.")
    args = parser.parse_args()

    computed = load_and_compute(args.books, args.classifications, args.parsed)
    wb = build_workbook(computed)
    args.out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.out_xlsx)
    print(f"Wrote Excel Workbook -> {args.out_xlsx} ({len(wb.sheetnames)} sheets)")

    if args.out_csv:
        write_csv_ledger(computed, args.out_csv)
        print(f"Wrote CSV Ledger -> {args.out_csv} ({len(computed['ledger'])} rows)")


if __name__ == "__main__":
    main()
