#!/usr/bin/env python3
"""Comprehensive Verification-Driven Coding (VDC) suite for the Small-Business Financial Insights package.

Checks:
1. Independent Primary-Source Tie-Out: computes raw bank settlements directly from parsed_statements.json
   (or raw + duplicates) without relying solely on ledger_engine.py.
2. Exact arithmetic & cash-bridge tie-out across ledger, Summary P&L, Detailed P&L, and Cash Flow.
3. Excel (.xlsx) 14-sheet priority structure, formula syntax, AND deterministic XML Formula Evaluator
   checking Reconciliations, Account Cash, Transfer Matches, Summary P&L, Detailed P&L, Cash Flow,
   and Executive Overview against independent expectations.
4. Standalone HTML dashboard JSON payload integrity, CSV tie-out, and round-trip override invariance.
5. Strict Non-Advisory Guardrail scan on the Financial Health Report (zero suggestions/advice).
6. Enforced Verdict Cap (VERIFIED vs PROVISIONAL_REVIEW_REQUIRED) and Explicit Unverified Range (U).
"""

from __future__ import annotations

import argparse
import copy
import csv
from datetime import date, timedelta
from decimal import Decimal, ROUND_HALF_UP
import json
from pathlib import Path
import re
import tempfile
from typing import Any
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from _contract import BALANCE_TOLERANCE, UNVERIFIED_PROPERTIES, make_delta, make_vdc_report
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

EXCEL_EPOCH = date(1899, 12, 30)


# ---------------------------------------------------------------------------
# Lightweight Deterministic Excel XML Formula Evaluator
# ---------------------------------------------------------------------------

def _col_to_idx(col: str) -> int:
    idx = 0
    for ch in col.upper():
        idx = idx * 26 + (ord(ch) - 64)
    return idx


def _idx_to_col(idx: int) -> str:
    res = []
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        res.append(chr(65 + rem))
    return "".join(reversed(res))


def _split_coord(coord: str) -> tuple[int, int]:
    m = re.fullmatch(r"\$?([A-Za-z]+)\$?(\d+)", coord.strip())
    if not m:
        raise ValueError(f"Invalid cell coordinate: {coord}")
    return _col_to_idx(m.group(1)), int(m.group(2))


def _is_num(val: Any) -> bool:
    return isinstance(val, (int, float)) and not isinstance(val, bool)


def _to_num(val: Any) -> float:
    if val is None or val == "":
        return 0.0
    if isinstance(val, bool):
        return 1.0 if val else 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _excel_round(val: float, digits: int) -> float:
    d = Decimal(str(val))
    q = Decimal("1") if digits == 0 else Decimal(f"1e-{digits}")
    return float(d.quantize(q, rounding=ROUND_HALF_UP))


def _matches_criterion(cell_val: Any, crit: Any) -> bool:
    if isinstance(crit, str):
        for op in ("<>", ">=", "<=", ">", "<", "="):
            if crit.startswith(op):
                rhs_raw = crit[len(op):]
                if op in (">", "<", ">=", "<="):
                    if not _is_num(cell_val):
                        return False
                    rhs = float(rhs_raw)
                    lhs = float(cell_val)
                    return (
                        lhs > rhs if op == ">"
                        else lhs < rhs if op == "<"
                        else lhs >= rhs if op == ">="
                        else lhs <= rhs
                    )
                lhs_s = "" if cell_val is None else str(cell_val)
                if op == "<>":
                    if _is_num(cell_val):
                        try:
                            return abs(float(cell_val) - float(rhs_raw)) >= 1e-9
                        except ValueError:
                            pass
                    return lhs_s.lower() != rhs_raw.lower()
                if op == "=":
                    if _is_num(cell_val):
                        try:
                            return abs(float(cell_val) - float(rhs_raw)) < 1e-9
                        except ValueError:
                            pass
                    return lhs_s.lower() == rhs_raw.lower()
        if _is_num(cell_val):
            try:
                return abs(float(cell_val) - float(crit)) < 1e-9
            except ValueError:
                return False
        return ("" if cell_val is None else str(cell_val)).lower() == crit.lower()

    if _is_num(crit):
        if not _is_num(cell_val):
            return False
        return abs(float(cell_val) - float(crit)) < 1e-9

    return cell_val == crit


class ExcelFormulaEvaluator:
    """Evaluates openpyxl-generated spreadsheet formulas directly from OOXML."""

    def __init__(self, sheets_data: dict[str, dict[str, tuple[Any, str | None]]]) -> None:
        self.sheets = sheets_data
        self.cache: dict[tuple[str, str], Any] = {}
        self.visiting: set[tuple[str, str]] = set()

    def get_cell(self, sheet: str, coord: str) -> Any:
        coord = coord.replace("$", "").upper()
        key = (sheet, coord)
        if key in self.cache:
            return self.cache[key]
        raw_val, formula = self.sheets.get(sheet, {}).get(coord, (None, None))
        if formula is None:
            self.cache[key] = raw_val
            return raw_val
        if key in self.visiting:
            raise ValueError(f"Circular reference at {sheet}!{coord}")
        self.visiting.add(key)
        try:
            val = self._eval_expr(formula, sheet)
        finally:
            self.visiting.remove(key)
        self.cache[key] = val
        return val

    def get_range(self, sheet: str, start_coord: str, end_coord: str) -> list[Any]:
        c1, r1 = _split_coord(start_coord)
        c2, r2 = _split_coord(end_coord)
        out = []
        for r in range(min(r1, r2), max(r1, r2) + 1):
            for c in range(min(c1, c2), max(c1, c2) + 1):
                out.append(self.get_cell(sheet, f"{_idx_to_col(c)}{r}"))
        return out

    def _tokenize(self, expr: str) -> list[tuple[str, str]]:
        tokens: list[tuple[str, str]] = []
        i = 0
        n = len(expr)
        while i < n:
            ch = expr[i]
            if ch.isspace():
                i += 1
                continue
            if ch == '"':
                j = i + 1
                buf = []
                while j < n:
                    if expr[j] == '"' and j + 1 < n and expr[j + 1] == '"':
                        buf.append('"')
                        j += 2
                    elif expr[j] == '"':
                        break
                    else:
                        buf.append(expr[j])
                        j += 1
                tokens.append(("STR", "".join(buf)))
                i = j + 1
                continue
            if expr[i : i + 2] in ("<>", ">=", "<="):
                tokens.append(("OP", expr[i : i + 2]))
                i += 2
                continue
            if ch in "+-*/=<>(),:":
                tokens.append(("SYM" if ch in "(),:" else "OP", ch))
                i += 1
                continue
            if ch == "'":
                j = expr.find("'", i + 1)
                if j == -1:
                    raise ValueError(f"Unterminated sheet quote in formula: {expr}")
                sheet_name = expr[i + 1 : j]
                if j + 1 < n and expr[j + 1] == "!":
                    k = j + 2
                    while k < n and (expr[k].isalnum() or expr[k] in "$"):
                        k += 1
                    cell_part = expr[j + 2 : k]
                    if k < n and expr[k] == ":":
                        k2 = k + 1
                        while k2 < n and (expr[k2].isalnum() or expr[k2] in "$"):
                            k2 += 1
                        end_part = expr[k + 1 : k2]
                        tokens.append(("RANGE", f"{sheet_name}!{cell_part}:{end_part}"))
                        i = k2
                    else:
                        tokens.append(("REF", f"{sheet_name}!{cell_part}"))
                        i = k
                    continue
            if ch.isdigit() or (ch == "." and i + 1 < n and expr[i + 1].isdigit()):
                j = i
                while j < n and (expr[j].isdigit() or expr[j] == "."):
                    j += 1
                tokens.append(("NUM", expr[i:j]))
                i = j
                continue
            if ch.isalpha() or ch == "$":
                j = i
                while j < n and (expr[j].isalnum() or expr[j] in "$_"):
                    j += 1
                ident = expr[i:j]
                if j < n and expr[j] == "!":
                    k = j + 1
                    while k < n and (expr[k].isalnum() or expr[k] in "$"):
                        k += 1
                    cell_part = expr[j + 1 : k]
                    if k < n and expr[k] == ":":
                        k2 = k + 1
                        while k2 < n and (expr[k2].isalnum() or expr[k2] in "$"):
                            k2 += 1
                        tokens.append(("RANGE", f"{ident}!{cell_part}:{expr[k + 1 : k2]}"))
                        i = k2
                    else:
                        tokens.append(("REF", f"{ident}!{cell_part}"))
                        i = k
                    continue
                if j < n and expr[j] == "(":
                    tokens.append(("FUNC", ident.upper()))
                    i = j
                    continue
                if j < n and expr[j] == ":" and re.fullmatch(r"\$?[A-Za-z]+\$?\d+", ident):
                    k = j + 1
                    while k < n and (expr[k].isalnum() or expr[k] in "$"):
                        k += 1
                    tokens.append(("RANGE", f"{ident}:{expr[j + 1 : k]}"))
                    i = k
                    continue
                if re.fullmatch(r"\$?[A-Za-z]+\$?\d+", ident):
                    tokens.append(("REF", ident))
                elif ident.upper() in ("TRUE", "FALSE"):
                    tokens.append(("BOOL", ident.upper()))
                else:
                    tokens.append(("IDENT", ident))
                i = j
                continue
            raise ValueError(f"Unexpected char '{ch}' in formula: {expr}")
        return tokens

    def _eval_expr(self, expr: str, default_sheet: str) -> Any:
        tokens = self._tokenize(expr.lstrip("="))
        pos = 0

        def peek() -> tuple[str, str] | None:
            return tokens[pos] if pos < len(tokens) else None

        def consume() -> tuple[str, str]:
            nonlocal pos
            t = tokens[pos]
            pos += 1
            return t

        def parse_cmp() -> Any:
            left = parse_add()
            while (t := peek()) is not None and t[0] == "OP" and t[1] in ("=", "<>", ">", "<", ">=", "<="):
                op = consume()[1]
                right = parse_add()
                if op == "=":
                    if _is_num(left) and _is_num(right):
                        left = abs(float(left) - float(right)) < 1e-9
                    else:
                        left = str(left or "") == str(right or "")
                elif op == "<>":
                    if _is_num(left) and _is_num(right):
                        left = abs(float(left) - float(right)) >= 1e-9
                    else:
                        left = str(left or "") != str(right or "")
                elif op == ">":
                    left = _to_num(left) > _to_num(right)
                elif op == "<":
                    left = _to_num(left) < _to_num(right)
                elif op == ">=":
                    left = _to_num(left) >= _to_num(right)
                elif op == "<=":
                    left = _to_num(left) <= _to_num(right)
            return left

        def parse_add() -> Any:
            left = parse_mul()
            while (t := peek()) is not None and t[0] == "OP" and t[1] in ("+", "-"):
                op = consume()[1]
                right = parse_mul()
                left = _to_num(left) + _to_num(right) if op == "+" else _to_num(left) - _to_num(right)
            return left

        def parse_mul() -> Any:
            left = parse_unary()
            while (t := peek()) is not None and t[0] == "OP" and t[1] in ("*", "/"):
                op = consume()[1]
                right = parse_unary()
                if op == "*":
                    left = _to_num(left) * _to_num(right)
                else:
                    denom = _to_num(right)
                    left = _to_num(left) / denom if abs(denom) > 1e-12 else 0.0
            return left

        def parse_unary() -> Any:
            if (t := peek()) is not None and t[0] == "OP" and t[1] in ("+", "-"):
                op = consume()[1]
                val = parse_unary()
                return -_to_num(val) if op == "-" else _to_num(val)
            return parse_atom()

        def _parse_arg_raw() -> Any:
            # Used for functions that need raw range references (SUM, SUMIFS, COUNTIFS, INDEX, MATCH)
            t = peek()
            if t is not None and t[0] == "RANGE":
                consume()
                rng = t[1]
                if "!" in rng:
                    sh, coords = rng.split("!", 1)
                else:
                    sh, coords = default_sheet, rng
                s_c, e_c = coords.split(":", 1)
                return ("__RANGE__", self.get_range(sh, s_c, e_c))
            return parse_cmp()

        def _unwrap_vals(arg: Any) -> list[Any]:
            if isinstance(arg, tuple) and len(arg) == 2 and arg[0] == "__RANGE__":
                return arg[1]
            return [arg]

        def parse_atom() -> Any:
            t = peek()
            if t is None:
                return 0.0
            kind, val = t
            if kind == "NUM":
                consume()
                return float(val) if "." in val else int(val)
            if kind == "STR":
                consume()
                return val
            if kind == "BOOL":
                consume()
                return val == "TRUE"
            if kind == "REF":
                consume()
                if "!" in val:
                    sh, coord = val.split("!", 1)
                else:
                    sh, coord = default_sheet, val
                return self.get_cell(sh, coord)
            if kind == "SYM" and val == "(":
                consume()
                inner = parse_cmp()
                if peek() == ("SYM", ")"):
                    consume()
                return inner
            if kind == "FUNC":
                fn = consume()[1]
                assert consume() == ("SYM", "(")
                args: list[Any] = []
                if peek() != ("SYM", ")"):
                    while True:
                        args.append(_parse_arg_raw())
                        if peek() == ("SYM", ","):
                            consume()
                        else:
                            break
                assert consume() == ("SYM", ")")
                return self._call_func(fn, args)
            raise ValueError(f"Unexpected token {t} in {expr}")

        return parse_cmp()

    def _call_func(self, fn: str, args: list[Any]) -> Any:
        def flat(a: Any) -> list[Any]:
            if isinstance(a, tuple) and len(a) == 2 and a[0] == "__RANGE__":
                return a[1]
            return [a]

        if fn == "SUM":
            total = 0.0
            for a in args:
                for v in flat(a):
                    if _is_num(v):
                        total += float(v)
            return total
        if fn == "ROUND":
            return _excel_round(_to_num(args[0]), int(_to_num(args[1])))
        if fn == "ABS":
            return abs(_to_num(args[0]))
        if fn == "MAX":
            vals = [_to_num(v) for a in args for v in flat(a)]
            return max(vals) if vals else 0.0
        if fn == "MIN":
            vals = [_to_num(v) for a in args for v in flat(a)]
            return min(vals) if vals else 0.0
        if fn == "IF":
            cond = bool(args[0])
            return args[1] if cond else (args[2] if len(args) > 2 else False)
        if fn == "OR":
            return any(bool(v) for a in args for v in flat(a))
        if fn == "AND":
            return all(bool(v) for a in args for v in flat(a))
        if fn == "ISNUMBER":
            return _is_num(args[0])
        if fn == "SEARCH":
            needle = str(args[0] or "").lower()
            hay = str(args[1] or "").lower()
            idx = hay.find(needle)
            return idx + 1 if idx >= 0 else False
        if fn == "YEAR":
            dt = EXCEL_EPOCH + timedelta(days=int(_to_num(args[0])))
            return dt.year
        if fn == "MONTH":
            dt = EXCEL_EPOCH + timedelta(days=int(_to_num(args[0])))
            return dt.month
        if fn == "DATE":
            dt = date(int(_to_num(args[0])), int(_to_num(args[1])), int(_to_num(args[2])))
            return (dt - EXCEL_EPOCH).days
        if fn == "SUMIFS":
            sum_vals = flat(args[0])
            pairs = [(flat(args[i]), args[i + 1]) for i in range(1, len(args), 2)]
            total = 0.0
            for idx in range(len(sum_vals)):
                if all(_matches_criterion(crit_rng[idx], crit_val) for crit_rng, crit_val in pairs):
                    if _is_num(sum_vals[idx]):
                        total += float(sum_vals[idx])
            return total
        if fn == "COUNTIFS":
            pairs = [(flat(args[i]), args[i + 1]) for i in range(0, len(args), 2)]
            n = len(pairs[0][0]) if pairs else 0
            cnt = 0
            for idx in range(n):
                if all(_matches_criterion(crit_rng[idx], crit_val) for crit_rng, crit_val in pairs):
                    cnt += 1
            return cnt
        if fn == "MATCH":
            needle = args[0]
            rng = flat(args[1])
            for idx, v in enumerate(rng, 1):
                if _matches_criterion(v, needle):
                    return idx
            return None
        if fn == "INDEX":
            rng = flat(args[0])
            pos = int(_to_num(args[1]))
            return rng[pos - 1] if 1 <= pos <= len(rng) else None
        raise ValueError(f"Unsupported Excel function in verifier: {fn}")


# ---------------------------------------------------------------------------
# Verification Layers
# ---------------------------------------------------------------------------

def verify_independent_source_tieout(
    books_path: Path,
    parsed_path: Path | None = None,
    csv_path: Path | None = None,
) -> list[dict[str, Any]]:
    """Verify final books & CSV directly against raw parsed_statements.json (independent of ledger_engine)."""
    books = json.loads(books_path.read_text(encoding="utf-8"))
    deltas: list[dict[str, Any]] = []

    if parsed_path and parsed_path.exists():
        parsed_docs = json.loads(parsed_path.read_text(encoding="utf-8"))
        excl_keys = {(d["excluded_statement"], int(d["excluded_seq"])) for d in books.get("duplicates", [])}

        # Check statement headers directly from parsed_statements.json
        src_net_total = 0.0
        included_count = 0
        for d in parsed_docs:
            if "rows" not in d or d.get("error"):
                continue
            rows = d["rows"]
            cr_sum = round(sum(float(r["credit"]) for r in rows), 2)
            dr_sum = round(sum(float(r["debit"]) for r in rows), 2)
            deltas.append(
                make_delta(
                    f"source.{d['id']}.credits",
                    f"{d['id']} primary credits tie-out",
                    d.get("credits"),
                    cr_sum,
                    provenance="statement_header_control",
                    tolerance=BALANCE_TOLERANCE,
                )
            )
            deltas.append(
                make_delta(
                    f"source.{d['id']}.debits",
                    f"{d['id']} primary debits tie-out",
                    d.get("debits"),
                    dr_sum,
                    provenance="statement_header_control",
                    tolerance=BALANCE_TOLERANCE,
                )
            )
            for r in rows:
                if (d["id"], int(r["seq"])) not in excl_keys:
                    src_net_total += float(r["credit"]) - float(r["debit"])
                    included_count += 1
        src_net_total = round(src_net_total, 2)
        deltas.append(
            make_delta(
                "source.included_row_count",
                "included source rows vs reconciled raw count",
                included_count,
                len(books["raw"]),
                provenance="input_premeasurement",
                tolerance=0.0,
            )
        )
    else:
        src_net_total = round(sum(float(r["credit"]) - float(r["debit"]) for r in books["raw"]), 2)

    raw_net = round(sum(float(r["credit"]) - float(r["debit"]) for r in books["raw"]), 2)
    ledger_net = round(sum(float(e["amount"]) for e in books["ledger"]), 2)

    deltas.append(
        make_delta(
            "source.raw_vs_primary_net",
            "reconciled raw net vs primary parsed rows net",
            src_net_total,
            raw_net,
            provenance="input_premeasurement",
            tolerance=BALANCE_TOLERANCE,
        )
    )
    deltas.append(
        make_delta(
            "source.ledger_vs_primary_net",
            "ledger net vs primary parsed rows net",
            src_net_total,
            ledger_net,
            provenance="input_premeasurement",
            tolerance=BALANCE_TOLERANCE,
        )
    )

    if csv_path and csv_path.exists():
        with csv_path.open(encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))
        deltas.append(
            make_delta(
                "csv.row_count",
                "CSV row count vs ledger component count",
                len(books["ledger"]),
                len(rows),
                provenance="input_premeasurement",
                tolerance=0.0,
            )
        )
        csv_sum = round(sum(float(r["amount"]) for r in rows), 2)
        deltas.append(
            make_delta(
                "csv.net_cash_tieout",
                "CSV net sum vs primary parsed rows net",
                src_net_total,
                csv_sum,
                provenance="input_premeasurement",
                tolerance=BALANCE_TOLERANCE,
            )
        )

    return deltas


def verify_math_and_bridge(computed: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    sp = computed["summary_pnl"]
    h = computed["health"]
    ledger = computed["ledger"]
    raw = computed["raw_transactions"]

    raw_net = round(sum(float(r["credit"]) - float(r["debit"]) for r in raw), 2)
    ledger_net = round(sum(float(e["amount"]) for e in ledger), 2)

    gp_calc = round(sp["revenue"]["total"] - sp["direct_costs"]["total"], 2)
    opex_sum = round(sum(sp["opex_buckets"][b]["total"] for b in OPEX_BUCKETS), 2)
    op_calc = round(sp["gross_profit"]["total"] - sp["total_opex"]["total"], 2)
    net_calc = round(sp["operating_profit"]["total"] - sp["taxes_and_fees"]["total"], 2)

    bridge_cash = round(
        sp["net_profit"]["total"]
        + sp["unclassified_in"]["total"]
        - sp["unclassified_out"]["total"]
        + sp["non_pnl_net"]["total"],
        2,
    )
    expected_close = round(h["opening_cash"] + bridge_cash, 2)

    deltas = [
        make_delta("math.ledger_vs_raw", "ledger net vs raw net", raw_net, ledger_net, "input_premeasurement"),
        make_delta("math.gross_profit", "gross profit equation", gp_calc, sp["gross_profit"]["total"], "specification"),
        make_delta("math.opex_buckets", "opex bucket sum vs total opex", opex_sum, sp["total_opex"]["total"], "specification"),
        make_delta("math.operating_profit", "operating profit equation", op_calc, sp["operating_profit"]["total"], "specification"),
        make_delta("math.net_profit", "net profit equation", net_calc, sp["net_profit"]["total"], "specification"),
        make_delta("math.cash_bridge", "cash bridge vs actual bank cash change", sp["actual_cash_change"]["total"], bridge_cash, "specification"),
        make_delta("math.bridge_vs_ledger", "cash bridge vs ledger net", ledger_net, bridge_cash, "input_premeasurement"),
        make_delta("math.opening_plus_change_vs_close", "opening cash + change vs closing cash", h["closing_cash"], expected_close, "statement_header_control"),
    ]
    for t in computed["transfers"]:
        pair_net = round(sum(e["amount"] for e in ledger if e.get("transfer_id") == t["id"]), 2)
        deltas.append(make_delta(f"math.transfer.{t['id']}", f"transfer {t['id']} pair net", 0.0, pair_net, "matched_transfer_pair"))

    for d in deltas:
        assert d["status"] == "pass", f"Math check failed: {d['check_id']} (expected={d['expected']}, measured={d['measured']}, delta={d['delta']})"

    summary = {
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
    return summary, deltas


def _read_xlsx_sheets(xlsx_path: Path) -> tuple[list[str], dict[str, dict[str, tuple[Any, str | None]]], int]:
    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    sheets_data: dict[str, dict[str, tuple[Any, str | None]]] = {}
    formula_count = 0

    with ZipFile(xlsx_path) as z:
        assert z.testzip() is None, "Corrupt XLSX zip archive"
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            sst = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sst.findall("s:si", ns):
                shared_strings.append("".join(t.text or "" for t in si.findall(".//s:t", ns)))

        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        sheets = [s.attrib["name"] for s in wb_xml.find("s:sheets", ns)]
        assert sheets == EXPECTED_SHEETS, f"Sheet order mismatch: {sheets} != {EXPECTED_SHEETS}"
        sheet_set = set(sheets)

        for idx, s_name in enumerate(sheets, 1):
            tree = ET.fromstring(z.read(f"xl/worksheets/sheet{idx}.xml"))
            cell_map: dict[str, tuple[Any, str | None]] = {}
            for cell in tree.findall(".//s:sheetData/s:row/s:c", ns):
                coord = cell.attrib.get("r", "").upper()
                ctype = cell.attrib.get("t", "n")
                assert ctype != "e", f"Excel error cell at {s_name}!{coord}"

                f_node = cell.find("s:f", ns)
                v_node = cell.find("s:v", ns)
                is_node = cell.find("s:is", ns)
                formula_txt = f_node.text if (f_node is not None and f_node.text) else None

                if formula_txt:
                    formula_count += 1
                    stripped = re.sub(r'"[^"]*"', '""', formula_txt)
                    assert stripped.count("(") == stripped.count(")"), (
                        f"Unbalanced parentheses in {s_name}!{coord}: {formula_txt}"
                    )
                    for ref_sheet in re.findall(r"'([^']+)'!", stripped):
                        assert ref_sheet in sheet_set, (
                            f"Unknown sheet reference '{ref_sheet}' in {s_name}!{coord}"
                        )

                val: Any = None
                if ctype == "inlineStr" and is_node is not None:
                    val = "".join(t.text or "" for t in is_node.findall(".//s:t", ns))
                elif ctype == "s" and v_node is not None and v_node.text is not None:
                    val = shared_strings[int(v_node.text)]
                elif ctype in ("str", "b") and v_node is not None and v_node.text is not None:
                    val = (v_node.text == "1") if ctype == "b" else v_node.text
                elif v_node is not None and v_node.text is not None:
                    raw_t = v_node.text
                    val = float(raw_t) if ("." in raw_t or "e" in raw_t.lower()) else int(raw_t)

                cell_map[coord] = (val, formula_txt)
            sheets_data[s_name] = cell_map

    return sheets, sheets_data, formula_count


def verify_xlsx(xlsx_path: Path, computed: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    sheets, sheets_data, formula_count = _read_xlsx_sheets(xlsx_path)
    ev = ExcelFormulaEvaluator(sheets_data)
    deltas: list[dict[str, Any]] = []
    evaluated_count = 0

    # Evaluate ALL formula cells across all 14 sheets (catches broken syntax, refs, or runtime formula errors)
    for s_name in sheets:
        for coord, (_, ftxt) in sheets_data[s_name].items():
            if ftxt is not None:
                ev.get_cell(s_name, coord)
                evaluated_count += 1

    # 1. Verify Reconciliations tab row checks & cross-workbook control checks evaluate to 0.00
    rc_sheet = sheets_data["Reconciliations"]
    rc_rows = sorted(
        {
            int(re.match(r"^[A-Z]+(\d+)$", k).group(1))
            for k in rc_sheet
            if re.match(r"^A(\d+)$", k) and int(re.match(r"^A(\d+)$", k).group(1)) >= 6
        }
    )
    for r in rc_rows:
        a_val = ev.get_cell("Reconciliations", f"A{r}")
        if a_val in ("Cross-Workbook Control Check", "Source ID", None, ""):
            continue
        # Statement row or Cross-Workbook audit row
        if isinstance(ev.get_cell("Reconciliations", f"C{r}"), (int, float)):
            for col in ("G", "J", "K", "L", "M"):
                v = _to_num(ev.get_cell("Reconciliations", f"{col}{r}"))
                d = make_delta(f"xlsx.Reconciliations.{col}{r}", f"Reconciliations!{col}{r}", 0.0, v, "statement_header_control")
                deltas.append(d)
                assert d["status"] == "pass", f"Excel formula check failed at Reconciliations!{col}{r}: {v}"
        else:
            v = _to_num(ev.get_cell("Reconciliations", f"B{r}"))
            d = make_delta(f"xlsx.Reconciliations.B{r}", f"Cross-workbook check '{a_val}'", 0.0, v, "specification")
            deltas.append(d)
            assert d["status"] == "pass", f"Excel Cross-Workbook check '{a_val}' failed: {v}"

    # 2. Verify Account Cash Closing Difference (col K) == 0.00
    ac_rows = [
        int(m.group(1))
        for k in sheets_data["Account Cash"]
        if (m := re.match(r"^A(\d+)$", k)) and int(m.group(1)) >= 6
    ]
    for r in ac_rows:
        v = _to_num(ev.get_cell("Account Cash", f"K{r}"))
        d = make_delta(f"xlsx.AccountCash.K{r}", f"Account Cash!K{r} closing diff", 0.0, v, "statement_header_control")
        deltas.append(d)
        assert d["status"] == "pass", f"Excel Account Cash!K{r} closing diff != 0: {v}"

    # 3. Verify Summary P&L cumulative column matches expected Summary P&L values
    sp = computed["summary_pnl"]
    cum_col = _idx_to_col(len(computed["months"]) + 2)
    label_to_expected = {
        "Revenue (Net of Refunds)": sp["revenue"]["total"],
        computed["profile"]["direct_cost_label"]: sp["direct_costs"]["total"],
        "GROSS PROFIT / (LOSS)": sp["gross_profit"]["total"],
        "Total Operating Expenses": sp["total_opex"]["total"],
        "OPERATING PROFIT / (LOSS)": sp["operating_profit"]["total"],
        "Taxes & Statutory Levies": sp["taxes"]["total"],
        "Bank & Payment Fees": sp["bank_fees"]["total"],
        "Total Taxes & Bank Fees": sp["taxes_and_fees"]["total"],
        "PROVISIONAL NET PROFIT / (LOSS)": sp["net_profit"]["total"],
        "Cash Bridge: Unclassified Receipts (Excluded from P&L)": sp["unclassified_in"]["total"],
        "Cash Bridge: Unclassified Payments (Excluded from P&L)": -sp["unclassified_out"]["total"],
        "Cash Bridge: Non-P&L Cash Movements (Loans, Assets, Owner & Other Accounts)": sp["non_pnl_net"]["total"],
        "ACTUAL NET BANK CASH CHANGE": sp["actual_cash_change"]["total"],
    }
    sp_rows_found = 0
    for coord, (val, _) in sheets_data["Summary P&L"].items():
        if coord.startswith("A") and val in label_to_expected:
            r_num = coord[1:]
            eval_val = _to_num(ev.get_cell("Summary P&L", f"{cum_col}{r_num}"))
            exp_val = label_to_expected[val]
            d = make_delta(
                f"xlsx.SummaryPnL.{cum_col}{r_num}",
                f"Summary P&L '{val}' formula evaluation",
                exp_val,
                round(eval_val, 2),
                "input_premeasurement",
            )
            deltas.append(d)
            assert d["status"] == "pass", f"Excel Summary P&L '{val}' evaluated to {eval_val}, expected {exp_val}"
            sp_rows_found += 1
    assert sp_rows_found == len(label_to_expected), "Missing expected rows in Summary P&L sheet"

    # 4. Verify Executive Overview headline net profit & actual bank cash change
    ov_net = _to_num(ev.get_cell("Executive Overview", "B10"))
    ov_cash = _to_num(ev.get_cell("Executive Overview", "B14"))
    d_ov1 = make_delta("xlsx.Overview.B10", "Executive Overview Net Profit", sp["net_profit"]["total"], round(ov_net, 2), "input_premeasurement")
    d_ov2 = make_delta("xlsx.Overview.B14", "Executive Overview Net Cash Change", sp["actual_cash_change"]["total"], round(ov_cash, 2), "input_premeasurement")
    deltas.extend([d_ov1, d_ov2])
    assert d_ov1["status"] == "pass", f"Executive Overview!B10 mismatch: {ov_net} != {sp['net_profit']['total']}"
    assert d_ov2["status"] == "pass", f"Executive Overview!B14 mismatch: {ov_cash} != {sp['actual_cash_change']['total']}"

    max_delta = max((abs(d["delta"]) for d in deltas if d["delta"] is not None), default=0.0)
    return {
        "sheets": len(sheets),
        "formula_cells": formula_count,
        "formula_cells_evaluated": evaluated_count,
        "max_formula_delta": round(max_delta, 4),
    }, deltas


def verify_dashboard_html(html_path: Path, computed: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    html = html_path.read_text(encoding="utf-8")
    m = re.search(r'<script id="books-payload" type="application/json">(.*?)</script>', html, re.DOTALL)
    assert m is not None, "Missing embedded books-payload JSON in dashboard HTML"
    payload = json.loads(m.group(1).replace("<\\/", "</"))

    deltas = [
        make_delta("html.ledger_count", "HTML ledger count", len(computed["ledger"]), len(payload["ledger"]), "input_premeasurement", tolerance=0.0),
        make_delta("html.net_profit", "HTML Summary P&L net profit", computed["summary_pnl"]["net_profit"]["total"], payload["summary_pnl"]["net_profit"]["total"], "input_premeasurement"),
        make_delta("html.actual_cash_change", "HTML actual bank cash change", computed["summary_pnl"]["actual_cash_change"]["total"], payload["summary_pnl"]["actual_cash_change"]["total"], "input_premeasurement"),
        make_delta("html.closing_cash", "HTML closing cash", computed["health"]["closing_cash"], payload["health"]["closing_cash"], "statement_header_control"),
    ]
    for d in deltas:
        assert d["status"] == "pass", f"HTML dashboard check failed: {d['check_id']} ({d['expected']} vs {d['measured']})"

    return {"html_kb": round(len(html) / 1024, 1), "smart_groups": len(payload["smart_review_groups"])}, deltas


def verify_health_report_guardrails(md_path: Path, computed: dict[str, Any]) -> dict[str, Any]:
    text = md_path.read_text(encoding="utf-8")
    for pat in FORBIDDEN_ADVISORY_PATTERNS:
        m = re.search(pat, text, re.IGNORECASE)
        assert m is None, f"Forbidden prescriptive language '{m.group(0)}' found in Health Report!"
    assert "## 1. Summary Profit & Loss" in text
    assert "## 3. Where Costs Are" in text
    # Verify key figures appear in report text
    assert f"{abs(computed['summary_pnl']['actual_cash_change']['total']):,.2f}" in text, "Actual Net Bank Cash Change missing from Health Report"
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
        "provenance": "owner_confirmed",
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

    src_deltas = verify_independent_source_tieout(books_path, parsed_path, csv_path)
    for d in src_deltas:
        assert d["status"] == "pass", f"Independent primary-source tie-out failed: {d['check_id']} ({d['expected']} vs {d['measured']})"

    math_res, math_deltas = verify_math_and_bridge(computed)
    xlsx_res, xlsx_deltas = verify_xlsx(xlsx_path, computed)
    html_res, html_deltas = verify_dashboard_html(html_path, computed)
    md_res = verify_health_report_guardrails(md_path, computed)
    roundtrip_ok = verify_override_roundtrip(books_path, classifications_path)

    all_deltas = [*src_deltas, *math_deltas, *xlsx_deltas, *html_deltas]
    vdc_st = computed["health"]["vdc_status"]
    unverified_range = sorted(set(UNVERIFIED_PROPERTIES["verify_package"]) | set(vdc_st["unverified_range"]))

    vdc_report = make_vdc_report(
        deltas=all_deltas,
        unverified_range=unverified_range,
        math_integrity="VERIFIED",
        accounting_completeness=vdc_st["accounting_completeness"],
        attribution_state="none",
        attribution_grounds="All primary-source, ledger, Excel formula, HTML, CSV, and guardrail checks passed.",
    )

    return {
        "status": "PASSED",
        "math_integrity_verdict": "VERIFIED",
        "accounting_completeness_verdict": vdc_st["accounting_completeness"],
        "math": math_res,
        "xlsx": xlsx_res,
        "dashboard": html_res,
        "health_report": md_res,
        "override_roundtrip": roundtrip_ok,
        "vdc_report": vdc_report,
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
