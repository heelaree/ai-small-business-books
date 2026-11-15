#!/usr/bin/env python3
"""Single-sourced machine-readable contract, JSON schemas, and VDC report helpers.

Implements Criterion 2 of Verification-Driven Coding (Kajiyama, 2026):
  - Fixes permitted operations, typed flags, tolerances, and unverified properties in one place.
  - Generates canonical JSON Schemas mechanically (`--write-schemas`).
  - Probes runtime dependencies (`doctor`).
  - Reconciles contract declarations against actual scripts, schemas, and SKILL.md (`check-drift`).
"""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import re
import sys
from typing import Any

CONTRACT_VERSION = "2.0.0-vdc"
REFERENCE_VERSION = "cash-basis-vdc-2026-11"
BALANCE_TOLERANCE = 0.01

PROVENANCE_CLASSES = (
    "statement_header_control",
    "input_premeasurement",
    "specification",
    "owner_confirmed",
    "statement_narration_rule",
    "matched_transfer_pair",
    "embedded_fee_split",
    "unclassified_default",
    "downgraded_unverified_claim",
    "synthetic_fixture",
)

UNVERIFIED_PROPERTIES: dict[str, list[str]] = {
    "probe_statements": [
        "contents_of_encrypted_pdfs",
        "off_bank_cash_and_petty_cash",
        "physical_receipt_and_invoice_authenticity",
    ],
    "reconcile_ledger": [
        "business_purpose_of_unconfirmed_counterparties",
        "contents_of_encrypted_pdfs",
        "off_bank_cash_and_petty_cash",
        "accrual_items_and_input_tax_recoverability",
    ],
    "build_all": [
        "visual_browser_rendering_and_js_interaction",
        "business_purpose_of_unconfirmed_counterparties",
        "contents_of_encrypted_pdfs",
        "off_bank_cash_and_petty_cash",
    ],
    "verify_package": [
        "visual_browser_rendering_and_js_interaction",
        "business_purpose_of_unconfirmed_counterparties",
        "contents_of_encrypted_pdfs",
        "off_bank_cash_and_petty_cash",
    ],
}

TOOLS: dict[str, dict[str, Any]] = {
    "probe_statements": {
        "script": "probe_statements.py",
        "role": "Pre-measurement of raw statement files or parsed_statements.json before reconciliation.",
        "args": ["--dir", "--parsed", "--out", "--json"],
        "unverified_properties": UNVERIFIED_PROPERTIES["probe_statements"],
    },
    "reconcile_ledger": {
        "script": "reconcile_ledger.py",
        "role": "Deduplicate consolidated rows, chain running balances, verify statement header controls, match transfers, split embedded fees, enforce provenance, and emit books_data.json.",
        "args": ["--parsed", "--out", "--entity-keyword", "--allow-unresolved", "--json"],
        "unverified_properties": UNVERIFIED_PROPERTIES["reconcile_ledger"],
    },
    "generate_dashboard": {
        "script": "generate_dashboard.py",
        "role": "Render standalone interactive HTML dashboard from reconciled books_data.json.",
        "args": ["--books", "--classifications", "--out"],
        "unverified_properties": UNVERIFIED_PROPERTIES["build_all"],
    },
    "generate_health_report": {
        "script": "generate_health_report.py",
        "role": "Render strictly factual Financial Health & Cost Report Markdown.",
        "args": ["--books", "--classifications", "--out"],
        "unverified_properties": UNVERIFIED_PROPERTIES["build_all"],
    },
    "generate_workbook": {
        "script": "generate_workbook.py",
        "role": "Generate 14-sheet formula-driven Excel workbook and CSV cash ledger.",
        "args": ["--books", "--classifications", "--parsed", "--out-xlsx", "--out-csv"],
        "unverified_properties": UNVERIFIED_PROPERTIES["build_all"],
    },
    "build_all": {
        "script": "build_all.py",
        "role": "Generate all 4 deliverables and run verify_package verification.",
        "args": ["--books", "--out-dir", "--slug", "--classifications", "--parsed", "--html-out"],
        "unverified_properties": UNVERIFIED_PROPERTIES["build_all"],
    },
    "verify_package": {
        "script": "verify_package.py",
        "role": "Independent primary-source tie-out, Excel formula evaluator, HTML/CSV cross-check, advisory guardrail scan, and enforced verdict cap.",
        "args": ["--books", "--xlsx", "--html", "--md", "--csv", "--classifications", "--parsed"],
        "unverified_properties": UNVERIFIED_PROPERTIES["verify_package"],
    },
    "probe_sensitivity": {
        "script": "probe_sensitivity.py",
        "role": "Layer 2 seeded-defect sensitivity probing across 10 financial defect classes + unseeded control.",
        "args": ["--parsed", "--classifications", "--entity-keyword", "--json"],
        "unverified_properties": [],
    },
}

VDC_DELTA_SCHEMA: dict[str, Any] = {
    "type": "object",
    "required": ["check_id", "property", "expected", "measured", "delta", "tolerance", "provenance", "status"],
    "properties": {
        "check_id": {"type": "string"},
        "property": {"type": "string"},
        "expected": {"type": ["number", "string", "null"]},
        "measured": {"type": ["number", "string", "null"]},
        "delta": {"type": ["number", "null"]},
        "tolerance": {"type": ["number", "null"]},
        "provenance": {"type": "string", "enum": list(PROVENANCE_CLASSES)},
        "status": {"type": "string", "enum": ["pass", "fail", "unknown"]},
        "notes": {"type": "string"},
    },
}

VDC_REPORT_SCHEMA: dict[str, Any] = {
    "type": "object",
    "required": ["reference_version", "deltas", "unverified_range", "verdict", "attribution"],
    "properties": {
        "reference_version": {"type": "string"},
        "deltas": {"type": "array", "items": VDC_DELTA_SCHEMA},
        "unverified_range": {"type": "array", "items": {"type": "string"}},
        "verdict": {
            "type": "object",
            "required": ["math_integrity", "accounting_completeness"],
            "properties": {
                "math_integrity": {"type": "string", "enum": ["VERIFIED", "FAILED", "UNRESOLVED"]},
                "accounting_completeness": {"type": "string", "enum": ["VERIFIED", "PROVISIONAL_REVIEW_REQUIRED"]},
            },
        },
        "attribution": {
            "type": "object",
            "required": ["state", "grounds"],
            "properties": {
                "state": {
                    "type": "string",
                    "enum": ["none", "artifact", "reference", "instrument", "provenance", "unresolved"],
                },
                "grounds": {"type": "string"},
                "next_measurement": {"type": "string"},
            },
        },
    },
}

PARSED_STATEMENTS_SCHEMA: dict[str, Any] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "title": "ParsedStatementsInput",
    "description": "Canonical schema for parsed_statements.json consumed by reconcile_ledger.py and probe_statements.py.",
    "type": "array",
    "items": {
        "type": "object",
        "required": ["id", "file"],
        "properties": {
            "id": {"type": "string"},
            "file": {"type": "string"},
            "sha256": {"type": "string"},
            "encrypted": {"type": "boolean"},
            "error": {"type": "string"},
            "page_count": {"type": ["integer", "null"]},
            "account": {"type": "string"},
            "currency": {"type": "string"},
            "owner": {"type": "string"},
            "start": {"type": "string", "format": "date"},
            "end": {"type": "string", "format": "date"},
            "opening": {"type": ["number", "null"]},
            "closing": {"type": ["number", "null"]},
            "credits": {"type": ["number", "null"]},
            "debits": {"type": ["number", "null"]},
            "is_consolidated": {"type": "boolean"},
            "rows": {
                "type": "array",
                "items": {
                    "type": "object",
                    "required": ["seq", "date", "description", "debit", "credit", "balance"],
                    "properties": {
                        "seq": {"type": "integer"},
                        "date": {"type": "string", "format": "date"},
                        "datetime": {"type": "string"},
                        "description": {"type": "string"},
                        "debit": {"type": "number"},
                        "credit": {"type": "number"},
                        "before": {"type": ["number", "null"]},
                        "balance": {"type": "number"},
                        "charge": {"type": "number"},
                        "principal": {"type": "number"},
                        "counterparty": {"type": "string"},
                        "source_account": {"type": "string"},
                        "reference": {"type": "string"},
                        "pages": {"type": "string"},
                        "statement": {"type": "string"},
                        "account": {"type": "string"},
                        "file": {"type": "string"},
                        "section": {"type": "string"},
                        "category": {"type": "string"},
                        "provenance": {"type": "string"},
                        "review_type": {"type": "string"},
                        "question": {"type": "string"},
                    },
                },
            },
        },
    },
}

BOOKS_DATA_SCHEMA: dict[str, Any] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "title": "SmallBusinessBooksData",
    "description": "Normalized contract between bank statement reconciliation (reconcile_ledger.py) and reporting generators (build_all.py).",
    "type": "object",
    "required": ["documents", "raw", "ledger", "duplicates", "transfers", "review", "issues"],
    "properties": {
        "documents": {
            "type": "array",
            "items": {
                "type": "object",
                "required": ["id", "file"],
                "properties": {
                    "id": {"type": "string"},
                    "file": {"type": "string"},
                    "sha256": {"type": "string"},
                    "encrypted": {"type": "boolean"},
                    "error": {"type": "string"},
                    "account": {"type": "string"},
                    "currency": {"type": "string"},
                    "owner": {"type": "string"},
                    "start": {"type": "string", "format": "date"},
                    "end": {"type": "string", "format": "date"},
                    "opening": {"type": ["number", "null"]},
                    "closing": {"type": ["number", "null"]},
                    "credits": {"type": ["number", "null"]},
                    "debits": {"type": ["number", "null"]},
                    "is_consolidated": {"type": "boolean"},
                    "running_breaks": {"type": "array", "items": {"type": "integer"}},
                    "reordered": {"type": "integer"},
                    "ending_diff": {"type": "number"},
                },
            },
        },
        "raw": {
            "type": "array",
            "items": {
                "type": "object",
                "required": ["raw_id", "statement", "seq", "date", "account", "description", "debit", "credit", "balance"],
                "properties": {
                    "raw_id": {"type": "string"},
                    "statement": {"type": "string"},
                    "seq": {"type": "integer"},
                    "date": {"type": "string", "format": "date"},
                    "datetime": {"type": "string"},
                    "account": {"type": "string"},
                    "description": {"type": "string"},
                    "debit": {"type": "number"},
                    "credit": {"type": "number"},
                    "balance": {"type": "number"},
                    "charge": {"type": "number"},
                    "counterparty": {"type": "string"},
                    "source_account": {"type": "string"},
                    "reference": {"type": "string"},
                    "pages": {"type": "string"},
                },
            },
        },
        "ledger": {
            "type": "array",
            "items": {
                "type": "object",
                "required": ["id", "raw_id", "date", "account", "description", "component", "amount", "section", "category"],
                "properties": {
                    "id": {"type": "string"},
                    "raw_id": {"type": "string"},
                    "date": {"type": "string", "format": "date"},
                    "account": {"type": "string"},
                    "description": {"type": "string"},
                    "component": {"enum": ["Principal", "Embedded fee"]},
                    "amount": {"type": "number"},
                    "section": {"type": "string"},
                    "category": {"type": "string"},
                    "provenance": {"type": "string"},
                    "transfer_id": {"type": "string"},
                    "question": {"type": "string"},
                },
            },
        },
        "duplicates": {"type": "array"},
        "transfers": {"type": "array"},
        "review": {"type": "array"},
        "issues": {"type": "array"},
        "vdc_report": VDC_REPORT_SCHEMA,
    },
}

CLASSIFICATIONS_SCHEMA: dict[str, Any] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "title": "OwnerClassificationsConfig",
    "description": "Durable owner confirmations and overrides exported from the HTML dashboard or recorded from chat answers.",
    "type": "object",
    "required": ["schema_version", "business_profile", "category_group_overrides", "counterparty_rules", "transaction_overrides"],
    "properties": {
        "schema_version": {"type": "string"},
        "business_profile": {
            "type": "object",
            "required": ["business_name", "legal_entity", "business_type", "currency", "direct_cost_label", "reporting_basis"],
            "properties": {
                "business_name": {"type": "string"},
                "legal_entity": {"type": "string"},
                "business_type": {"type": "string"},
                "industry": {"type": "string"},
                "currency": {"type": "string"},
                "direct_cost_label": {"type": "string"},
                "reporting_basis": {"type": "string"},
            },
        },
        "category_group_overrides": {"type": "object", "additionalProperties": {"type": "string"}},
        "counterparty_rules": {
            "type": "array",
            "items": {
                "type": "object",
                "required": ["section", "category"],
                "properties": {
                    "counterparty_key": {"type": "string"},
                    "match_text": {"type": "string"},
                    "section": {"type": "string"},
                    "category": {"type": "string"},
                    "note": {"type": "string"},
                    "provenance": {"type": "string"},
                },
            },
        },
        "transaction_overrides": {
            "type": "object",
            "additionalProperties": {
                "type": "object",
                "required": ["section", "category"],
                "properties": {
                    "section": {"type": "string"},
                    "category": {"type": "string"},
                    "note": {"type": "string"},
                    "provenance": {"type": "string"},
                },
            },
        },
    },
}

SCHEMA_FILES: dict[str, dict[str, Any]] = {
    "parsed_statements.schema.json": PARSED_STATEMENTS_SCHEMA,
    "books_data.schema.json": BOOKS_DATA_SCHEMA,
    "classifications.schema.json": CLASSIFICATIONS_SCHEMA,
}


def make_delta(
    check_id: str,
    prop: str,
    expected: float | str | None,
    measured: float | str | None,
    provenance: str,
    tolerance: float | None = BALANCE_TOLERANCE,
    notes: str = "",
) -> dict[str, Any]:
    """Build a single VDC measurement delta record."""
    if expected is None or measured is None:
        return {
            "check_id": check_id,
            "property": prop,
            "expected": expected,
            "measured": measured,
            "delta": None,
            "tolerance": tolerance,
            "provenance": provenance,
            "status": "unknown",
            **({"notes": notes} if notes else {}),
        }
    if isinstance(expected, (int, float)) and isinstance(measured, (int, float)):
        diff = round(float(measured) - float(expected), 4)
        tol = 0.0 if tolerance is None else float(tolerance)
        status = "pass" if abs(diff) <= tol + 1e-9 else "fail"
        return {
            "check_id": check_id,
            "property": prop,
            "expected": round(float(expected), 4),
            "measured": round(float(measured), 4),
            "delta": diff,
            "tolerance": tol,
            "provenance": provenance,
            "status": status,
            **({"notes": notes} if notes else {}),
        }
    status = "pass" if expected == measured else "fail"
    return {
        "check_id": check_id,
        "property": prop,
        "expected": expected,
        "measured": measured,
        "delta": None,
        "tolerance": None,
        "provenance": provenance,
        "status": status,
        **({"notes": notes} if notes else {}),
    }


def make_vdc_report(
    deltas: list[dict[str, Any]],
    unverified_range: list[str],
    math_integrity: str,
    accounting_completeness: str,
    attribution_state: str = "none",
    attribution_grounds: str = "All measured mathematical invariants within tolerance.",
    next_measurement: str = "",
) -> dict[str, Any]:
    """Build a canonical VDC operation report tuple r(i) = (Delta, U, W, Pi, Y, A)."""
    att: dict[str, Any] = {
        "state": attribution_state,
        "grounds": attribution_grounds,
    }
    if next_measurement:
        att["next_measurement"] = next_measurement
    return {
        "reference_version": REFERENCE_VERSION,
        "deltas": deltas,
        "unverified_range": sorted(set(unverified_range)),
        "verdict": {
            "math_integrity": math_integrity,
            "accounting_completeness": accounting_completeness,
        },
        "attribution": att,
    }


def build_contract_document() -> dict[str, Any]:
    return {
        "contract_version": CONTRACT_VERSION,
        "reference_version": REFERENCE_VERSION,
        "tolerances": {"monetary_tieout": BALANCE_TOLERANCE},
        "provenance_classes": list(PROVENANCE_CLASSES),
        "tools": TOOLS,
        "schemas": SCHEMA_FILES,
    }


def run_doctor() -> dict[str, Any]:
    """Pre-measure Python runtime and required library availability."""
    deps = {}
    for pkg in ("openpyxl", "pdfplumber", "pypdf"):
        spec = importlib.util.find_spec(pkg)
        if spec is None:
            deps[pkg] = {"status": "MISSING", "version": None}
        else:
            mod = __import__(pkg)
            deps[pkg] = {"status": "AVAILABLE", "version": getattr(mod, "__version__", "unknown")}
    all_ok = all(v["status"] == "AVAILABLE" for v in deps.values())
    return {
        "status": "AVAILABLE" if all_ok else "DEGRADED",
        "python": sys.version.split()[0],
        "python_ok": sys.version_info >= (3, 10),
        "dependencies": deps,
    }


def write_schemas(ref_dir: Path) -> list[str]:
    ref_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for fname, schema in SCHEMA_FILES.items():
        p = ref_dir / fname
        p.write_text(json.dumps(schema, indent=2) + "\n", encoding="utf-8")
        written.append(str(p))
    return written


def check_drift(skill_root: Path) -> dict[str, Any]:
    """Reconcile _contract.py against scripts/, references/*.schema.json, and SKILL.md."""
    scripts_dir = skill_root / "scripts"
    ref_dir = skill_root / "references"
    skill_md = skill_root / "SKILL.md"
    mismatches: list[str] = []

    for tool_name, spec in TOOLS.items():
        script_path = scripts_dir / spec["script"]
        if not script_path.exists():
            mismatches.append(f"Missing script for tool '{tool_name}': {script_path.name}")
            continue
        src = script_path.read_text(encoding="utf-8")
        for flag in spec["args"]:
            if flag not in src:
                mismatches.append(f"Flag '{flag}' declared in contract for '{tool_name}' not found in {script_path.name}")

    for fname, schema in SCHEMA_FILES.items():
        p = ref_dir / fname
        if not p.exists():
            mismatches.append(f"Missing schema file: references/{fname}")
        else:
            disk = json.loads(p.read_text(encoding="utf-8"))
            if disk != schema:
                mismatches.append(f"Schema drift detected in references/{fname} vs _contract.py")

    if skill_md.exists():
        md_txt = skill_md.read_text(encoding="utf-8")
        for tool_name, spec in TOOLS.items():
            if spec["script"] not in md_txt:
                mismatches.append(f"Script '{spec['script']}' from contract is not documented in SKILL.md")
        for fname in SCHEMA_FILES:
            if fname not in md_txt:
                mismatches.append(f"Schema '{fname}' from contract is not referenced in SKILL.md")

    return {
        "status": "PASSED" if not mismatches else "FAILED",
        "checked_tools": len(TOOLS),
        "checked_schemas": len(SCHEMA_FILES),
        "mismatches": mismatches,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="Single-sourced VDC contract, schema generator, and drift checker.")
    parser.add_argument("command", nargs="?", default="contract", choices=["contract", "doctor", "check-drift", "write-schemas"])
    parser.add_argument("--json", action="store_true", help="Output machine-readable JSON.")
    parser.add_argument("--write-schemas", action="store_true", help="Write canonical JSON schemas to references/.")
    args = parser.parse_args()

    skill_root = Path(__file__).resolve().parent.parent

    if args.write_schemas or args.command == "write-schemas":
        written = write_schemas(skill_root / "references")
        print(json.dumps({"status": "PASSED", "written_schemas": written}, indent=2))
        return

    if args.command == "doctor":
        res = run_doctor()
        print(json.dumps(res, indent=2))
        if not res["python_ok"] or res["status"] != "AVAILABLE":
            sys.exit(1)
        return

    if args.command == "check-drift":
        res = check_drift(skill_root)
        print(json.dumps(res, indent=2))
        if res["status"] != "PASSED":
            sys.exit(1)
        return

    print(json.dumps(build_contract_document(), indent=2))


if __name__ == "__main__":
    main()
