"""Regression tests: Inventory Audit Status tab must never show 'Unknown' district
when the Store List (master) or count file has the district.

User-reported bug (2026-10-24): Store List tab matched stores<->districts fine and
Variance Audit showed districts correctly, but the Inventory Audit Status tab
showed District = "Unknown" for every row.

Root cause: build_inventory_status_rows ignored master_store_records (and, in the
old copy, the count file's own District column) — districts came only from the
Employee Time Sheet, which the new flow no longer requires. Fix: district
priority chain = count file District -> Store List (master) -> time sheet -> Unknown.

Same AST-extraction technique as test_timesheet_store_matching.py.
"""
import ast
import datetime as dt
import hashlib
import re
import sys
import unittest
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

REPO_ROOT = Path(__file__).resolve().parent.parent

NEEDED_FUNCS = [
    "safe_text",
    "normalize_header",
    "normalize_store",
    "display_store",
    "normalize_district",
    "excel_serial_to_datetime",
    "numeric_excel_date",
    "find_column",
    "filter_latest_inventory_records",
    "build_inventory_status_rows",
]

# build_store_maps may take 1 arg (old copies) or 2 (new copies); both are fine.
_CANDIDATE_FUNCS = NEEDED_FUNCS + [
    "build_store_maps",
    "build_ts_store_to_employee",
    "store_name_tokens",
    "match_store_employee",
]


def _script_path() -> Path:
    for name in (REPO_ROOT / "GFH_Audit_Automation.py", REPO_ROOT / "GFH_Inventory_Audit.py",
                 REPO_ROOT / "GFH_Inventory_Audit_Timesheet.py"):
        if name.exists():
            return name
    raise RuntimeError("No audit app script found next to tests/")


def _extract_sources():
    tree = ast.parse(_script_path().read_text(encoding="utf-8"), filename=str(_script_path()))
    found = {}
    for node in tree.body:
        if isinstance(node, ast.FunctionDef) and node.name in _CANDIDATE_FUNCS:
            found[node.name] = ast.unparse(node)
        if isinstance(node, ast.ClassDef) and node.name == "InventoryStatusRow":
            found["InventoryStatusRow"] = ast.unparse(node)
    missing = set(NEEDED_FUNCS) - set(found)
    if missing:
        raise RuntimeError(f"Could not extract from script: {sorted(missing)}")
    return found


def _make_namespace():
    ns = {
        "re": re, "dt": dt, "hashlib": hashlib, "dataclass": dataclass,
        "Dict": Dict, "List": List, "Optional": Optional, "Tuple": Tuple,
        "Iterable": Iterable, "Any": Any,
    }
    for name, src in _extract_sources().items():
        exec(src, ns)
    return ns


def _count_row(store, created_by, created="45870.5", district=None, with_district_col=False):
    row = {
        "Store": store, "Created By": created_by, "Created Date/Time": created,
        "Serial #": "111111111111111", "Status": "Matched", "Product": "Phone",
    }
    if with_district_col:
        row["District"] = district if district is not None else ""
    return row


MASTER = [
    {"Store": "Belleview", "District": "Colorado"},
    {"Store": "Boulder", "District": "Colorado"},
    {"Store": "Clayton Store", "District": "Atlanta"},
    {"Store": "Fuqua", "District": "Houston"},
]


class TestStatusDistrictFallback(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        ns = _make_namespace()
        for name, val in ns.items():
            if name in _CANDIDATE_FUNCS or name == "InventoryStatusRow":
                setattr(cls, name, staticmethod(val) if callable(val) else val)

    def _rows(self, inv, ts=None, master=None):
        rows, _summary = self.build_inventory_status_rows(
            inv, ts or [], master_store_records=master)
        return {r.store: r.district for r in rows}

    def test_empty_timesheet_master_fills_districts(self):
        """THE BUG: count file + master present, timesheet empty -> real districts."""
        inv = [_count_row("Belleview", "alice"), _count_row("Clayton Store", "carl")]
        out = self._rows(inv, master=MASTER)
        self.assertEqual(out.get("Belleview"), "Colorado")
        self.assertEqual(out.get("Clayton Store"), "Atlanta")

    def test_district_column_empty_still_resolves(self):
        """Count file HAS a District column but it is blank -> master wins."""
        inv = [_count_row("Belleview", "alice", with_district_col=True),
               _count_row("Boulder", "bob", with_district_col=True)]
        out = self._rows(inv, master=MASTER)
        self.assertEqual(out.get("Belleview"), "Colorado")
        self.assertEqual(out.get("Boulder"), "Colorado")

    def test_count_file_district_wins_over_master(self):
        """Count file district is authoritative when present and non-empty."""
        inv = [_count_row("Belleview", "alice", with_district_col=True, district="Colorado West")]
        out = self._rows(inv, master=MASTER)
        self.assertEqual(out.get("Belleview"), "Colorado West")

    def test_count_file_unknown_literal_ignored(self):
        """A literal 'Unknown' in the count file must not stick."""
        inv = [_count_row("Belleview", "alice", with_district_col=True, district="Unknown")]
        out = self._rows(inv, master=MASTER)
        self.assertEqual(out.get("Belleview"), "Colorado")

    def test_master_only_pending_store_appears_with_district(self):
        """Store not in the count file but in Store List still gets its district."""
        inv = [_count_row("Belleview", "alice")]
        out = self._rows(inv, master=MASTER)
        self.assertEqual(out.get("Fuqua"), "Houston")

    def test_no_master_no_crash(self):
        inv = [_count_row("Belleview", "alice")]
        out = self._rows(inv, master=None)
        self.assertIn("Belleview", out)

    def test_status_still_correct(self):
        inv = [_count_row("Belleview", "alice")]
        rows, _summary = self.build_inventory_status_rows(inv, [], master_store_records=MASTER)
        by_store = {r.store: r for r in rows}
        self.assertEqual(by_store["Belleview"].status, "Completed")
        self.assertEqual(by_store["Fuqua"].status, "Pending")


if __name__ == "__main__":
    unittest.main()
