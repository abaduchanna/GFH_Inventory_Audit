"""Regression tests: audit table row tags must follow the active theme.

User report: in the GFH audit timesheet app, the Inventory Audit Status and
Variance Audit tabs rendered populated rows with LIGHT pastel backgrounds
(yellow/blue/green) while the app was in dark theme. Root cause: the row
tag backgrounds were hardcoded light in _build_status_tab/_build_audit_tab
and never re-applied on theme toggle (only "status_completed" was).

These tests pin the fix for BOTH shipped apps in this repo
(GFH_Inventory_Audit_Timesheet.py and GFH_Inventory_Audit.py): a single
theme-aware _apply_row_tag_colors() that configures all seven tags, is
called from both tab builders AND from _apply_theme(), and contains no
hardcoded light values outside the light theme palette.
"""
import ast
import os
import re

import pytest

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

MONOLITHS = (
    os.path.join(ROOT, "GFH_Inventory_Audit_Timesheet.py"),
    os.path.join(ROOT, "GFH_Inventory_Audit.py"),
)

# Light-theme pastels are allowed ONLY inside the light palette branch of
# _apply_row_tag_colors() — never in tag_configure(...) calls.
PASTELS = ("#FFF3CD", "#D7ECFF", "#D9F7DF")

# Dark-theme tints (yellow/blue/green families with light text)
DARK_TINTS = ("#3a3117", "#ffd66e", "#16324f", "#9ecbff", "#17331f", "#8fe3a8")

ALL_TAGS = (
    "status_pending", "status_completed_after_update", "status_completed_sent",
    "status_completed", "variance_pending", "variance_sent", "variance_cleared",
)


def _read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


@pytest.mark.parametrize("src_path", MONOLITHS, ids=os.path.basename)
def test_row_tag_colors_method_exists(src_path):
    src = _read(src_path)
    assert "def _apply_row_tag_colors(self" in src, \
        "_apply_row_tag_colors() helper missing"


@pytest.mark.parametrize("src_path", MONOLITHS, ids=os.path.basename)
def test_no_hardcoded_pastel_tag_configures(src_path):
    src = _read(src_path)
    for line in src.splitlines():
        if "tag_configure(" in line and any(p.lower() in line.lower() for p in PASTELS):
            raise AssertionError(
                f"hardcoded pastel row background in tag_configure: {line.strip()}")


@pytest.mark.parametrize("src_path", MONOLITHS, ids=os.path.basename)
def test_apply_theme_reapplies_row_tags(src_path):
    src = _read(src_path)
    m = re.search(r"def _apply_theme\(self.*?(?=\n    def )", src, re.S)
    assert m, "_apply_theme not found"
    assert "_apply_row_tag_colors()" in m.group(0), \
        "_apply_theme must re-apply row tag colors on every toggle"


@pytest.mark.parametrize("src_path", MONOLITHS, ids=os.path.basename)
def test_both_tab_builders_use_row_tag_colors(src_path):
    src = _read(src_path)
    for fn in ("_build_status_tab", "_build_audit_tab"):
        m = re.search(rf"def {fn}\(self.*?(?=\n    def )", src, re.S)
        assert m, f"{fn} not found"
        assert "_apply_row_tag_colors()" in m.group(0), \
            f"{fn} must apply theme-aware row tags (no hardcoded pastels)"


@pytest.mark.parametrize("src_path", MONOLITHS, ids=os.path.basename)
def test_dark_tints_and_all_tags_configured(src_path):
    src = _read(src_path)
    m = re.search(r"def _apply_row_tag_colors\(self.*?(?=\n    def )", src, re.S)
    body = m.group(0)
    for tint in DARK_TINTS:
        assert tint.lower() in body.lower(), f"dark tint {tint} missing"
    for tag in ALL_TAGS:
        assert f'"{tag}"' in body, f"tag {tag} not configured in _apply_row_tag_colors"


@pytest.mark.parametrize("src_path", MONOLITHS, ids=os.path.basename)
def test_monolith_still_parses(src_path):
    ast.parse(_read(src_path))
