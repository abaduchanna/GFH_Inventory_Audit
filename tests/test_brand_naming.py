"""Brand naming regression tests.

Static string checks - no tkinter/display needed, CI-safe. The Verge-style
band texture was REVERTED (it belongs to the Verge apps only); these tests
lock in the branding changes that were actually requested:

* settings dir prefix is GFH-Telecom (prefix in caps)
* ThemeManager app_name is VidaPay-GFH (no lowercase vidapay-gfh)
* header_manager keeps the lazy-tkinter import in add_copyright that
  prevents the frozen-exe NameError
"""
import pathlib
import re

ROOT = pathlib.Path(__file__).resolve().parents[1]

THEME_MANAGER = ROOT / "theme_manager.py"
HEADER_MANAGER = ROOT / "header_manager.py"
APP_FILES = [
    ROOT / "GFH_Inventory_Audit.py",
    ROOT / "GFH_Inventory_Audit_Timesheet.py",
]


def _source(path):
    assert path.exists(), f"missing expected source file: {path}"
    return path.read_text(encoding="utf-8")


def test_config_dir_uses_caps_prefix():
    src = _source(THEME_MANAGER)
    assert '"GFH-Telecom"' in src
    assert '"gfh-telecom"' not in src


def test_band_texture_reverted():
    # The Verge texture must NOT be wired into the GFH apps.
    for path in (THEME_MANAGER, HEADER_MANAGER, *APP_FILES):
        src = _source(path)
        assert "draw_band_texture" not in src, f"{path.name} still references the band texture"


def test_header_manager_keeps_lazy_tkinter_import_in_add_copyright():
    src = _source(HEADER_MANAGER)
    assert re.search(
        r"def add_copyright\(.*?\)\s*:\s*\n(?:.*\n)*?\s*import tkinter as tk",
        src,
    ), "add_copyright must import tkinter lazily before using tk"


def test_app_uses_vidapay_gfh_app_name():
    for app_path in APP_FILES:
        src = _source(app_path)
        assert 'app_name="VidaPay-GFH"' in src
        assert "vidapay-gfh" not in src
