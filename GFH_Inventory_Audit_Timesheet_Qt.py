"""
GFH Telecom LLC Inventory Audit — Timesheet Edition (PySide6/Qt build)
=======================================================================

Same tool as the Tkinter build, same data folder, same SQLite log, same
WhatsApp Desktop automation - a PySide6 UI layer on the same business
logic. See the Tkinter build's module docstring for the full feature
description.

What this app does
1. Upload Inventory_Count_Result_Details.xlsx.
2. Rebuilds the GFH Inventory Status / Inventory Count Results workflow:
   - Reads serial-level inventory count results.
   - Filters Inventory_Count_Result_Details to the latest Created Date/Time per Store + Created By.
   - Deduplicates by (Store, IMEI) keeping the latest record.
   - Treats non-matched serialized rows as variances.
   - Pulls District from the saved Store List. Rep/employee name comes straight from the
     "Created By" column in Inventory_Count_Result_Details (no separate time sheet file).
   - Merges Arizona - D1, Arizona - D2, and Arizona into Arizona.
   - Skips SIM, eSIM, SIM Card, SIM Kit, and similar SIM products from UI and sending.
3. Shows variances in a GUI with District, Store, Product, IMEI, Status, Rep Name, and Checkbox.
4. Checkbox marks the variance as cleared/resolved. Cleared rows are skipped in pending auto-send.
5. Sends selected or pending variances as PNG images to WhatsApp Desktop.
6. Send mode options:
   - District: one image batch per district.
   - Store: one image batch per store, routed to that store's district WhatsApp group.
   - Sales Rep: one image batch per sales rep per district, routed to that district WhatsApp group.
7. Opens with a blank UI and loads data only after the file path is provided and Load Variances is clicked.
9. Keeps a SQLite log of cleared, pending, and sent variance status.

Store list and employee phone numbers are stored in the local SQLite database. Store list and employees can be imported from XLSX. WhatsApp group names are saved in the database and configurable per district. WhatsApp sends captions together with images. Final district result sending is available per district. Starting message and 3-reminder sending for uncleared variances is supported.

WhatsApp Desktop notes:
- Keep WhatsApp Desktop installed and logged in before sending.
- Default search shortcut is Ctrl+F. If your WhatsApp uses Ctrl+K for chat search,
  change SEARCH_SHORTCUT below to "ctrl+k".
- District group routing uses the WhatsApp group names saved in the District DMs tab.

Developed by www.3SVerse.com
"""
from __future__ import annotations

import sys
import os
import csv
import datetime as dt
import hashlib
import re
import shutil
import sqlite3
import struct
import subprocess
import tempfile
import threading
import time
import traceback
import zipfile
import zlib
from datetime import date
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import xml.etree.ElementTree as ET
import base64

# ── Required package availability ──────────────────────────────────────────
# The Tkinter build auto-installed missing packages behind a Tk splash
# (see _auto_install_packages there). This Qt build follows the
# GFH_Inventory_Aging_Processor_Qt pattern instead: packages are probed at
# import time and any shortfall is reported by _check_required_packages()
# in main(), AFTER the QApplication exists so a QMessageBox can be shown.
_MISSING: List[str] = []

try:
    import openpyxl
except Exception:
    _MISSING.append("openpyxl")

try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk
except Exception:
    Image = None
    ImageDraw = None
    ImageFont = None
    ImageTk = None
    _MISSING.append("pillow")

try:
    # Presence checks only - no import. pyautogui/win32api/pygetwindow are
    # imported lazily inside the WhatsApp sender itself, same as the
    # Tkinter build.
    import importlib.util as _ilu
    if _ilu.find_spec("pyautogui") is None:
        _MISSING.append("pyautogui")
    if _ilu.find_spec("win32api") is None:
        _MISSING.append("pywin32")
    if _ilu.find_spec("pygetwindow") is None:
        _MISSING.append("pygetwindow")
except Exception:
    pass

from theme_manager_qt import ThemeManagerQt


def _check_required_packages(app) -> None:
    """Call once a QApplication exists. Mirrors the Tkinter build's
    auto-install list; the Qt build shows a dialog instead of a Tk splash
    when something is missing."""
    from PySide6.QtWidgets import QMessageBox
    if _MISSING:
        QMessageBox.critical(
            None, "GFH Telecom LLC Inventory Audit - missing packages",
            "Required packages not installed:\n\n"
            + "\n".join(f"  - {p}" for p in _MISSING)
            + "\n\nInstall and reopen:\n\npip install " + " ".join(_MISSING),
        )
        sys.exit(1)


APP_NAME = "GFH Telecom LLC Inventory Audit"
EMBEDDED_LOGO_B64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "embedded_logo_b64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "embedded_logo_b64.txt"), "r").read().strip()
EMBEDDED_ICON_B64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "embedded_icon_b64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "embedded_icon_b64.txt"), "r").read().strip()

if getattr(sys, "frozen", False):
    PACKAGE_DIR = Path(sys.executable).resolve().parent
else:
    PACKAGE_DIR = Path(__file__).resolve().parent
PORTABLE_APP_DIR = PACKAGE_DIR / "GFH_Inventory_Audit_Data"
LEGACY_APP_DIR = Path.home() / "GFH_Inventory_Variance_GUI"

def _choose_app_dir() -> Path:
    """Always place data folder next to the EXE (or script), regardless of location."""
    candidate = PORTABLE_APP_DIR
    try:
        candidate.mkdir(parents=True, exist_ok=True)
        test_file = candidate / ".write_test"
        test_file.write_text("ok", encoding="utf-8")
        try:
            test_file.unlink()
        except Exception:
            pass
        return candidate
    except Exception as e:
        # Qt port of the Tkinter build's messagebox: a QApplication usually
        # does not exist yet at import time, so fall back to stderr.
        try:
            from PySide6.QtWidgets import QApplication, QMessageBox
            _app = QApplication.instance()
            if _app is not None:
                QMessageBox.critical(
                    None,
                    "Folder not writable",
                    f"Cannot write to the data folder:\n"
                    f"  {candidate}\n\n"
                    f"Error: {e}\n\n"
                    f"Please ensure the folder is writable.",
                )
            else:
                raise RuntimeError("no QApplication")
        except Exception:
            print(
                f"Cannot write to the data folder:\n  {candidate}\n\nError: {e}\n\n"
                f"Please ensure the folder is writable.",
                file=sys.stderr,
            )
        raise SystemExit(1)

APP_DIR = _choose_app_dir()
DB_PATH = APP_DIR / "inventory_variance_log.sqlite3"
IMAGE_DIR = APP_DIR / "whatsapp_images"
EXPORT_DIR = APP_DIR / "exports"
STORE_CONFIG_PATH = APP_DIR / "store_master_list.csv"
PACKAGED_STORE_LIST_PATH = PACKAGE_DIR / "store_master_list.csv"


def _resource_path(name: str) -> Path:
    """Resolve a bundled asset for a PyInstaller onefile EXE.

    onefile builds extract data files to sys._MEIPASS at runtime; fall back
    to the directory next to the app/script for normal runs."""
    if getattr(sys, "frozen", False):
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            bundled = Path(meipass) / name
            if bundled.exists():
                return bundled
    return PACKAGE_DIR / name


HEADER_LOGO_PATH = _resource_path("GFH_Telecom_Logo.png")
STATUS_LOGO_PATH = HEADER_LOGO_PATH
APP_ICON_PATH = PACKAGE_DIR / "gfh_telecom_llc_icon.ico"
APP_ICON_PNG_PATH = PACKAGE_DIR / "gfh_telecom_llc_logo.png"

# ---------------------------------------------------------------------------
# Embedded icon assets — written to PACKAGE_DIR at startup so no external
# file is needed. Swap the base64 strings to change the window icon.
# ---------------------------------------------------------------------------
APP_ICON_ICO_BASE64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "app_icon_ico_base64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "app_icon_ico_base64.txt"), "r").read().strip()

APP_ICON_PNG_BASE64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "app_icon_png_base64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "app_icon_png_base64.txt"), "r").read().strip()


def ensure_app_icon_files() -> None:
    """Decode and write embedded ICO + PNG icons to PACKAGE_DIR once."""
    try:
        ico_bytes = base64.b64decode(APP_ICON_ICO_BASE64.strip().replace("\n", ""))
        if not APP_ICON_PATH.exists() or APP_ICON_PATH.stat().st_size != len(ico_bytes):
            APP_ICON_PATH.write_bytes(ico_bytes)
    except Exception:
        pass
    try:
        png_bytes = base64.b64decode(APP_ICON_PNG_BASE64.strip().replace("\n", ""))
        if not APP_ICON_PNG_PATH.exists() or APP_ICON_PNG_PATH.stat().st_size != len(png_bytes):
            APP_ICON_PNG_PATH.write_bytes(png_bytes)
    except Exception:
        pass


LOGO_PNG_NAME = "GFH_Telecom_Logo.png"   # used in the header (FixedHeaderManagerQt)
ICON_ICO_NAME = "GFH_icon.ico"           # taskbar + titlebar icon (Windows .ico)
SEARCH_SHORTCUT = "ctrl+f"
SEND_ONLY_UNSENT_BY_DEFAULT = True
AUTO_SEND_CHUNK_SIZE = 22

EXCEL_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

SEND_MODE_LABELS = {
    "district": "District",
    "store": "Store",
    "rep": "Sales Rep",
}



DEFAULT_STORE_MASTER = []

def default_store_master_records() -> List[Dict[str, str]]:
    return [{"District": district, "Store": store} for district, store in DEFAULT_STORE_MASTER]


def migrate_legacy_app_data_if_needed() -> None:
    """No-op: legacy per-user folder migration removed.
    All instances now share the single portable DB next to the script."""
    pass


def save_store_master_records(records: List[Dict[str, str]], path: Path = STORE_CONFIG_PATH) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    normalized_rows: Dict[str, Dict[str, str]] = {}
    for rec in records:
        district = normalize_district(rec.get("District", ""))
        store = display_store(rec.get("Store", ""))
        norm = normalize_store(store)
        if not norm or not district or district == "Unknown":
            continue
        normalized_rows[norm] = {"District": district, "Store": store}
    ordered = sorted(normalized_rows.values(), key=lambda r: (normalize_district(r["District"]), r["Store"].lower()))
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=["District", "Store"])
        writer.writeheader()
        writer.writerows(ordered)


def ensure_store_master_file() -> None:
    if STORE_CONFIG_PATH.exists():
        return
    STORE_CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    if PACKAGED_STORE_LIST_PATH.exists():
        shutil.copy2(PACKAGED_STORE_LIST_PATH, STORE_CONFIG_PATH)
    else:
        save_store_master_records(default_store_master_records(), STORE_CONFIG_PATH)


def load_store_master_records(path: Path = STORE_CONFIG_PATH) -> List[Dict[str, str]]:
    ensure_store_master_file()
    rows: List[Dict[str, str]] = []
    try:
        with path.open("r", newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for rec in reader:
                district = normalize_district(rec.get("District", ""))
                store = display_store(rec.get("Store", ""))
                if district and district != "Unknown" and store:
                    rows.append({"District": district, "Store": store})
    except Exception:
        rows = default_store_master_records()
        save_store_master_records(rows, path)
    if not rows:
        rows = default_store_master_records()
        save_store_master_records(rows, path)
    return rows


@dataclass
class VarianceRow:
    key: str
    district: str
    store: str
    product: str
    imei: str
    status: str
    created_by: str = ""
    rep_name: str = ""
    created_date: str = ""
    document_status: str = ""
    source_file: str = ""
    cleared: bool = False
    sent_count: int = 0
    last_sent_at: str = ""
    cleared_at: str = ""
    notes: str = ""


@dataclass
class InventoryStatusRow:
    key: str
    district: str
    store: str
    status: str
    rep_name: str = ""
    source_file: str = ""

def now_text() -> str:
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def normalize_header(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", safe_text(text).lower())


def normalize_store(text: str) -> str:
    value = safe_text(text)
    value = re.sub(r"\s+", " ", value).strip()
    value = value.replace("–", "-").replace("—", "-")
    return value.lower()


def display_store(text: str) -> str:
    value = safe_text(text)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def normalize_phone(text: str) -> str:
    value = safe_text(text)
    if not value:
        return ""
    value = value.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
    value = value.replace(".", "")
    if value.startswith("00"):
        value = "+" + value[2:]
    return value


def whatsapp_mention(phone: str) -> str:
    phone = normalize_phone(phone)
    if not phone:
        return ""
    if phone.startswith("@"):
        return phone
    return "@" + phone


def person_name_key(text: str) -> str:
    value = safe_text(text).lower()
    value = value.replace(",", " ")
    value = re.sub(r"[^a-z0-9]+", " ", value)
    parts = [p for p in value.split() if p]
    return " ".join(sorted(parts))


def device_rule_key(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", safe_text(text).lower())


def device_matches_rule(product: str, imei: str, rule_text: str, match_type: str) -> bool:
    rule_clean = device_rule_key(rule_text)
    if not rule_clean:
        return False

    product_clean = device_rule_key(product)
    imei_clean = device_rule_key(imei)
    match_type_clean = normalize_header(match_type)

    if match_type_clean in {"productexact", "exactproduct"}:
        return product_clean == rule_clean
    if match_type_clean in {"imeiexact", "serialexact", "esnexact"}:
        return imei_clean == rule_clean
    if match_type_clean in {"anycontains", "containsany"}:
        return rule_clean in product_clean or rule_clean in imei_clean
    return rule_clean in product_clean


def normalize_district(text: str) -> str:
    value = safe_text(text).replace("–", "-").replace("—", "-")
    value = re.sub(r"\s+", " ", value).strip()
    if not value:
        return "Unknown"
    lower = value.lower()
    if lower.startswith("arizona"):
        return "Arizona"
    aliases = {
        "atlanta": "Atlanta",
        "colorado west": "Colorado West",
        "houston": "Houston",
        "colorado east": "Colorado East",
        "tennessee": "Tennessee",
        "louisiana": "Louisiana",
    }
    return aliases.get(lower, value)


def group_name_for_district(district: str, db=None) -> str:
    normalized = normalize_district(district)
    if db is not None:
        try:
            saved = db.find_whatsapp_group(normalized)
            if saved:
                return saved
        except Exception:
            pass
    return f"GFH TELECOM {normalized.upper()}"


def is_sim_product(product: str) -> bool:
    text = safe_text(product).lower()
    compact = normalize_header(product)
    if not text and not compact:
        return False
    if re.search(r"(^|[^a-z0-9])e?[-\s]?sim(s)?([^a-z0-9]|$)", text):
        return True
    if "simcard" in compact or compact in {"sim", "sims", "esim", "esims"}:
        return True
    if compact.startswith(("sim", "esim")) and any(token in compact for token in ("card", "kit", "pack", "starter")):
        return True
    return False


def excel_serial_to_datetime(value) -> Optional[dt.datetime]:
    text = safe_text(value)
    if not text:
        return None
    try:
        serial = float(text)
    except ValueError:
        # Try common date strings if present.
        for fmt in ("%m/%d/%Y", "%m/%d/%Y %I:%M %p", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return dt.datetime.strptime(text, fmt)
            except ValueError:
                continue
        return None
    base = dt.datetime(1899, 12, 30)
    return base + dt.timedelta(days=serial)


def excel_serial_to_date_text(value) -> str:
    parsed = excel_serial_to_datetime(value)
    if parsed is None:
        return safe_text(value)
    return parsed.strftime("%m/%d/%Y %I:%M %p")


def numeric_excel_date(value) -> float:
    text = safe_text(value)
    try:
        return float(text)
    except ValueError:
        parsed = excel_serial_to_datetime(text)
        if parsed is None:
            return -1.0
        base = dt.datetime(1899, 12, 30)
        return (parsed - base).total_seconds() / 86400.0


def variance_key(store: str, imei: str, product: str, status: str, created_by: str = "", created_date: str = "") -> str:
    raw = "|".join([
        normalize_store(store),
        safe_text(imei).lower(),
        safe_text(product).lower(),
        safe_text(status).lower(),
        safe_text(created_by).lower(),
        safe_text(created_date).lower(),
    ])
    return hashlib.sha1(raw.encode("utf-8", errors="ignore")).hexdigest()



# ── win32api keyboard fallback (for Python 3.14 where pynput breaks) ────────
_WA_VK_MAP = {
    "ctrl": 0x11, "control": 0x11, "alt": 0x12, "menu": 0x12,
    "shift": 0x10, "win": 0x5B, "enter": 0x0D, "return": 0x0D,
    "tab": 0x09, "esc": 0x1B, "escape": 0x1B, "backspace": 0x08,
    "delete": 0x2E, "del": 0x2E, "home": 0x24, "end": 0x23,
    "pageup": 0x21, "pagedown": 0x22, "up": 0x26, "down": 0x28,
    "left": 0x25, "right": 0x27, "space": 0x20, "a": 0x41, "c": 0x43,
    "f": 0x46, "v": 0x56, "x": 0x58,
}

def _wa_win32_key_down(vk_code):
    import ctypes
    ctypes.windll.user32.keybd_event(vk_code, 0, 0, 0)

def _wa_win32_key_up(vk_code):
    import ctypes
    ctypes.windll.user32.keybd_event(vk_code, 0, 0x0002, 0)

def _wa_win32_hotkey(*keys):
    codes = []
    for k in keys:
        k_lower = k.lower()
        if k_lower in _WA_VK_MAP:
            codes.append(_WA_VK_MAP[k_lower])
        elif len(k) == 1:
            codes.append(ord(k.upper()))
        else:
            return
    for code in codes:
        _wa_win32_key_down(code)
    for code in reversed(codes):
        _wa_win32_key_up(code)

def _wa_win32_press(key):
    k_lower = key.lower()
    if k_lower in _WA_VK_MAP:
        code = _WA_VK_MAP[k_lower]
    elif len(key) == 1:
        code = ord(key.upper())
    else:
        return
    _wa_win32_key_down(code)
    _wa_win32_key_up(code)

def _wa_win32_write(text, interval=0.0):
    import ctypes
    for char in text:
        vk = ctypes.windll.user32.VkKeyScanW(ord(char))
        if vk == -1:
            continue
        code = vk & 0xFF
        shift = (vk >> 8) & 1
        if shift:
            _wa_win32_key_down(_WA_VK_MAP["shift"])
        _wa_win32_key_down(code)
        _wa_win32_key_up(code)
        if shift:
            _wa_win32_key_up(_WA_VK_MAP["shift"])
        if interval > 0:
            import time as _t
            _t.sleep(interval)

_wa_hotkey = _wa_win32_hotkey
_wa_press = _wa_win32_press
_wa_write = _wa_win32_write


class RobustXlsxReader:
    """Small XLSX reader with a fallback for exports missing a central ZIP directory."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        self.parts = self._read_parts()
        self.shared_strings = self._read_shared_strings()

    def _read_parts(self) -> Dict[str, bytes]:
        parts: Dict[str, bytes] = {}
        try:
            with zipfile.ZipFile(self.path) as zf:
                for name in zf.namelist():
                    parts[name] = zf.read(name)
            return parts
        except Exception:
            data = self.path.read_bytes()
            pos = 0
            while pos + 30 <= len(data) and data[pos:pos + 4] == b"PK\x03\x04":
                try:
                    (
                        _sig,
                        _ver,
                        _flag,
                        method,
                        _mtime,
                        _mdate,
                        _crc,
                        csize,
                        _usize,
                        nlen,
                        xlen,
                    ) = struct.unpack_from("<IHHHHHIIIHH", data, pos)
                    name_start = pos + 30
                    name_end = name_start + nlen
                    name = data[name_start:name_end].decode("utf-8", errors="replace")
                    start = name_end + xlen
                    comp = data[start:min(start + csize, len(data))]
                    if method == 8:
                        try:
                            content = zlib.decompress(comp, -15)
                        except Exception:
                            content = b""
                    elif method == 0:
                        content = comp
                    else:
                        content = b""
                    if content:
                        parts[name] = content
                    pos = start + csize
                except Exception:
                    break
            if not parts:
                raise RuntimeError(f"Could not read Excel file: {self.path}")
            return parts

    @staticmethod
    def _clean_xml(raw: bytes) -> bytes:
        return raw.lstrip(b"\xef\xbb\xbf")

    def _read_shared_strings(self) -> List[str]:
        raw = self.parts.get("xl/sharedStrings.xml")
        if not raw:
            return []
        root = ET.fromstring(self._clean_xml(raw))
        strings: List[str] = []
        for si in root.findall(EXCEL_NS + "si"):
            strings.append("".join((t.text or "") for t in si.iter(EXCEL_NS + "t")))
        return strings

    def _workbook_rels(self) -> Dict[str, str]:
        raw = self.parts.get("xl/_rels/workbook.xml.rels")
        if not raw:
            return {}
        root = ET.fromstring(self._clean_xml(raw))
        rels: Dict[str, str] = {}
        for rel in root.findall(REL_NS + "Relationship"):
            rid = rel.attrib.get("Id", "")
            target = rel.attrib.get("Target", "")
            if not target:
                continue
            if not target.startswith("/"):
                target = "xl/" + target
            else:
                target = target.lstrip("/")
            rels[rid] = target
        return rels

    def sheet_paths(self) -> List[Tuple[str, str]]:
        raw = self.parts.get("xl/workbook.xml")
        if not raw:
            if "xl/worksheets/sheet1.xml" in self.parts:
                return [("Sheet1", "xl/worksheets/sheet1.xml")]
            return []
        root = ET.fromstring(self._clean_xml(raw))
        rels = self._workbook_rels()
        sheets: List[Tuple[str, str]] = []
        for sheet in root.findall(".//" + EXCEL_NS + "sheet"):
            name = sheet.attrib.get("name", "Sheet")
            rid = sheet.attrib.get(R_NS + "id", "")
            path = rels.get(rid)
            if path and path in self.parts:
                sheets.append((name, path))
        if not sheets and "xl/worksheets/sheet1.xml" in self.parts:
            sheets.append(("Sheet1", "xl/worksheets/sheet1.xml"))
        return sheets

    @staticmethod
    def _cell_to_col(ref: str) -> Optional[int]:
        match = re.match(r"([A-Z]+)(\d+)", ref or "")
        if not match:
            return None
        col = 0
        for ch in match.group(1):
            col = col * 26 + ord(ch) - 64
        return col

    def _cell_value(self, cell: ET.Element) -> str:
        cell_type = cell.attrib.get("t")
        value_node = cell.find(EXCEL_NS + "v")
        if cell_type == "s":
            if value_node is None or value_node.text is None:
                return ""
            try:
                idx = int(value_node.text)
            except ValueError:
                return ""
            if 0 <= idx < len(self.shared_strings):
                return self.shared_strings[idx]
            return ""
        if cell_type == "inlineStr":
            inline = cell.find(EXCEL_NS + "is")
            if inline is None:
                return ""
            return "".join((t.text or "") for t in inline.iter(EXCEL_NS + "t"))
        if value_node is None or value_node.text is None:
            return ""
        return value_node.text

    def read_sheet(self, preferred_name: Optional[str] = None) -> List[Dict[str, str]]:
        sheets = self.sheet_paths()
        if not sheets:
            raise RuntimeError(f"No worksheet found in {self.path.name}")
        selected_name, selected_path = sheets[0]
        if preferred_name:
            want = preferred_name.lower().strip()
            for name, path in sheets:
                if want in name.lower().strip():
                    selected_name, selected_path = name, path
                    break
        raw = self.parts[selected_path]
        root = ET.fromstring(self._clean_xml(raw))
        rows: List[Tuple[int, Dict[int, str]]] = []
        for row in root.findall(".//" + EXCEL_NS + "row"):
            row_num = int(row.attrib.get("r", "0") or 0)
            values: Dict[int, str] = {}
            for cell in row.findall(EXCEL_NS + "c"):
                col = self._cell_to_col(cell.attrib.get("r", ""))
                if not col:
                    continue
                values[col] = self._cell_value(cell)
            if values:
                rows.append((row_num, values))
        if not rows:
            return []

        header_row_idx, header_values = rows[0]
        max_col = max(header_values)
        headers = [safe_text(header_values.get(i, "")) for i in range(1, max_col + 1)]
        seen: Dict[str, int] = {}
        clean_headers = []
        for i, header in enumerate(headers, start=1):
            name = header if header else f"Column{i}"
            base = name
            if base in seen:
                seen[base] += 1
                name = f"{base}_{seen[base]}"
            else:
                seen[base] = 1
            clean_headers.append(name)

        records: List[Dict[str, str]] = []
        for row_num, value_map in rows[1:]:
            if row_num <= header_row_idx:
                continue
            record: Dict[str, str] = {}
            max_data_col = max(max_col, max(value_map.keys()) if value_map else max_col)
            for i in range(1, max_data_col + 1):
                header = clean_headers[i - 1] if i <= len(clean_headers) else f"Column{i}"
                record[header] = safe_text(value_map.get(i, ""))
            if any(safe_text(v) for v in record.values()):
                records.append(record)
        return records


def find_column(record: Dict[str, str], candidates: Iterable[str]) -> Optional[str]:
    normalized = {normalize_header(k): k for k in record.keys()}
    for candidate in candidates:
        key = normalize_header(candidate)
        if key in normalized:
            return normalized[key]
    for candidate in candidates:
        key = normalize_header(candidate)
        for norm, original in normalized.items():
            if key and (key in norm or norm in key):
                return original
    return None


def read_xlsx_records(path: str | Path) -> List[Dict[str, str]]:
    reader = RobustXlsxReader(path)
    return reader.read_sheet()


def build_ts_store_to_employee(time_sheet_records: List[Dict[str, str]]) -> Dict[str, str]:
    """Map normalized store → employee name from Timesheet rows.

    Keeps rows without a Clock In too — the employee may still be at the
    store with the count pending, and the pending-store rep lookup is
    built from these rows. When a store has several rows, the latest
    Clock In wins (falling back to row order when Clock In is missing).
    """
    ts_store_to_employee: Dict[str, str] = {}
    if not time_sheet_records:
        return ts_store_to_employee
    sample_ts    = time_sheet_records[0]
    ts_store_col = find_column(sample_ts, ["Store"])
    ts_emp_col   = find_column(sample_ts, ["Employee", "Employee Name",
                                           "Salesperson", "Sales Person", "Rep Name"])
    ts_clock_col = find_column(sample_ts, ["Clock In", "Clock-In", "Clock In "])
    latest_ts: Dict[str, float] = {}
    for idx, rec in enumerate(time_sheet_records):
        emp_raw = safe_text(rec.get(ts_emp_col, "")) if ts_emp_col else ""
        # Skip TOTAL/subtotal rows
        if not emp_raw or "TOTAL" in emp_raw.upper() or "\u2014 " in emp_raw or "-- " in emp_raw:
            continue
        store_raw = rec.get(ts_store_col, "") if ts_store_col else ""
        norm = normalize_store(store_raw)
        if not norm:
            continue
        clock_in_val = safe_text(rec.get(ts_clock_col, "")) if ts_clock_col else ""
        score = numeric_excel_date(clock_in_val) if clock_in_val.strip() else float(idx)
        if score >= latest_ts.get(norm, -1.0):
            ts_store_to_employee[norm] = emp_raw.strip()
            latest_ts[norm] = score
    return ts_store_to_employee


def store_name_tokens(norm_store: str) -> Tuple[str, ...]:
    """Alphanumeric tokens of an already-normalized store name, deduped."""
    return tuple(dict.fromkeys(t for t in re.findall(r"[a-z0-9]+", norm_store or "") if t))


def match_store_employee(norm_store: str, ts_store_to_employee: Dict[str, str]) -> str:
    """Resolve the timesheet employee for a store.

    Exact normalized-store lookup first, then a tolerant token-overlap
    match. The timesheet export and the count file often spell the same
    store slightly differently ("Kings Highway #1204" vs "Kings Highway
    Store 1204"), which left pending stores without an employee name even
    when the timesheet had a clock-in for them. A token match is accepted
    when one name's tokens are a subset of the other's — but never when a
    distinguishing store number conflicts ("Store 1204" vs "Store 1205"
    do not match). Ties go to the candidate with the most token overlap,
    then the shortest name.
    """
    if not norm_store or not ts_store_to_employee:
        return ""
    exact = ts_store_to_employee.get(norm_store)
    if exact:
        return exact
    target_tokens = set(store_name_tokens(norm_store))
    if not target_tokens:
        return ""
    best_name = ""
    best_overlap = 0
    best_len = -1
    for ts_norm, employee in ts_store_to_employee.items():
        if not ts_norm or not employee:
            continue
        ts_tokens = set(store_name_tokens(ts_norm))
        if not ts_tokens:
            continue
        if target_tokens <= ts_tokens or ts_tokens <= target_tokens:
            overlap = len(target_tokens & ts_tokens)
            if overlap > best_overlap or (overlap == best_overlap and best_len != -1 and len(ts_norm) < best_len):
                best_name = employee
                best_overlap = overlap
                best_len = len(ts_norm)
            elif best_len == -1:
                best_name = employee
                best_overlap = overlap
                best_len = len(ts_norm)
    return best_name


def build_store_maps(
    inventory_records: List[Dict[str, str]],
    time_sheet_records: List[Dict[str, str]],
) -> Tuple[Dict[str, str], Dict[str, str], Dict[str, str]]:
    """Build store lookups using Count Details as the source of truth.

    - Store names and Districts come from Inventory_Count_Result_Details.
    - Employee (rep) names come from the Timesheet by matching on Store,
      tolerating slight store-name differences between the two exports.
    - Timesheet is ONLY used to resolve who worked at each store.
    """
    district_by_store: Dict[str, str] = {}
    display_by_store:  Dict[str, str] = {}
    rep_by_store:      Dict[str, str] = {}   # norm_store → full Employee name

    # ── Step 1: Build store → employee map from Timesheet ─────────────────
    ts_store_to_employee = build_ts_store_to_employee(time_sheet_records)

    # ── Step 2: Build district/display maps from Count Details ────────────
    if not inventory_records:
        return district_by_store, display_by_store, rep_by_store

    sample_inv   = inventory_records[0]
    inv_store_col    = find_column(sample_inv, ["Store"])
    inv_district_col = find_column(sample_inv, ["District"])

    if not inv_store_col:
        return district_by_store, display_by_store, rep_by_store

    for rec in inventory_records:
        store_raw = rec.get(inv_store_col, "")
        norm = normalize_store(store_raw)
        if not norm:
            continue

        # Store display name from count file
        display_by_store[norm] = display_store(store_raw)

        # District from count file
        if inv_district_col:
            district = normalize_district(rec.get(inv_district_col, ""))
            if district and district != "Unknown":
                district_by_store[norm] = district

        # Employee name from timesheet (matched by store, tolerant of
        # slight name differences between the two exports)
        if norm not in rep_by_store:
            employee = ts_store_to_employee.get(norm) or match_store_employee(norm, ts_store_to_employee)
            if employee:
                rep_by_store[norm] = employee

    return district_by_store, display_by_store, rep_by_store


def build_rep_by_store_from_created_by(
    inventory_records: List[Dict[str, str]],
) -> Dict[str, str]:
    """Latest "Created By" value per store from the Inventory_Count_Result_Details rows.

    Replaces the old Employee_Time_Sheet-based rep lookup. Employee/rep name for a store is
    simply whoever most recently counted that store, per the Created Date/Time column.
    """
    rep_by_store: Dict[str, str] = {}
    latest_score_by_store: Dict[str, float] = {}
    if not inventory_records:
        return rep_by_store

    sample = inventory_records[0]
    store_col = find_column(sample, ["Store"])
    created_by_col = find_column(sample, ["Created By", "Count By", "User Login"])
    created_date_col = find_column(sample, ["Created Date", "Created Date/Time", "Date Time", "Date"])
    if not store_col or not created_by_col:
        return rep_by_store

    for index, rec in enumerate(inventory_records):
        norm = normalize_store(rec.get(store_col, ""))
        if not norm:
            continue
        created_by = safe_text(rec.get(created_by_col, ""))
        if not created_by:
            continue
        score = numeric_excel_date(rec.get(created_date_col, "")) if created_date_col else float(index)
        if score < 0:
            score = float(index) / 1000000.0
        if score >= latest_score_by_store.get(norm, -1.0):
            rep_by_store[norm] = created_by
            latest_score_by_store[norm] = score

    return rep_by_store


def filter_latest_inventory_records(inventory_records: List[Dict[str, str]]) -> Tuple[List[Dict[str, str]], Dict[str, int]]:
    """Keep only the latest Inventory_Count_Result_Details rows per Store + Created By.

    GFH exports can include more than one count from the same Created By user for the same store.
    Older rows should not drive the current variance UI because a later recount can convert earlier
    Deficit rows into Matched rows. The Created Date value includes date and time as an Excel serial.
    """
    metrics = {
        "raw_inventory_rows": len(inventory_records),
        "latest_inventory_rows": len(inventory_records),
        "stale_inventory_rows": 0,
        "latest_created_by_groups": 0,
    }
    if not inventory_records:
        return inventory_records, metrics

    sample = inventory_records[0]
    store_col = find_column(sample, ["Store"])
    created_by_col = find_column(sample, ["Created By", "Count By", "User Login"])
    created_date_col = find_column(sample, ["Created Date", "Created Date/Time", "Date Time", "Date"])

    if not store_col or not created_by_col or not created_date_col:
        return inventory_records, metrics

    latest_by_group: Dict[Tuple[str, str], float] = {}
    scores: List[Tuple[Tuple[str, str], float]] = []

    for index, rec in enumerate(inventory_records):
        store_key = normalize_store(rec.get(store_col, "")) or "__unknown_store__"
        created_by_key = safe_text(rec.get(created_by_col, "")).lower().strip() or "__unknown_created_by__"
        group_key = (store_key, created_by_key)
        score = numeric_excel_date(rec.get(created_date_col, ""))
        if score < 0:
            score = float(index) / 1000000.0
        scores.append((group_key, score))
        if group_key not in latest_by_group or score > latest_by_group[group_key]:
            latest_by_group[group_key] = score

    filtered: List[Dict[str, str]] = []
    for rec, (group_key, score) in zip(inventory_records, scores):
        latest_score = latest_by_group.get(group_key, score)
        if abs(score - latest_score) <= 0.0000001:
            filtered.append(rec)

    metrics["latest_inventory_rows"] = len(filtered)
    metrics["stale_inventory_rows"] = len(inventory_records) - len(filtered)
    metrics["latest_created_by_groups"] = len(latest_by_group)
    return filtered, metrics


def extract_variances(
    inventory_records: List[Dict[str, str]],
    time_sheet_records: List[Dict[str, str]],
    master_store_records: Optional[List[Dict[str, str]]] = None,
    source_file: str = "",
) -> Tuple[List[VarianceRow], Dict[str, int]]:
    if not inventory_records:
        return [], {"completed": 0, "pending": 0, "stores_total": 0, "skipped_sims": 0, "raw_inventory_rows": 0, "latest_inventory_rows": 0, "stale_inventory_rows": 0, "latest_created_by_groups": 0}

    # Employee (rep) names come from timesheet matched by store.
    # Store names and Districts come from count details (source of truth).
    _, _, rep_by_store = build_store_maps(inventory_records, time_sheet_records)

    # Build district/display maps from count details directly
    district_by_store: Dict[str, str] = {}
    display_by_store: Dict[str, str] = {}
    sample0 = inventory_records[0]
    _store_col0    = find_column(sample0, ["Store"])
    _district_col0 = find_column(sample0, ["District"])
    for rec in inventory_records:
        store_raw = rec.get(_store_col0, "") if _store_col0 else ""
        norm = normalize_store(store_raw)
        if not norm:
            continue
        display_by_store[norm] = display_store(store_raw)
        if _district_col0:
            d = normalize_district(rec.get(_district_col0, ""))
            if d and d != "Unknown":
                district_by_store[norm] = d

    # Fill gaps from master store list
    for rec in master_store_records or []:
        district = normalize_district(rec.get("District", ""))
        store = display_store(rec.get("Store", ""))
        norm = normalize_store(store)
        if not norm:
            continue
        if district and district != "Unknown" and norm not in district_by_store:
            district_by_store[norm] = district
        if store and norm not in display_by_store:
            display_by_store[norm] = store
    sample = inventory_records[0]
    store_col = find_column(sample, ["Store"])
    product_col = find_column(sample, ["Product Description", "Product"])
    imei_col = find_column(sample, ["Serial #", "Serial", "IMEI", "ESN"])
    status_col = find_column(sample, ["Status"])
    created_by_col = find_column(sample, ["Created By", "Count By", "User Login"])
    created_date_col = find_column(sample, ["Created Date", "Date"])
    document_status_col = find_column(sample, ["Document Status"])

    missing = [
        name for name, col in [
            ("Store", store_col),
            ("Product Description", product_col),
            ("Serial # / IMEI", imei_col),
            ("Status", status_col),
        ] if not col
    ]
    if missing:
        raise RuntimeError("Missing required inventory column(s): " + ", ".join(missing))

    inventory_records, latest_metrics = filter_latest_inventory_records(inventory_records)

    # Deduplicate by (Store, IMEI) to prevent duplicate IMEI rows
    dedup_map: Dict[Tuple[str, str], Tuple[Dict[str, str], float]] = {}
    for rec in inventory_records:
        store_raw = rec.get(store_col, "")
        imei = safe_text(rec.get(imei_col, ""))
        if not imei:
            continue
        norm_store = normalize_store(store_raw)
        key = (norm_store, imei.lower())
        date_score = numeric_excel_date(rec.get(created_date_col, "")) if created_date_col else 0
        if key not in dedup_map or date_score > dedup_map[key][1]:
            dedup_map[key] = (rec, date_score)
    inventory_records = [rec for rec, _ in dedup_map.values()]
    latest_metrics["dedup_inventory_rows"] = len(inventory_records)

    completed_store_norms = set()
    for rec in inventory_records:
        norm = normalize_store(rec.get(store_col, ""))
        if norm:
            completed_store_norms.add(norm)

    all_store_norms = set(display_by_store.keys()) | completed_store_norms
    pending_store_norms = all_store_norms - completed_store_norms

    variance_rows: List[VarianceRow] = []
    skipped_sims = 0

    for rec in inventory_records:
        status = safe_text(rec.get(status_col, ""))
        if not status:
            continue
        if normalize_header(status) in {"matched", "match", "ok", "balanced"}:
            continue

        store = display_store(rec.get(store_col, ""))
        product = safe_text(rec.get(product_col, ""))
        imei = safe_text(rec.get(imei_col, ""))

        if not imei:
            continue
        if product and normalize_header(product) in {"accessorycommission", "accessoriescommission"}:
            continue
        if is_sim_product(product):
            skipped_sims += 1
            continue

        norm_store = normalize_store(store)
        district = normalize_district(district_by_store.get(norm_store, "Unknown"))
        created_by = safe_text(rec.get(created_by_col, "")) if created_by_col else ""
        # Rep/employee name now comes straight from the "Created By" column of the
        # inventory count Excel — no more separate Employee_Time_Sheet lookup.
        rep_name = created_by
        created_date = excel_serial_to_date_text(rec.get(created_date_col, "")) if created_date_col else ""
        document_status = safe_text(rec.get(document_status_col, "")) if document_status_col else ""
        key = variance_key(store, imei, product, status, created_by, created_date)

        variance_rows.append(
            VarianceRow(
                key=key,
                district=district,
                store=store,
                product=product,
                imei=imei,
                status=status,
                created_by=created_by,
                rep_name=rep_name,
                created_date=created_date,
                document_status=document_status,
                source_file=source_file,
            )
        )

    summary = {
        "completed": len(completed_store_norms),
        "pending": len(pending_store_norms),
        "stores_total": len(all_store_norms),
        "skipped_sims": skipped_sims,
        **latest_metrics,
    }
    return variance_rows, summary




def build_inventory_status_rows(
    inventory_records: List[Dict[str, str]],
    time_sheet_records: List[Dict[str, str]],
    master_store_records: Optional[List[Dict[str, str]]] = None,
    source_file: str = "",
) -> Tuple[List[InventoryStatusRow], Dict[str, int]]:
    # build_store_maps gives us employee (rep) names matched by store.
    # It does NOT drive which stores appear — count details is the source of truth.
    _, _, rep_by_store = build_store_maps(inventory_records, time_sheet_records)
    # Timesheet store → employee map used as a fallback for pending stores
    # that never appear in the count file (master-list-only stores): those
    # stores are missing from rep_by_store because build_store_maps only
    # transfers employees for count-file stores, yet the timesheet clearly
    # shows who clocked in there.
    ts_store_to_employee = build_ts_store_to_employee(time_sheet_records)
    # Timesheet district fallback — third layer after the count file and the
    # Store List (master) so a store known only to the timesheet still gets
    # its district instead of "Unknown".
    ts_district_by_store: Dict[str, str] = {}
    if time_sheet_records:
        ts_sample = time_sheet_records[0]
        ts_store_col = find_column(ts_sample, ["Store"])
        ts_district_col = find_column(ts_sample, ["District"])
        if ts_store_col and ts_district_col:
            for rec in time_sheet_records:
                norm = normalize_store(rec.get(ts_store_col, ""))
                if not norm:
                    continue
                d = normalize_district(rec.get(ts_district_col, ""))
                if d and d != "Unknown":
                    ts_district_by_store.setdefault(norm, d)

    # Build store display names AND districts directly from count details so
    # every store in the count file appears regardless of timesheet coverage.
    count_display_by_store: Dict[str, str] = {}
    count_district_by_store: Dict[str, str] = {}
    completed_store_norms: set[str] = set()
    inv_display_by_store: Dict[str, str] = {}

    latest_metrics = {
        "raw_inventory_rows": len(inventory_records),
        "latest_inventory_rows": len(inventory_records),
        "stale_inventory_rows": 0,
        "latest_created_by_groups": 0,
    }

    if inventory_records:
        sample = inventory_records[0]
        store_col    = find_column(sample, ["Store"])
        district_col = find_column(sample, ["District"])

        # First pass: collect ALL stores+districts from the count file
        for rec in inventory_records:
            store_raw = rec.get(store_col, "") if store_col else ""
            norm = normalize_store(store_raw)
            if not norm:
                continue
            count_display_by_store[norm] = display_store(store_raw)
            if district_col:
                d = normalize_district(rec.get(district_col, ""))
                if d and d != "Unknown":
                    count_district_by_store[norm] = d

        # Second pass: find completed stores (those that have records in latest batch)
        if store_col:
            filtered_records, latest_metrics = filter_latest_inventory_records(inventory_records)
            for rec in filtered_records:
                store_raw = rec.get(store_col, "")
                norm = normalize_store(store_raw)
                if norm:
                    inv_display_by_store[norm] = display_store(store_raw)
                    completed_store_norms.add(norm)

    # Fallback: also include master store list for any stores not in count file
    master_display_by_store: Dict[str, str] = {}
    master_district_by_store: Dict[str, str] = {}
    for rec in master_store_records or []:
        store = display_store(rec.get("Store", ""))
        norm = normalize_store(store)
        if not norm:
            continue
        master_display_by_store[norm] = store
        d = normalize_district(rec.get("District", ""))
        if d and d != "Unknown":
            master_district_by_store[norm] = d

    # All stores = everything from count file (primary) + master list (fallback)
    all_store_norms = sorted(
        set(count_display_by_store.keys())
        | set(master_display_by_store.keys())
    )

    rows: List[InventoryStatusRow] = []
    for norm in all_store_norms:
        # Display name: count file first, then master list
        store = (count_display_by_store.get(norm)
                 or master_display_by_store.get(norm)
                 or norm.title())
        # District: count file first, then master list
        district = normalize_district(
            count_district_by_store.get(norm)
            or master_district_by_store.get(norm)
            or ts_district_by_store.get(norm)
            or "Unknown"
        )
        rep_name = safe_text(rep_by_store.get(norm, ""))
        if not rep_name:
            # Pending stores that never appear in the count file still get
            # their timesheet employee here (exact match, then tolerant
            # token-overlap match on the store name).
            rep_name = safe_text(match_store_employee(norm, ts_store_to_employee))
        status = "Completed" if norm in completed_store_norms else "Pending"
        key_raw = f"{norm}|{district}|{status}|status"
        key = hashlib.sha1(key_raw.encode("utf-8", errors="ignore")).hexdigest()
        rows.append(InventoryStatusRow(
            key=key, district=district, store=store,
            status=status, rep_name=rep_name, source_file=source_file
        ))

    summary = {
        "completed": len(completed_store_norms),
        "pending": len([r for r in rows if r.status == "Pending"]),
        "stores_total": len(rows),
        **latest_metrics,
    }
    return rows, summary


class VarianceDatabase:
    def __init__(self, db_path: Path):
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._ensure_wal_mode()
        self._backup_on_startup()
        self._init_db()

    def _ensure_wal_mode(self) -> None:
        try:
            con = sqlite3.connect(self.db_path, timeout=10, isolation_level=None)
            con.execute("PRAGMA journal_mode=WAL")
            con.execute("PRAGMA synchronous=NORMAL")
            con.close()
        except Exception:
            pass

    def _backup_on_startup(self) -> None:
        backup_path = self.db_path.with_suffix(".bak.sqlite3")
        try:
            if self.db_path.exists():
                con = sqlite3.connect(self.db_path, timeout=10)
                result = con.execute("PRAGMA integrity_check").fetchone()
                if result and result[0] == "ok":
                    import shutil
                    shutil.copy2(self.db_path, backup_path)
                else:
                    if backup_path.exists():
                        import shutil
                        shutil.copy2(backup_path, self.db_path)
                con.close()
        except Exception:
            pass

    def connect(self):
        con = sqlite3.connect(self.db_path, timeout=30,
                              isolation_level=None,
                              check_same_thread=False)
        try:
            con.execute("PRAGMA journal_mode=WAL")
            con.execute("PRAGMA synchronous=NORMAL")
            con.execute("PRAGMA busy_timeout=30000")
            con.execute("PRAGMA cache_size=-8000")
            con.execute("PRAGMA temp_store=MEMORY")
            con.execute("PRAGMA mmap_size=134217728")
        except Exception:
            pass
        con.isolation_level = ""
        return con

    @staticmethod
    def _table_columns(con: sqlite3.Connection, table: str) -> set[str]:
        try:
            return {row[1] for row in con.execute(f"PRAGMA table_info({table})").fetchall()}
        except Exception:
            return set()

    def _init_db(self) -> None:
        with self.connect() as con:
            con.execute(
                """
                CREATE TABLE IF NOT EXISTS variances (
                    key TEXT PRIMARY KEY,
                    district TEXT,
                    store TEXT,
                    product TEXT,
                    imei TEXT,
                    status TEXT,
                    created_by TEXT,
                    rep_name TEXT,
                    created_date TEXT,
                    document_status TEXT,
                    source_file TEXT,
                    first_seen_at TEXT,
                    last_seen_at TEXT,
                    cleared INTEGER DEFAULT 0,
                    cleared_at TEXT,
                    sent_count INTEGER DEFAULT 0,
                    last_sent_at TEXT,
                    notes TEXT
                )
                """
            )
            variance_columns = self._table_columns(con, "variances")
            for column, definition in {
                "rep_name": "TEXT DEFAULT ''",
                "created_by": "TEXT DEFAULT ''",
                "created_date": "TEXT DEFAULT ''",
                "document_status": "TEXT DEFAULT ''",
                "source_file": "TEXT DEFAULT ''",
                "notes": "TEXT DEFAULT ''",
            }.items():
                if column not in variance_columns:
                    con.execute(f"ALTER TABLE variances ADD COLUMN {column} {definition}")

            con.execute(
                """
                CREATE TABLE IF NOT EXISTS send_log (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    variance_key TEXT,
                    district TEXT,
                    group_name TEXT,
                    batch_title TEXT,
                    image_path TEXT,
                    sent_at TEXT,
                    mode TEXT,
                    error TEXT
                )
                """
            )
            send_columns = self._table_columns(con, "send_log")
            if "batch_title" not in send_columns:
                con.execute("ALTER TABLE send_log ADD COLUMN batch_title TEXT DEFAULT ''")

            con.execute(
                """
                CREATE TABLE IF NOT EXISTS inventory_status_state (
                    store_key TEXT PRIMARY KEY,
                    district TEXT,
                    store TEXT,
                    last_status TEXT,
                    previous_status TEXT,
                    status_changed_at TEXT,
                    last_sent_at TEXT,
                    last_loaded_at TEXT
                )
                """
            )

            con.execute(
                """
                CREATE TABLE IF NOT EXISTS whatsapp_groups (
                    district TEXT PRIMARY KEY,
                    group_name TEXT NOT NULL,
                    created_at TEXT,
                    updated_at TEXT
                )
                """
            )

            con.execute(
                """
                CREATE TABLE IF NOT EXISTS district_managers (
                    district TEXT PRIMARY KEY,
                    dm_name TEXT,
                    phone TEXT,
                    created_at TEXT,
                    updated_at TEXT
                )
                """
            )

            con.execute(
                """
                CREATE TABLE IF NOT EXISTS store_accounts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    district TEXT NOT NULL,
                    store TEXT NOT NULL,
                    store_key TEXT UNIQUE NOT NULL,
                    account_id TEXT,
                    username TEXT,
                    password TEXT,
                    created_at TEXT,
                    updated_at TEXT
                )
                """
            )
            con.execute(
                """
                CREATE TABLE IF NOT EXISTS sales_reps (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    rep_name TEXT NOT NULL,
                    rep_key TEXT UNIQUE NOT NULL,
                    phone TEXT NOT NULL,
                    created_at TEXT,
                    updated_at TEXT
                )
                """
            )

            con.execute(
                """
                CREATE TABLE IF NOT EXISTS device_exclusions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    rule_text TEXT NOT NULL,
                    rule_key TEXT UNIQUE NOT NULL,
                    match_type TEXT NOT NULL,
                    created_at TEXT,
                    updated_at TEXT
                )
                """
            )
            exclusion_columns = self._table_columns(con, "device_exclusions")
            if "product" not in exclusion_columns:
                con.execute("ALTER TABLE device_exclusions ADD COLUMN product TEXT DEFAULT ''")
            if "imei" not in exclusion_columns:
                con.execute("ALTER TABLE device_exclusions ADD COLUMN imei TEXT DEFAULT ''")
            if "comments" not in exclusion_columns:
                con.execute("ALTER TABLE device_exclusions ADD COLUMN comments TEXT DEFAULT ''")
            if "district" not in exclusion_columns:
                con.execute("ALTER TABLE device_exclusions ADD COLUMN district TEXT DEFAULT ''")
            # IMEI-ONLY exclusions (user rule 2026-09-15): purge legacy
            # product-level rules on every startup — a whole product can
            # never be excluded; a device only goes missing when its exact
            # IMEI is in the exclude list. Pre-schema IMEI-Exact rows
            # (imei column empty, rule_text = IMEI) survive the purge.
            con.execute(
                "DELETE FROM device_exclusions WHERE TRIM(COALESCE(imei, '')) = '' "
                "AND NOT (TRIM(COALESCE(rule_text, '')) <> '' AND REPLACE(LOWER(COALESCE(match_type, '')), ' ', '') IN ('imeiexact', 'serialexact', 'esnexact'))"
            )

            # Created By → Employee mapping table (Timesheet Edition)
            con.execute("""
                CREATE TABLE IF NOT EXISTS created_by_mappings (
                    created_by TEXT PRIMARY KEY,
                    employee_name TEXT DEFAULT '',
                    phone TEXT DEFAULT ''
                )
            """)

    def get_created_by_mappings(self) -> List[Dict[str, str]]:
        """Return all created_by → employee mappings."""
        with self.connect() as con:
            rows = con.execute(
                "SELECT created_by, employee_name, phone FROM created_by_mappings ORDER BY created_by"
            ).fetchall()
        return [{"created_by": r[0], "employee_name": r[1], "phone": r[2]} for r in rows]

    def upsert_created_by_mapping(self, created_by: str, employee_name: str, phone: str) -> None:
        with self.connect() as con:
            con.execute(
                "INSERT INTO created_by_mappings(created_by, employee_name, phone) VALUES(?,?,?) "
                "ON CONFLICT(created_by) DO UPDATE SET employee_name=excluded.employee_name, phone=excluded.phone",
                (created_by.strip(), employee_name.strip(), phone.strip()),
            )

    def delete_created_by_mapping(self, created_by: str) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM created_by_mappings WHERE created_by=?", (created_by,))

    def resolve_employee_for_created_by(self, created_by: str) -> Dict[str, str]:
        """Return {employee_name, phone} for a given Created By value, or empty strings."""
        with self.connect() as con:
            row = con.execute(
                "SELECT employee_name, phone FROM created_by_mappings WHERE created_by=?",
                (created_by.strip(),),
            ).fetchone()
        if row:
            return {"employee_name": row[0], "phone": row[1]}
        return {"employee_name": "", "phone": ""}

    def upsert_rows(self, rows: List[VarianceRow]) -> None:
        with self.connect() as con:
            for row in rows:
                existing = con.execute(
                    "SELECT cleared, cleared_at, sent_count, last_sent_at, notes, first_seen_at FROM variances WHERE key=?",
                    (row.key,),
                ).fetchone()
                if existing:
                    con.execute(
                        """
                        UPDATE variances
                        SET district=?, store=?, product=?, imei=?, status=?, created_by=?, rep_name=?, created_date=?,
                            document_status=?, source_file=?, last_seen_at=?
                        WHERE key=?
                        """,
                        (
                            row.district, row.store, row.product, row.imei, row.status,
                            row.created_by, row.rep_name, row.created_date, row.document_status,
                            row.source_file, now_text(), row.key,
                        ),
                    )
                else:
                    con.execute(
                        """
                        INSERT INTO variances (
                            key, district, store, product, imei, status, created_by, rep_name, created_date,
                            document_status, source_file, first_seen_at, last_seen_at, cleared,
                            cleared_at, sent_count, last_sent_at, notes
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, '', 0, '', '')
                        """,
                        (
                            row.key, row.district, row.store, row.product, row.imei, row.status,
                            row.created_by, row.rep_name, row.created_date, row.document_status, row.source_file,
                            now_text(), now_text(),
                        ),
                    )

    def rows(self, include_cleared: bool = False) -> List[VarianceRow]:
        query = """
            SELECT key, district, store, product, imei, status, created_by, rep_name, created_date,
                   document_status, source_file, cleared, sent_count, last_sent_at, cleared_at, notes
            FROM variances
        """
        if not include_cleared:
            query += " WHERE cleared=0"
        query += " ORDER BY district, store, rep_name, status, product"
        with self.connect() as con:
            data = con.execute(query).fetchall()
        rows: List[VarianceRow] = []
        for item in data:
            rows.append(
                VarianceRow(
                    key=item[0], district=item[1] or "", store=item[2] or "", product=item[3] or "",
                    imei=item[4] or "", status=item[5] or "", created_by=item[6] or "",
                    rep_name=item[7] or "", created_date=item[8] or "", document_status=item[9] or "",
                    source_file=item[10] or "", cleared=bool(item[11]), sent_count=int(item[12] or 0),
                    last_sent_at=item[13] or "", cleared_at=item[14] or "", notes=item[15] or "",
                )
            )
        return rows

    def get_rows_by_keys(self, keys: Iterable[str]) -> List[VarianceRow]:
        keys = list(keys)
        if not keys:
            return []
        placeholders = ",".join("?" for _ in keys)
        with self.connect() as con:
            data = con.execute(
                f"""
                SELECT key, district, store, product, imei, status, created_by, rep_name, created_date,
                       document_status, source_file, cleared, sent_count, last_sent_at, cleared_at, notes
                FROM variances WHERE key IN ({placeholders})
                ORDER BY district, store, rep_name, status, product
                """,
                keys,
            ).fetchall()
        return [
            VarianceRow(
                key=i[0], district=i[1] or "", store=i[2] or "", product=i[3] or "", imei=i[4] or "",
                status=i[5] or "", created_by=i[6] or "", rep_name=i[7] or "", created_date=i[8] or "",
                document_status=i[9] or "", source_file=i[10] or "", cleared=bool(i[11]),
                sent_count=int(i[12] or 0), last_sent_at=i[13] or "", cleared_at=i[14] or "", notes=i[15] or "",
            )
            for i in data
        ]

    def set_cleared(self, key: str, cleared: bool) -> None:
        with self.connect() as con:
            con.execute(
                "UPDATE variances SET cleared=?, cleared_at=? WHERE key=?",
                (1 if cleared else 0, now_text() if cleared else "", key),
            )

    def mark_sent(self, rows: Iterable[VarianceRow], group_name: str, batch_title: str, image_path: str, mode: str, error: str = "") -> None:
        sent_at = now_text()
        with self.connect() as con:
            for row in rows:
                con.execute(
                    """
                    INSERT INTO send_log (variance_key, district, group_name, batch_title, image_path, sent_at, mode, error)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (row.key, row.district, group_name, batch_title, image_path, sent_at, mode, error),
                )
                if not error:
                    con.execute(
                        "UPDATE variances SET sent_count=COALESCE(sent_count,0)+1, last_sent_at=? WHERE key=?",
                        (sent_at, row.key),
                    )


    def upsert_inventory_status_rows(self, rows: List[InventoryStatusRow]) -> None:
        loaded_at = now_text()
        with self.connect() as con:
            for row in rows:
                store_key = normalize_store(row.store)
                if not store_key:
                    continue
                existing = con.execute(
                    "SELECT last_status, previous_status, status_changed_at, last_sent_at FROM inventory_status_state WHERE store_key=?",
                    (store_key,),
                ).fetchone()

                last_status = safe_text(row.status)
                if existing:
                    old_status = safe_text(existing[0])
                    old_previous = safe_text(existing[1])
                    old_changed_at = safe_text(existing[2])
                    old_sent_at = safe_text(existing[3])
                    if old_status and normalize_header(old_status) != normalize_header(last_status):
                        previous_status = old_status
                        changed_at = loaded_at
                        # New status means new status still needs its own send state.
                        sent_at = ""
                    else:
                        previous_status = old_previous
                        changed_at = old_changed_at
                        sent_at = old_sent_at
                    con.execute(
                        """
                        UPDATE inventory_status_state
                        SET district=?, store=?, last_status=?, previous_status=?, status_changed_at=?, last_sent_at=?, last_loaded_at=?
                        WHERE store_key=?
                        """,
                        (row.district, row.store, last_status, previous_status, changed_at, sent_at, loaded_at, store_key),
                    )
                else:
                    con.execute(
                        """
                        INSERT INTO inventory_status_state
                        (store_key, district, store, last_status, previous_status, status_changed_at, last_sent_at, last_loaded_at)
                        VALUES (?, ?, ?, ?, '', '', '', ?)
                        """,
                        (store_key, row.district, row.store, last_status, loaded_at),
                    )

    def mark_status_sent(self, rows: Iterable[InventoryStatusRow]) -> None:
        sent_at = now_text()
        with self.connect() as con:
            for row in rows:
                store_key = normalize_store(row.store)
                if not store_key:
                    continue
                con.execute(
                    """
                    INSERT INTO inventory_status_state
                    (store_key, district, store, last_status, previous_status, status_changed_at, last_sent_at, last_loaded_at)
                    VALUES (?, ?, ?, ?, '', '', ?, ?)
                    ON CONFLICT(store_key) DO UPDATE SET
                        district=excluded.district,
                        store=excluded.store,
                        last_status=excluded.last_status,
                        last_sent_at=excluded.last_sent_at,
                        last_loaded_at=excluded.last_loaded_at
                    """,
                    (store_key, row.district, row.store, row.status, sent_at, sent_at),
                )

    def inventory_status_state_map(self) -> Dict[str, Dict[str, str]]:
        with self.connect() as con:
            rows = con.execute(
                """
                SELECT store_key, district, store, last_status, previous_status, status_changed_at, last_sent_at, last_loaded_at
                FROM inventory_status_state
                """
            ).fetchall()
        return {
            r[0]: {
                "District": r[1] or "",
                "Store": r[2] or "",
                "Last Status": r[3] or "",
                "Previous Status": r[4] or "",
                "Status Changed At": r[5] or "",
                "Last Sent At": r[6] or "",
                "Last Loaded At": r[7] or "",
            }
            for r in rows
        }

    def save_whatsapp_group(self, district: str, group_name: str) -> None:
        district = normalize_district(district)
        group_name = safe_text(group_name).strip()
        if not district or district == "Unknown":
            raise ValueError("District is required.")
        if not group_name:
            raise ValueError("WhatsApp group name is required.")
        with self.connect() as con:
            con.execute(
                """
                INSERT INTO whatsapp_groups (district, group_name, created_at, updated_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(district) DO UPDATE SET
                    group_name=excluded.group_name,
                    updated_at=excluded.updated_at
                """,
                (district, group_name, now_text(), now_text()),
            )

    def delete_whatsapp_group(self, district: str) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM whatsapp_groups WHERE district=?", (normalize_district(district),))

    def whatsapp_groups(self) -> List[Dict[str, str]]:
        with self.connect() as con:
            rows = con.execute(
                "SELECT district, group_name FROM whatsapp_groups ORDER BY district"
            ).fetchall()
        return [{"District": r[0] or "", "Group Name": r[1] or ""} for r in rows]

    def find_whatsapp_group(self, district: str) -> str:
        with self.connect() as con:
            row = con.execute(
                "SELECT group_name FROM whatsapp_groups WHERE district=?",
                (normalize_district(district),),
            ).fetchone()
        return safe_text(row[0]) if row else ""

    def get_setting(self, key: str, default: str = "") -> str:
        """Read a key from the app_settings table (created on first use)."""
        try:
            with self.connect() as con:
                con.execute(
                    "CREATE TABLE IF NOT EXISTS app_settings "
                    "(key TEXT PRIMARY KEY, value TEXT DEFAULT '')"
                )
                row = con.execute(
                    "SELECT value FROM app_settings WHERE key=?", (key,)
                ).fetchone()
            return safe_text(row[0]) if row else default
        except Exception:
            return default

    def save_setting(self, key: str, value: str) -> None:
        """Persist a key/value pair in the app_settings table."""
        try:
            with self.connect() as con:
                con.execute(
                    "CREATE TABLE IF NOT EXISTS app_settings "
                    "(key TEXT PRIMARY KEY, value TEXT DEFAULT '')"
                )
                con.execute(
                    "INSERT INTO app_settings(key,value) VALUES(?,?) "
                    "ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                    (key, value),
                )
        except Exception:
            pass

    def save_district_manager(self, district: str, dm_name: str, phone: str) -> None:
        district = normalize_district(district)
        dm_name = safe_text(dm_name)
        phone = normalize_phone(phone)
        if not district or district == "Unknown":
            raise ValueError("District is required.")
        if not phone:
            raise ValueError("DM phone number is required.")
        with self.connect() as con:
            con.execute(
                """
                INSERT INTO district_managers (district, dm_name, phone, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(district) DO UPDATE SET
                    dm_name=excluded.dm_name,
                    phone=excluded.phone,
                    updated_at=excluded.updated_at
                """,
                (district, dm_name, phone, now_text(), now_text()),
            )

    def delete_district_manager(self, district: str) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM district_managers WHERE district=?", (normalize_district(district),))

    def district_managers(self) -> List[Dict[str, str]]:
        with self.connect() as con:
            rows = con.execute(
                "SELECT district, dm_name, phone FROM district_managers ORDER BY district"
            ).fetchall()
        return [{"District": r[0] or "", "DM Name": r[1] or "", "Phone": r[2] or ""} for r in rows]

    def all_known_districts(self) -> List[str]:
        """Return sorted list of unique districts from all district-having tables."""
        districts = set()
        with self.connect() as con:
            for table in ("store_accounts", "whatsapp_groups", "district_managers", "device_exclusions", "variance_log"):
                try:
                    rows = con.execute(f"SELECT DISTINCT district FROM {table} WHERE district IS NOT NULL AND district != ''").fetchall()
                    for r in rows:
                        d = safe_text(r[0]).strip()
                        if d:
                            districts.add(normalize_district(d))
                except Exception:
                    pass
        return sorted(districts)

    def find_district_manager_phone(self, district: str) -> str:
        with self.connect() as con:
            row = con.execute(
                "SELECT phone FROM district_managers WHERE district=?",
                (normalize_district(district),),
            ).fetchone()
        return normalize_phone(row[0]) if row else ""

    def save_store_account(self, district: str, store: str, account_id: str, username: str, password: str) -> None:
        district = normalize_district(district)
        store = display_store(store)
        store_key = normalize_store(store)
        if not store_key:
            raise ValueError("Store name is required.")
        with self.connect() as con:
            con.execute(
                """
                INSERT INTO store_accounts (district, store, store_key, account_id, username, password, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(store_key) DO UPDATE SET
                    district=excluded.district,
                    store=excluded.store,
                    account_id=excluded.account_id,
                    username=excluded.username,
                    password=excluded.password,
                    updated_at=excluded.updated_at
                """,
                (district, store, store_key, safe_text(account_id), safe_text(username), safe_text(password), now_text(), now_text()),
            )

    def delete_store_account(self, store_key: str) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM store_accounts WHERE store_key=?", (store_key,))

    def store_accounts(self) -> List[Dict[str, str]]:
        with self.connect() as con:
            rows = con.execute(
                """
                SELECT district, store, store_key, account_id, username, password
                FROM store_accounts
                ORDER BY district, store
                """
            ).fetchall()
        return [
            {
                "District": row[0] or "",
                "Store": row[1] or "",
                "StoreKey": row[2] or "",
                "Account ID": row[3] or "",
                "Username": row[4] or "",
                "Password": row[5] or "",
            }
            for row in rows
        ]

    def replace_store_accounts(self, records: List[Dict[str, str]]) -> int:
        cleaned: Dict[str, Dict[str, str]] = {}
        for rec in records:
            district = normalize_district(rec.get("District", ""))
            store = display_store(rec.get("Store", ""))
            store_key = normalize_store(store)
            if not store_key or not district or district == "Unknown":
                continue
            cleaned[store_key] = {"District": district, "Store": store, "StoreKey": store_key}

        now = now_text()
        last_error = None
        for attempt in range(5):
            try:
                with self.connect() as con:
                    con.execute("BEGIN IMMEDIATE")
                    con.execute("DELETE FROM store_accounts")
                    for rec in sorted(cleaned.values(), key=lambda r: (r["District"], r["Store"].lower())):
                        con.execute(
                            """
                            INSERT INTO store_accounts (district, store, store_key, account_id, username, password, created_at, updated_at)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                            """,
                            (rec["District"], rec["Store"], rec["StoreKey"], "", "", "", now, now),
                        )
                    con.commit()
                return len(cleaned)
            except sqlite3.OperationalError as exc:
                last_error = exc
                if "locked" not in str(exc).lower():
                    raise
                time.sleep(0.8 + attempt * 0.6)
        raise RuntimeError("Database is locked. Close any other running copy of this GUI, then import again.") from last_error

    def store_master_records(self) -> List[Dict[str, str]]:
        return [{"District": row["District"], "Store": row["Store"]} for row in self.store_accounts()]

    def save_sales_rep(self, rep_name: str, phone: str) -> None:
        rep_name = safe_text(rep_name)
        phone = normalize_phone(phone)
        rep_key = person_name_key(rep_name)
        if not rep_key:
            raise ValueError("Sales rep name is required.")
        if not phone:
            raise ValueError("Phone number is required.")
        with self.connect() as con:
            con.execute(
                """
                INSERT INTO sales_reps (rep_name, rep_key, phone, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(rep_key) DO UPDATE SET
                    rep_name=excluded.rep_name,
                    phone=excluded.phone,
                    updated_at=excluded.updated_at
                """,
                (rep_name, rep_key, phone, now_text(), now_text()),
            )

    def delete_sales_rep(self, rep_key: str) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM sales_reps WHERE rep_key=?", (rep_key,))

    def sales_reps(self) -> List[Dict[str, str]]:
        with self.connect() as con:
            rows = con.execute(
                "SELECT rep_name, rep_key, phone FROM sales_reps ORDER BY rep_name"
            ).fetchall()
        return [{"Rep Name": r[0] or "", "RepKey": r[1] or "", "Phone": r[2] or ""} for r in rows]

    def sales_rep_phone_map(self) -> Dict[str, str]:
        return {row["RepKey"]: row["Phone"] for row in self.sales_reps()}

    def find_sales_rep_phone(self, rep_name: str) -> str:
        key = person_name_key(rep_name)
        if not key:
            return ""
        reps = self.sales_reps()
        exact = {row["RepKey"]: row["Phone"] for row in reps}
        if key in exact:
            return exact[key]
        # Drop single-character tokens (initials like "E" from "Menatallah.E",
        # "M" from "Abdullah.M"). person_name_key() splits on punctuation, so
        # dotted usernames become e.g. {"e", "menatallah"} — and the single
        # letter becomes a required word in the subset check below, which
        # breaks matching against employees whose names don't contain that
        # letter as a standalone word. Filtering both sides by len > 1
        # restores the subset match: {"menatallah"} <= {"elsafty", "menatallah"}.
        key_parts = {p for p in set(key.split()) if len(p) > 1}
        for row in reps:
            rep_parts = {p for p in set((row["RepKey"] or "").split()) if len(p) > 1}
            if key_parts and rep_parts and (key_parts <= rep_parts or rep_parts <= key_parts):
                return row["Phone"]
        return ""

    def resolve_phone_for_rep(self, rep_name: str, created_by: str = "") -> str:
        """Timesheet-aware phone resolver.

        Resolution order:
        1. created_by_mappings by created_by username (most direct — user set this)
        2. created_by_mappings by employee_name (fuzzy/normalized match)
        3. sales_reps table by rep_name (fuzzy/normalized match)
        4. created_by as a loose sales_reps key (last resort)

        Name matching uses person_name_key() normalization: lowercases,
        strips punctuation, sorts word tokens.  So "John Smith" matches
        "smith, john" or "John  Smith" (double space) or "JOHN SMITH".
        """
        # 1. Created By mapping by username
        if created_by:
            cb = created_by.strip()
            mapping = self.resolve_employee_for_created_by(cb)
            if mapping.get("phone"):
                return mapping["phone"]

        # 2. Created By mapping by employee_name (fuzzy match)
        if rep_name:
            rn_key = person_name_key(rep_name)
            if rn_key:
                for m in self.get_created_by_mappings():
                    emp_name = m.get("employee_name", "").strip()
                    if emp_name and person_name_key(emp_name) == rn_key and m.get("phone"):
                        return m["phone"]

        # 3. Standard sales_reps lookup (fuzzy match)
        phone = self.find_sales_rep_phone(rep_name)
        if phone:
            return phone

        # 4. Try created_by as a name in sales_reps
        if created_by:
            phone = self.find_sales_rep_phone(created_by)
            if phone:
                return phone

        return ""

    def update_device_exclusion_comment(self, rule_key: str, comments: str) -> None:
        with self.connect() as con:
            con.execute(
                "UPDATE device_exclusions SET comments=?, updated_at=? WHERE rule_key=?",
                (safe_text(comments), now_text(), rule_key),
            )

    def save_device_exclusion(self, product: str = "", imei: str = "", comments: str = "", district: str = "") -> None:
        product = safe_text(product)
        imei = safe_text(imei)
        comments = safe_text(comments)
        district = normalize_district(district) if safe_text(district) else ""

        # IMEI-ONLY exclusions (user rule 2026-09-15): the exclude list hides
        # a device only by its exact IMEI. Whole products can never be
        # excluded — a product-only request is rejected and the app asks for
        # the IMEI(s) instead. Product is kept on the row for display only.
        if not imei:
            raise ValueError(
                "IMEI is required — whole products cannot be excluded. "
                "Enter the IMEI(s) of the specific device(s) to hide."
            )
        # IMEI exclusions stay global across all districts.
        district = ""
        rule_text = imei
        match_type = "IMEI Exact"

        rule_key = device_rule_key(f"|{imei}")
        if not rule_key:
            raise ValueError("IMEI is required.")

        with self.connect() as con:
            # Dedupe: one rule per IMEI.
            con.execute("DELETE FROM device_exclusions WHERE imei=?", (imei,))

            con.execute(
                """
                INSERT INTO device_exclusions (rule_text, rule_key, match_type, product, imei, comments, district, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(rule_key) DO UPDATE SET
                    rule_text=excluded.rule_text,
                    match_type=excluded.match_type,
                    product=excluded.product,
                    imei=excluded.imei,
                    comments=excluded.comments,
                    district=excluded.district,
                    updated_at=excluded.updated_at
                """,
                (rule_text, rule_key, match_type, product, imei, comments, district, now_text(), now_text()),
            )

    def delete_device_exclusion(self, rule_key: str) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM device_exclusions WHERE rule_key=?", (rule_key,))

    def delete_all_device_exclusions(self) -> None:
        with self.connect() as con:
            con.execute("DELETE FROM device_exclusions")

    def device_exclusions(self) -> List[Dict[str, str]]:
        with self.connect() as con:
            rows = con.execute(
                """
                SELECT rule_text, rule_key, match_type, product, imei, comments, district
                FROM device_exclusions
                ORDER BY COALESCE(district, ''), COALESCE(product, ''), COALESCE(imei, '')
                """
            ).fetchall()

        out: List[Dict[str, str]] = []
        for r in rows:
            rule_text = r[0] or ""
            rule_key = r[1] or ""
            match_type = r[2] or ""
            product = r[3] or ""
            imei = r[4] or ""
            comments = r[5] or ""
            district = normalize_district(r[6] or "") if safe_text(r[6] or "") else ""

            # Backfill display values for exclusions created in old versions.
            if not product and not imei and rule_text:
                if normalize_header(match_type) in {"imeiexact", "serialexact", "esnexact"}:
                    imei = rule_text
                    district = ""
                else:
                    product = rule_text

            if imei:
                district = ""

            out.append({
                "District": district,
                "Product": product,
                "IMEI": imei,
                "Comments": comments,
                "Rule Text": rule_text,
                "RuleKey": rule_key,
                "Match Type": match_type,
            })
        return out

    def is_device_excluded(self, district: str, product: str, imei: str) -> bool:
        # IMEI-ONLY exclusions (user rule 2026-09-15): a device is hidden
        # only when its exact IMEI is in the exclude list — every other
        # device must show in the Variance Audit. Product-level exclusion is
        # not allowed; the district argument stays for call-site
        # compatibility but no longer participates in matching.
        imei_clean = device_rule_key(imei)
        if not imei_clean:
            return False

        for rule in self.device_exclusions():
            rule_imei = device_rule_key(rule.get("IMEI", ""))
            if rule_imei and rule_imei == imei_clean:
                return True

        return False

    def exclusion_reason(self, district: str, product: str, imei: str) -> str:
        # IMEI-ONLY exclusions (user rule 2026-09-15) — see
        # is_device_excluded: only an exact IMEI match produces a reason.
        imei_clean = device_rule_key(imei)

        for rule in self.device_exclusions():
            rule_imei = device_rule_key(rule.get("IMEI", ""))
            if rule_imei and imei_clean and rule_imei == imei_clean:
                return f"IMEI: {rule.get('IMEI', '')}"

        return ""

    def export_xlsx(self, output_path: Path, keys: Optional[Iterable[str]] = None) -> None:
        keys = list(keys) if keys is not None else None
        where_sql = ""
        params: List[str] = []
        if keys is not None:
            if not keys:
                headers = [
                    "District", "Store", "Product", "IMEI", "Status", "Rep Name", "Count By", "Created Date",
                    "Document Status", "Clearance Status", "Cleared At", "Sent Count", "Last Sent At",
                    "First Seen At", "Last Seen At", "Source File", "Notes"
                ]
                output_path.parent.mkdir(parents=True, exist_ok=True)
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(headers)
                wb.save(output_path)
                return
            placeholders = ",".join("?" for _ in keys)
            where_sql = f"WHERE key IN ({placeholders})"
            params = keys

        query = f"""
                SELECT district, store, product, imei, status, rep_name, created_by, created_date,
                       document_status,
                       CASE WHEN cleared=1 THEN 'Cleared' ELSE 'Not Cleared' END,
                       cleared_at, sent_count, last_sent_at, first_seen_at, last_seen_at, source_file, notes
                FROM variances
                {where_sql}
                ORDER BY district, store, rep_name, status, product
                """
        with self.connect() as con:
            rows = con.execute(query, params).fetchall()

        headers = [
            "District", "Store", "Product", "IMEI", "Status", "Rep Name", "Count By", "Created Date",
            "Document Status", "Clearance Status", "Cleared At", "Sent Count", "Last Sent At",
            "First Seen At", "Last Seen At", "Source File", "Notes"
        ]

        output_path.parent.mkdir(parents=True, exist_ok=True)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(headers)

        imei_col_idx = headers.index("IMEI") + 1

        for r_idx, row in enumerate(rows, start=2):
            for c_idx, val in enumerate(row, start=1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                if c_idx == imei_col_idx:
                    cell.data_type = 's'
                    cell.number_format = '@'

        wb.save(output_path)

class ImageRenderer:
    def __init__(self):
        IMAGE_DIR.mkdir(parents=True, exist_ok=True)
        if Image is None:
            raise RuntimeError("Pillow is required. Install with: py -m pip install pillow")

    @staticmethod
    def _font(size: int, bold: bool = False):
        candidates = []
        if bold:
            candidates.extend(["arialbd.ttf", "Arial Bold.ttf", "segoeuib.ttf"])
        candidates.extend(["arial.ttf", "Segoe UI.ttf", "DejaVuSans.ttf"])
        for candidate in candidates:
            try:
                return ImageFont.truetype(candidate, size)
            except Exception:
                continue
        return ImageFont.load_default()

    @staticmethod
    def _wrap(draw, text: str, font, max_width: int) -> List[str]:
        text = safe_text(text)
        if not text:
            return [""]
        words = text.split()
        lines: List[str] = []
        current = ""
        for word in words:
            probe = word if not current else current + " " + word
            bbox = draw.textbbox((0, 0), probe, font=font)
            if bbox[2] - bbox[0] <= max_width:
                current = probe
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)
        return lines or [""]

    def render_rows(self, batch_title: str, rows: List[VarianceRow], mode: str = "pending") -> Path:
        if not rows:
            raise ValueError("No rows to render")

        width = 1560
        margin = 32
        title_font = self._font(34, True)
        sub_font = self._font(18, False)
        header_font = self._font(20, True)
        cell_font = self._font(18, False)
        small_font = self._font(16, False)
        bold_small_font = self._font(16, True)
        row_height_base = 52

        tmp = Image.new("RGB", (width, 400), "white")
        draw = ImageDraw.Draw(tmp)
        col_widths = [165, 220, 480, 205, 140, 250]
        product_width = col_widths[2] - 20
        rep_width = col_widths[5] - 20
        row_heights = []
        for row in rows:
            product_lines = self._wrap(draw, str(row.product or ""), cell_font, product_width)
            rep_lines = self._wrap(draw, str(row.rep_name or ""), small_font, rep_width)
            row_heights.append(max(row_height_base, 26 * max(len(product_lines), len(rep_lines)) + 22))

        dark = (18, 20, 43)
        gray = (95, 95, 102)
        red = (233, 27, 47)
        light = (246, 247, 249)
        border = (215, 218, 223)

        logo_img = None
        logo_w = 0
        logo_h_used = 0
        if STATUS_LOGO_PATH.exists():
            try:
                logo_img = Image.open(STATUS_LOGO_PATH).convert("RGBA")
                scale = min(640 / logo_img.width, 150 / logo_img.height)
                size = (max(1, int(logo_img.width * scale)), max(1, int(logo_img.height * scale)))
                logo_img = logo_img.resize(size)
                logo_w, logo_h_used = size
            except Exception:
                logo_img = None
                logo_w = 0
                logo_h_used = 0

        header_area_h = max(logo_h_used, 74) + 18
        table_top = margin + header_area_h + 18
        table_height = 54 + sum(row_heights)
        height = table_top + table_height + 58

        img = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(img)

        clean_mode = safe_text(mode).replace("manual_", "").replace("pending_", "")
        mode_label = SEND_MODE_LABELS.get(clean_mode, clean_mode.replace("_", " ").title())

        # Standard GFH header: logo + red divider on the left, heading
        # CENTERED across the full width (matches every other GFH/VDS app
        # window's header layout) - was logo-on-the-right, title left-aligned.
        title_text = "GFH Inventory Variance"
        sub_text = f"{mode_label}: {batch_title}   Rows: {len(rows)}   Generated: {now_text()}"
        title_bbox = draw.textbbox((0, 0), title_text, font=title_font)
        title_w = title_bbox[2] - title_bbox[0]
        sub_bbox = draw.textbbox((0, 0), sub_text, font=sub_font)
        sub_w = sub_bbox[2] - sub_bbox[0]
        draw.text(((width - title_w) // 2, margin + 4), title_text, fill=dark, font=title_font)
        draw.text(((width - sub_w) // 2, margin + 48), sub_text, fill=gray, font=sub_font)

        if logo_img is not None:
            logo_x = margin
            logo_y = margin + (max(logo_h_used, 74) - logo_h_used) // 2
            img.paste(logo_img, (logo_x, logo_y), logo_img)
            divider_x = margin + logo_w + 14
            draw.rectangle(
                (divider_x, margin, divider_x + 3, margin + max(logo_h_used, 74)), fill=red)

        y = table_top
        x = margin
        headers = ["District", "Store", "Product", "IMEI", "Status", "Rep Name"]
        draw.rectangle((x, y, width - margin, y + 54), fill=red)
        cx = x
        for idx, header in enumerate(headers):
            draw.text((cx + 10, y + 14), header, fill="white", font=header_font)
            cx += col_widths[idx]
        y += 54

        for idx, row in enumerate(rows):
            fill = light if idx % 2 == 0 else (255, 255, 255)
            rh = row_heights[idx]
            draw.rectangle((x, y, width - margin, y + rh), fill=fill, outline=border)
            values = [row.district, row.store, row.product, row.imei, row.status, row.rep_name]
            cx = x
            for cidx, value in enumerate(values):
                max_w = col_widths[cidx] - 20
                font = bold_small_font if cidx in {0, 1, 5} else (small_font if cidx == 4 else cell_font)
                lines = self._wrap(draw, str(value or ""), font, max_w)
                yy = y + 11
                for line in lines[:4]:
                    draw.text((cx + 10, yy), line, fill=dark, font=font)
                    yy += 25
                cx += col_widths[cidx]
            y += rh

        footer = "Please provide resolution image or valid variance explanation. Cleared variances will not be auto-sent again."
        draw.text((margin, height - 36), footer, fill=gray, font=small_font)

        safe_title = re.sub(r"[^A-Za-z0-9_-]+", "_", batch_title)[:60] or "Variance"
        safe_mode = re.sub(r"[^A-Za-z0-9_-]+", "_", safe_text(mode).replace("manual_", "").replace("pending_", ""))[:30] or "mode"
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = IMAGE_DIR / f"GFH_Variance_{safe_mode}_{safe_title}_{stamp}.png"
        img.save(path)
        return path


# ── ROBUST WHATSAPP DESKTOP LAYER ───────────────────────────────────────────
# Root causes of "WhatsApp opens but nothing is typed":
#  1. _wa_find_hwnd matched ANY window with "WhatsApp" in the title (an Edge
#     "WhatsApp Web" tab, our own window...), so keystrokes went elsewhere.
#  2. SetForegroundWindow is silently refused by Windows (foreground lock) and
#     the result was never checked, so typing started in the wrong window.
#  3. keybd_event without scan codes / per-char VkKeyScan typing is ignored by
#     WhatsApp's WebView2 UI; text now goes through the clipboard (Ctrl+V) with
#     SendInput-Unicode as the fallback.
# Every step now re-verifies WhatsApp is the foreground window and ABORTS
# instead of typing into whatever else happens to be in front.
import ctypes as _wa_ct
from ctypes import wintypes as _wa_wt

_WA_BROWSER_PROCS = ("msedge", "chrome", "firefox", "brave", "opera",
                     "iexplore", "vivaldi")
_WA_SELF_PROCS = ("vidapay", "python")
_WA_EXT_VKS = (0x21, 0x22, 0x23, 0x24, 0x25, 0x26, 0x27, 0x28, 0x2D, 0x2E)


def _wa_user32():
    u = _wa_ct.windll.user32
    u.GetForegroundWindow.restype = _wa_wt.HWND
    u.GetAncestor.restype = _wa_wt.HWND
    u.GetAncestor.argtypes = [_wa_wt.HWND, _wa_wt.UINT]
    return u


def _wa_proc_name(pid):
    """Lower-case exe name of a process id ('' when unknown)."""
    try:
        k32 = _wa_ct.windll.kernel32
        k32.OpenProcess.restype = _wa_ct.c_void_p
        k32.OpenProcess.argtypes = [_wa_wt.DWORD, _wa_wt.BOOL, _wa_wt.DWORD]
        k32.QueryFullProcessImageNameW.argtypes = [
            _wa_ct.c_void_p, _wa_wt.DWORD, _wa_wt.LPWSTR,
            _wa_ct.POINTER(_wa_wt.DWORD)]
        k32.CloseHandle.argtypes = [_wa_ct.c_void_p]
        h = k32.OpenProcess(0x1000, False, pid)   # QUERY_LIMITED_INFORMATION
        if not h:
            return ""
        try:
            buf = _wa_ct.create_unicode_buffer(520)
            n = _wa_wt.DWORD(520)
            if k32.QueryFullProcessImageNameW(h, 0, buf, _wa_ct.byref(n)):
                return os.path.basename(buf.value).lower()
        finally:
            k32.CloseHandle(h)
    except Exception:
        pass
    return ""


def _wa_window_pid(hwnd):
    pid = _wa_wt.DWORD(0)
    _wa_user32().GetWindowThreadProcessId(hwnd, _wa_ct.byref(pid))
    return pid.value


def _wa_find_hwnd():
    """HWND of the real WhatsApp Desktop window (never a browser tab or this
    app's own window), or None. Minimised windows are included."""
    if sys.platform != "win32":
        return None
    try:
        u = _wa_user32()
        own = os.getpid()
        found = []
        enum_proc = _wa_ct.WINFUNCTYPE(_wa_wt.BOOL, _wa_wt.HWND, _wa_wt.LPARAM)

        def _cb(hwnd, _lp):
            try:
                if not u.IsWindowVisible(hwnd):
                    return True
                n = u.GetWindowTextLengthW(hwnd)
                if n <= 0:
                    return True
                buf = _wa_ct.create_unicode_buffer(n + 1)
                u.GetWindowTextW(hwnd, buf, n + 1)
                title = buf.value
                if "whatsapp" not in title.lower():
                    return True
                pid = _wa_window_pid(hwnd)
                if pid == own:
                    return True
                pn = _wa_proc_name(pid)
                if any(b in pn for b in _WA_BROWSER_PROCS):
                    return True
                if any(pn.startswith(s) for s in _WA_SELF_PROCS):
                    return True
                if "whatsapp" in pn:
                    score = 0
                elif pn.startswith("applicationframehost"):
                    score = 1
                elif title.strip().lower() == "whatsapp":
                    score = 2
                else:
                    return True
                found.append((score, hwnd))
            except Exception:
                pass
            return True

        u.EnumWindows(enum_proc(_cb), 0)
        found.sort(key=lambda t: t[0])
        return found[0][1] if found else None
    except Exception:
        return None


def _wa_is_foreground(hwnd):
    """True when WhatsApp (or its child/core window) owns the foreground."""
    try:
        u = _wa_user32()
        fg = u.GetForegroundWindow()
        if not fg:
            return False
        if fg == hwnd or u.GetAncestor(fg, 2) == hwnd:   # GA_ROOT
            return True
        return "whatsapp" in _wa_proc_name(_wa_window_pid(fg))
    except Exception:
        return False


def _wa_focus_hwnd(hwnd, log=None):
    """Bring WhatsApp to the foreground and VERIFY it. Returns True/False.
    Escalates: ALT-tap unlock + AttachThreadInput -> SwitchToThisWindow ->
    minimise/restore. Only un-minimises when actually minimised, so a snapped
    or maximised window keeps its layout."""
    try:
        u = _wa_user32()
        k32 = _wa_ct.windll.kernel32
        try:   # allow this process to take the foreground
            u.SystemParametersInfoW(0x2001, 0, 0, 0x0002)
        except Exception:
            pass
        if u.IsIconic(hwnd):
            u.ShowWindow(hwnd, 9)          # SW_RESTORE (from minimised only)
            time.sleep(0.5)
        for attempt in range(1, 5):
            _wa_win32_key_down(0x12)       # ALT tap unlocks SetForegroundWindow
            _wa_win32_key_up(0x12)
            fg = u.GetForegroundWindow()
            fg_thread = u.GetWindowThreadProcessId(fg, None) if fg else 0
            our_thread = k32.GetCurrentThreadId()
            wa_thread = u.GetWindowThreadProcessId(hwnd, None)
            pairs = [(fg_thread, our_thread), (fg_thread, wa_thread)]
            for a, b in pairs:
                if a and b and a != b:
                    u.AttachThreadInput(a, b, True)
            u.BringWindowToTop(hwnd)
            u.ShowWindow(hwnd, 5)          # SW_SHOW (keeps maximised state)
            u.SetForegroundWindow(hwnd)
            for a, b in pairs:
                if a and b and a != b:
                    u.AttachThreadInput(a, b, False)
            if attempt >= 2:
                u.SwitchToThisWindow(hwnd, True)
            if attempt >= 3:
                u.ShowWindow(hwnd, 6)      # SW_MINIMIZE
                time.sleep(0.3)
                u.ShowWindow(hwnd, 9)      # restores previous state
            time.sleep(0.4)
            if _wa_is_foreground(hwnd):
                return True
    except Exception as exc:
        if log:
            log(f"WhatsApp focus error: {exc}")
    return False


def _wa_ensure_focused(log=None):
    log = log or _wa_log
    """Find (launching if needed) WhatsApp Desktop, focus it, verify it.
    Returns hwnd on success, None on failure."""
    hwnd = _wa_find_hwnd()
    if hwnd is None:
        log("WhatsApp Desktop not open — launching...")
        _wa_launch()
        for _ in range(40):                # wait up to 20 s for the window
            time.sleep(0.5)
            hwnd = _wa_find_hwnd()
            if hwnd:
                time.sleep(2.0)            # let the chat list finish loading
                break
    if hwnd is None:
        log("WhatsApp Desktop window not found — is it installed and signed in?")
        return None
    for _ in range(3):
        if _wa_focus_hwnd(hwnd, log):
            time.sleep(0.4)
            log("WhatsApp Desktop focused.")
            return hwnd
        hwnd = _wa_find_hwnd() or hwnd
        time.sleep(1.0)
    hint = ""
    try:
        if _wa_ct.windll.shell32.IsUserAnAdmin():
            hint = (" This app is running as Administrator; Windows blocks it "
                    "from controlling a normal-user WhatsApp - run both at "
                    "the same level.")
    except Exception:
        pass
    log("Could not bring WhatsApp Desktop to the front." + hint)
    return None


def _wa_guard(hwnd, log=None):
    log = log or _wa_log
    """Re-verify (and re-take) focus right before typing. Never type blind."""
    if _wa_is_foreground(hwnd):
        return True
    log("WhatsApp lost focus — re-focusing...")
    if _wa_focus_hwnd(hwnd, log):
        time.sleep(0.3)
        return True
    log("WhatsApp is not in front — aborting so nothing is typed elsewhere.")
    return False


# ---- keyboard: scan-code aware key events + Unicode SendInput text ---------
class _WA_KEYBDINPUT(_wa_ct.Structure):
    _fields_ = [("wVk", _wa_wt.WORD), ("wScan", _wa_wt.WORD),
                ("dwFlags", _wa_wt.DWORD), ("time", _wa_wt.DWORD),
                ("dwExtraInfo", _wa_ct.c_size_t)]


class _WA_MOUSEINPUT(_wa_ct.Structure):
    _fields_ = [("dx", _wa_ct.c_long), ("dy", _wa_ct.c_long),
                ("mouseData", _wa_wt.DWORD), ("dwFlags", _wa_wt.DWORD),
                ("time", _wa_wt.DWORD), ("dwExtraInfo", _wa_ct.c_size_t)]


class _WA_INPUT(_wa_ct.Structure):
    class _U(_wa_ct.Union):
        _fields_ = [("ki", _WA_KEYBDINPUT), ("mi", _WA_MOUSEINPUT)]
    _anonymous_ = ("u",)
    _fields_ = [("type", _wa_wt.DWORD), ("u", _U)]


def _wa_send_key(vk, scan, flags):
    inp = _WA_INPUT(type=1)
    inp.ki = _WA_KEYBDINPUT(vk, scan, flags, 0, 0)
    _wa_ct.windll.user32.SendInput(1, _wa_ct.byref(inp), _wa_ct.sizeof(inp))


def _wa_win32_key_down(vk_code):
    scan = _wa_ct.windll.user32.MapVirtualKeyW(vk_code, 0)
    _wa_send_key(vk_code, scan, 0x1 if vk_code in _WA_EXT_VKS else 0)


def _wa_win32_key_up(vk_code):
    scan = _wa_ct.windll.user32.MapVirtualKeyW(vk_code, 0)
    _wa_send_key(vk_code, scan, 0x2 | (0x1 if vk_code in _WA_EXT_VKS else 0))


def _wa_win32_hotkey(*keys):
    codes = []
    for k in keys:
        kl = k.lower()
        if kl in _WA_VK_MAP:
            codes.append(_WA_VK_MAP[kl])
        elif len(k) == 1:
            codes.append(ord(k.upper()))
        else:
            return
    for c in codes:
        _wa_win32_key_down(c)
        time.sleep(0.04)
    time.sleep(0.04)
    for c in reversed(codes):
        _wa_win32_key_up(c)
        time.sleep(0.04)


def _wa_win32_press(key):
    kl = key.lower()
    if kl in _WA_VK_MAP:
        code = _WA_VK_MAP[kl]
    elif len(key) == 1:
        code = ord(key.upper())
    else:
        return
    _wa_win32_key_down(code)
    time.sleep(0.04)
    _wa_win32_key_up(code)
    time.sleep(0.04)


def _wa_win32_write(text, interval=0.0):
    """Type text as Unicode characters (layout independent, any language)."""
    for ch in str(text):
        for up in (0, 2):
            _wa_send_key(0, ord(ch), 0x4 | up)      # KEYEVENTF_UNICODE
        time.sleep(max(0.01, interval))


def _wa_set_clipboard_text(text):
    """Put plain text on the clipboard via ctypes (no pyperclip needed)."""
    u = _wa_ct.windll.user32
    k = _wa_ct.windll.kernel32
    k.GlobalAlloc.argtypes = [_wa_wt.UINT, _wa_ct.c_size_t]
    k.GlobalAlloc.restype = _wa_ct.c_void_p
    k.GlobalLock.argtypes = [_wa_ct.c_void_p]
    k.GlobalLock.restype = _wa_ct.c_void_p
    k.GlobalUnlock.argtypes = [_wa_ct.c_void_p]
    u.OpenClipboard.argtypes = [_wa_ct.c_void_p]
    u.SetClipboardData.argtypes = [_wa_wt.UINT, _wa_ct.c_void_p]
    u.SetClipboardData.restype = _wa_ct.c_void_p
    data = (str(text) + "\x00").encode("utf-16-le")
    for _ in range(10):                      # clipboard may be briefly busy
        if u.OpenClipboard(None):
            break
        time.sleep(0.05)
    else:
        return False
    try:
        u.EmptyClipboard()
        h = k.GlobalAlloc(0x0002, len(data))
        if not h:
            return False
        p = k.GlobalLock(h)
        _wa_ct.memmove(p, data, len(data))
        k.GlobalUnlock(h)
        return bool(u.SetClipboardData(13, h))   # CF_UNICODETEXT
    finally:
        u.CloseClipboard()



_WA_LOGGER = None


def _wa_log(msg):
    try:
        (_WA_LOGGER or print)(msg)
    except Exception:
        pass


def _wa_launch():
    """Open WhatsApp Desktop (several methods)."""
    import subprocess as _sp
    for cmd in (["explorer.exe", "whatsapp:"],
                "start whatsapp:",
                r'explorer.exe "shell:AppsFolder\5319275A.WhatsAppDesktop_cv1g1gvanyjgm!App"'):
        try:
            _sp.Popen(cmd, shell=isinstance(cmd, str))
            return
        except Exception:
            continue


def _wa_active_guard():
    """Never type blind: WhatsApp must be the foreground window right now."""
    hwnd = _wa_find_hwnd()
    if hwnd is None:
        raise RuntimeError("WhatsApp Desktop window not found - nothing was typed.")
    if not _wa_guard(hwnd):
        raise RuntimeError("WhatsApp Desktop is not in front - nothing was typed.")


def _wa_paste_text(text):
    """Enter text into the focused field: clipboard + Ctrl+V, falling back to
    Unicode typing (guarded)."""
    _wa_active_guard()
    try:
        if _wa_set_clipboard_text(text):
            time.sleep(0.2)
            _wa_win32_hotkey("ctrl", "v")
            time.sleep(0.3)
            return
    except Exception:
        pass
    _wa_win32_write(text, 0.02)


def _wa_hotkey(*keys):
    _wa_active_guard()
    _wa_win32_hotkey(*keys)


def _wa_press(key):
    _wa_active_guard()
    _wa_win32_press(key)


def _wa_write(text, interval=0.0):
    _wa_active_guard()
    _wa_win32_write(text, interval)
# ── END ROBUST WHATSAPP DESKTOP LAYER ───────────────────────────────────────



class WhatsAppSender:
    def __init__(self, status_callback=None, mode="desktop"):
        self.status_callback = status_callback or (lambda text: None)
        # "desktop" uses pyautogui to drive WhatsApp Desktop app.
        # "web" opens web.whatsapp.com in the default browser.
        self.mode = mode if mode in ("desktop", "web") else "desktop"

    def log(self, text: str) -> None:
        self.status_callback(text)

    def _force_focus_whatsapp(self) -> bool:
        """Find WhatsApp Desktop (launching it if needed), bring it to the front
        and VERIFY it is the foreground window. False => nothing may be typed."""
        global _WA_LOGGER
        _WA_LOGGER = self.log
        return _wa_ensure_focused(self.log) is not None

    @staticmethod
    def _import_pyautogui():
        try:
            import pyautogui
            pyautogui.FAILSAFE = True
            return pyautogui
        except Exception as exc:
            raise RuntimeError("pyautogui is required. Install with: py -m pip install pyautogui") from exc

    @staticmethod
    def _open_whatsapp() -> None:
        if sys.platform.startswith("win"):
            try:
                subprocess.Popen("start whatsapp:", shell=True)
                time.sleep(3)
                return
            except Exception:
                pass
        time.sleep(1)

    @staticmethod
    def _activate_whatsapp_window() -> bool:
        try:
            import pygetwindow as gw
            windows = [w for w in gw.getAllWindows() if "whatsapp" in (w.title or "").lower()]
            if not windows:
                return False
            win = windows[0]
            if win.isMinimized:
                win.restore()
            win.activate()
            time.sleep(1)
            return True
        except Exception:
            return False

    @staticmethod
    def _copy_image_to_clipboard(image_path: Path) -> None:
        """Copy image to clipboard — same as GFH_Inventory_Audit.py (proven working)."""
        if Image is None:
            raise RuntimeError("Pillow is required to copy images to clipboard.")
        if not sys.platform.startswith("win"):
            raise RuntimeError("Automatic image clipboard paste is implemented for Windows only.")
        try:
            import win32clipboard
            import win32con
        except Exception as exc:
            raise RuntimeError("pywin32 is required. Install with: py -m pip install pywin32") from exc
        image = Image.open(image_path).convert("RGB")
        output = BytesIO()
        image.save(output, "BMP")
        data = output.getvalue()[14:]
        output.close()
        win32clipboard.OpenClipboard()
        try:
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardData(win32con.CF_DIB, data)
        finally:
            win32clipboard.CloseClipboard()

    @staticmethod
    def _type_group_search(pyautogui, group_name: str) -> None:
        shortcut = SEARCH_SHORTCUT.lower().strip()
        if "+" in shortcut:
            keys = [k.strip() for k in shortcut.split("+") if k.strip()]
        else:
            keys = [k.strip() for k in shortcut.split() if k.strip()]
        if not keys:
            keys = ["ctrl", "f"]
        _wa_hotkey(*keys)
        time.sleep(0.7)
        _wa_hotkey("ctrl", "a")
        time.sleep(0.2)
        _wa_paste_text(group_name)
        time.sleep(1.2)
        _wa_press("enter")
        time.sleep(1.5)

    @staticmethod
    def _paste_text(pyautogui, text: str) -> None:
        text = safe_text(text)
        if not text:
            return
        _wa_paste_text(text)

    def send_image(self, group_name: str, image_path: Path, text_message: str = "") -> None:
        if self.mode == "web":
            self._send_image_web(group_name, image_path, text_message)
            return
        pyautogui = self._import_pyautogui()
        self.log(f"Opening WhatsApp Desktop for {group_name}...")
        self._open_whatsapp()
        if not self._force_focus_whatsapp():
            raise RuntimeError("Could not bring WhatsApp Desktop to the front - nothing was typed.")
        self.log(f"Searching group: {group_name}")
        self._type_group_search(pyautogui, group_name)

        self.log("Copying image to clipboard...")
        self._copy_image_to_clipboard(image_path)

        # Paste the image — this opens WhatsApp's image-preview/send dialog
        _wa_hotkey("ctrl", "v")
        # Wait for the preview dialog to fully load before doing anything else
        time.sleep(3.0)

        caption = safe_text(text_message)
        if caption:
            self.log("Typing caption into image preview field...")
            # WhatsApp Desktop focuses the caption field automatically when the
            # image preview dialog opens. Type directly into it.
            self._paste_text(pyautogui, caption)
            time.sleep(0.8)

        # Send the image (Enter confirms the image-preview dialog)
        _wa_press("enter")
        time.sleep(2.0)
        self.log(f"Sent image to {group_name}")

    def _send_image_web(self, group_name: str, image_path: Path, text_message: str = "") -> None:
        """Send image via WhatsApp Web by opening it in the default browser."""
        self.log(f"WhatsApp Web: attaching image for {group_name} (manual step required — "
                 "open web.whatsapp.com, find the group, and attach the image manually).")
        self.log(f"Image path: {image_path}")

    def send_text(self, group_name: str, text_message: str) -> None:
        if self.mode == "web":
            self._send_text_web(group_name, text_message)
            return
        message = safe_text(text_message)
        if not message:
            return
        pyautogui = self._import_pyautogui()
        self.log(f"Opening WhatsApp Desktop for {group_name}...")
        self._open_whatsapp()
        if not self._force_focus_whatsapp():
            raise RuntimeError("Could not bring WhatsApp Desktop to the front - nothing was typed.")
        self.log(f"Searching group: {group_name}")
        self._type_group_search(pyautogui, group_name)
        self.log("Sending WhatsApp text message...")
        self._paste_text(pyautogui, message)
        time.sleep(0.7)
        _wa_press("enter")
        time.sleep(1.0)
        self.log(f"Sent text to {group_name}")

    def _send_text_web(self, group_name: str, text_message: str) -> None:
        """Send text via WhatsApp Web by opening a direct wa.me link."""
        message = safe_text(text_message)
        if not message:
            return
        import urllib.parse, webbrowser
        # wa.me won't work for group chats (no phone number) — open web.whatsapp.com
        # and log the message so the user can paste it manually if needed.
        url = f"https://web.whatsapp.com"
        webbrowser.open(url)
        self.log(f"WhatsApp Web opened. Group: {group_name!r} | Message: {message!r}")
        self.log("Tip: use WhatsApp Web search (Ctrl+F) to find the group and paste the message.")


GFH_SQUARE_ICON_B64 = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "gfh_square_icon_b64.txt"), "r").read().strip() if not getattr(sys, "frozen", False) else open(os.path.join(getattr(sys, "_MEIPASS", "."), "assets", "gfh_square_icon_b64.txt"), "r").read().strip()


# ==========================================================
# QT UI LAYER
# ==========================================================
# The Tkinter UI class becomes MainWindow(QMainWindow). Same window
# title, same tab names/order, same button labels, same dialogs
# (Toplevel -> QDialog, messagebox -> QMessageBox,
# filedialog -> QFileDialog), same behaviors; worker-thread UI updates go
# through UIBridge(QObject) signals only.
from PySide6.QtCore import Qt, QObject, Signal, QTimer
from PySide6.QtGui import QIcon, QColor, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QLineEdit, QComboBox, QCheckBox, QRadioButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
    QTabWidget, QGroupBox, QSplitter, QMessageBox, QFileDialog,
    QInputDialog,
)

from header_manager_qt import FixedHeaderManagerQt
from native_titlebar_qt import sync_titlebar_qt


def _extract_embedded_icon(b64, filename):
    try:
        if not b64:
            return None
        import base64 as _b64
        target = os.path.join(tempfile.gettempdir(), filename)
        with open(target, "wb") as fh:
            fh.write(_b64.b64decode(b64))
        return target if os.path.isfile(target) else None
    except Exception:
        return None


def _set_window_icon_qt(window):
    """Set taskbar + titlebar icon - same 3-tier resolution order as the
    Tkinter build's window iconbitmap chain: _MEIPASS ICO files
    first, then the .ico shipped beside the script, then the embedded
    base64 icon decoded to %TEMP%."""
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        for ico_name in ("GFH_icon.ico", "gfh_telecom_llc_icon.ico"):
            ico_path = os.path.join(meipass, ico_name)
            if os.path.exists(ico_path):
                window.setWindowIcon(QIcon(ico_path))
                return
    base_dir = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
        else os.path.dirname(os.path.abspath(__file__))
    ico_path = os.path.join(base_dir, "GFH_icon.ico")
    if os.path.exists(ico_path):
        window.setWindowIcon(QIcon(ico_path))
        return
    ico_path = os.path.join(base_dir, "gfh_telecom_llc_icon.ico")
    if os.path.exists(ico_path):
        window.setWindowIcon(QIcon(ico_path))
        return
    try:
        ico_path = _extract_embedded_icon(EMBEDDED_ICON_B64, "gfh_audit_icon.ico")
        if ico_path:
            window.setWindowIcon(QIcon(ico_path))
    except Exception:
        pass


class UIBridge(QObject):
    """Thread-safe UI update delivery (WhatsApp send worker threads -> GUI
    thread). Replaces the Tkinter build's direct (technically unsafe)
    cross-thread StringVar.set / messagebox / Treeview calls."""
    status = Signal(str)
    dialog = Signal(str, str, str)          # kind: info|warning|error
    refresh_audit = Signal()
    refresh_status = Signal()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        _vp_hide(self)
        self.setWindowTitle(APP_NAME + " — Timesheet Edition")
        self._app_icon = None
        # Set the window/taskbar icon. Same resolution order as the Tkinter
        # build: _MEIPASS .ico first, the script-side .ico second, the
        # embedded base64 icon last.
        ensure_app_icon_files()
        _set_window_icon_qt(self)
        self.zoom_scale = 1.0
        APP_DIR.mkdir(parents=True, exist_ok=True)
        IMAGE_DIR.mkdir(parents=True, exist_ok=True)
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        migrate_legacy_app_data_if_needed()
        self.db = VarianceDatabase(DB_PATH)
        # Track file mtime for cross-instance sync.
        try:
            self._db_mtime: float = DB_PATH.stat().st_mtime if DB_PATH.exists() else 0.0
        except Exception:
            self._db_mtime = 0.0
        self._db_sync_paused: bool = False   # paused while THIS instance is writing
        self._db_sync_after_id = None        # cancellable after() handle
        self.master_store_records = self.db.store_master_records()
        self.current_inventory_records: List[Dict[str, str]] = []
        self.current_time_sheet_records: List[Dict[str, str]] = []

        # tk.StringVar/BooleanVar -> plain Python attributes (R4)
        self.selected_exclusion_key: str = ""
        self.status_text: str = "Select the Inventory_Count_Result_Details file, then click Load Variances. No data loaded yet."
        self.summary_text: str = "No data loaded"
        self.include_cleared: bool = False
        self.send_only_unsent: bool = SEND_ONLY_UNSENT_BY_DEFAULT

        self.loaded_keys: set = set()
        self.data_loaded = False
        self.key_by_iid: Dict[str, str] = {}

        self.status_rows: List[InventoryStatusRow] = []
        self.status_row_by_key: Dict[str, InventoryStatusRow] = {}
        self.status_key_by_iid: Dict[str, str] = {}
        self.status_checked_keys: set = set()
        self.audit_checked_keys: set = set()

        self._apply_dynamic_geometry()
        self.theme_manager = ThemeManagerQt(default="dark", app_name="VidaPay-GFH")
        # Always start dark — override any saved light preference
        self.theme_manager.current_theme = "dark"

        self._bridge = UIBridge()
        self._bridge.status.connect(self._set_status_label)
        self._bridge.dialog.connect(self._show_dialog)
        self._bridge.refresh_audit.connect(self.refresh_table)
        self._bridge.refresh_status.connect(self.refresh_status_table)

        # ── Outer layout: header (navy) / body / footer copyright — exact
        # FixedHeaderManagerQt order required by the port pattern.
        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.header_mgr = FixedHeaderManagerQt(outer, title=APP_NAME)
        self.header_mgr.set_logo(logo_path=str(HEADER_LOGO_PATH), text="GFH")
        self.header_mgr.add_theme_toggle(self.theme_manager, callback=self._apply_theme)

        body = QWidget()
        body_layout = QVBoxLayout(body)
        body_layout.setContentsMargins(14, 12, 14, 10)
        body_layout.setSpacing(6)
        outer.addWidget(body, 1)

        self._build_ui_body(body_layout)

        self.header_mgr.add_copyright(self.theme_manager, outer)   # LAST

        self.set_status(f"Ready. Data folder: {APP_DIR}")
        self._start_db_sync_poll()
        self._apply_row_tag_colors()
        self._apply_theme()
        self.resize(1400, 900)
        # The Tk app opens maximized — replicate.
        QTimer.singleShot(0, self.showMaximized)

    # ── Window state save/restore for WhatsApp sending ──────────────────
    # When WhatsApp Desktop is brought to the foreground, Windows can push
    # this window behind it. These methods save the window geometry+state
    # before WhatsApp is focused and restore it after (Qt port of the Tk
    # geometry/state() pair).
    def _save_window_state(self) -> dict:
        """Save the current window geometry and state (maximized/normal)
        so it can be restored after WhatsApp steals focus."""
        try:
            return {
                "geometry": self.geometry(),
                "state": self.windowState(),
            }
        except Exception:
            return {}

    def _restore_window_state(self, saved: dict) -> None:
        """Restore window geometry + state saved by _save_window_state.
        Called after WhatsApp finishes sending — brings the app back to
        the exact position/size/state it was in before."""
        if not saved:
            return
        try:
            state = saved.get("state", Qt.WindowNoState)
            geom = saved.get("geometry", None)
            if state & Qt.WindowMaximized:
                self.setWindowState(Qt.WindowMaximized)
            elif geom is not None:
                self.setGeometry(geom)
            # Bring the app back to the foreground
            try:
                self.raise_()
                self.activateWindow()
            except Exception:
                pass
        except Exception:
            pass

    def _apply_dynamic_geometry(self) -> None:
        """Fill the entire work area: the window opens TRUE MAXIMIZED
        (QTimer.singleShot(0, showMaximized) in __init__), and the
        maximized state is re-asserted once mapped (Qt port of the Tk
        build's work-area geometry + state('zoomed') + reassert)."""
        try:
            screen = QApplication.primaryScreen()
            geo = screen.availableGeometry() if screen is not None else None
            if geo is not None:
                sw, sh = geo.width(), geo.height()
                self.setMinimumSize(min(1220, max(560, sw // 2)),
                                    min(740, max(420, sh // 2)))
            self.setWindowState(Qt.WindowMaximized)
            QTimer.singleShot(60, self._reassert_zoomed)
        except Exception:
            pass

    def _reassert_zoomed(self) -> None:
        """Re-assert the maximized state once the window has mapped."""
        try:
            self.showMaximized()
        except Exception:
            pass

    def zoom_in(self, event=None) -> None:
        if self.zoom_scale < 2.0:
            self.zoom_scale = round(self.zoom_scale + 0.1, 1)
            self.apply_zoom()

    def zoom_out(self, event=None) -> None:
        if self.zoom_scale > 0.5:
            self.zoom_scale = round(self.zoom_scale - 0.1, 1)
            self.apply_zoom()

    def apply_zoom(self) -> None:
        """Scale the UI font (the Tk build scaled ttk style fonts; the Qt
        build appends one app-wide font-size rule on top of the theme QSS)."""
        self._apply_zoom_qss()

    def _apply_zoom_qss(self) -> None:
        app = QApplication.instance()
        if app is None:
            return
        try:
            size = max(6, round(13 * self.zoom_scale))
            base = re.sub(r"\n?/\* gfh-zoom \*/ QWidget \{ font-size: \d+px; \}", "",
                          app.styleSheet() or "")
            app.setStyleSheet(base + f"\n/* gfh-zoom */ QWidget {{ font-size: {size}px; }}")
        except Exception:
            pass

    def _build_ui_body(self, body_layout) -> None:
        """Upload Files box + notebook (6 tabs) + status bar. Same widget
        labels and order as the Tkinter _build_ui()."""
        file_box = QGroupBox("Upload Files")
        grid = QGridLayout(file_box)
        self.inventory_edit = QLineEdit()
        self.time_sheet_edit = QLineEdit()
        self._file_row(grid, 0, "Inventory_Count_Result_Details", self.inventory_edit, self.pick_inventory)
        self._file_row(grid, 1, "Timesheet (timesheets_*.xlsx)", self.time_sheet_edit, self.pick_time_sheet)
        load_btn = QPushButton("Load Variances")
        load_btn.clicked.connect(self.load_variances)
        grid.addWidget(load_btn, 0, 3, 2, 1)
        # + / − zoom buttons stacked vertically (+ on top, − on bottom)
        zoom_frame = QWidget()
        zl = QVBoxLayout(zoom_frame)
        zl.setContentsMargins(0, 0, 0, 0)
        plus_btn = QPushButton("+")
        plus_btn.setFixedWidth(44)
        plus_btn.clicked.connect(self.zoom_in)
        minus_btn = QPushButton("−")
        minus_btn.setFixedWidth(44)
        minus_btn.clicked.connect(self.zoom_out)
        zl.addWidget(plus_btn)
        zl.addWidget(minus_btn)
        grid.addWidget(zoom_frame, 0, 4, 2, 1)
        grid.setColumnStretch(1, 1)
        body_layout.addWidget(file_box)

        sc_zoom_in = QShortcut(QKeySequence("Ctrl+="), self)
        sc_zoom_in.activated.connect(self.zoom_in)
        sc_zoom_plus = QShortcut(QKeySequence("Ctrl++"), self)
        sc_zoom_plus.activated.connect(self.zoom_in)
        sc_zoom_out = QShortcut(QKeySequence("Ctrl+-"), self)
        sc_zoom_out.activated.connect(self.zoom_out)

        self.tabs = QTabWidget()
        body_layout.addWidget(self.tabs, 1)
        self.status_tab = QWidget()
        self.audit_tab = QWidget()
        self.store_tab = QWidget()
        self.rep_tab = QWidget()
        self.dm_tab = QWidget()
        self.exclusion_tab = QWidget()
        self.tabs.addTab(self.status_tab, "Inventory Audit Status")
        self.tabs.addTab(self.audit_tab, "Variance Audit")
        self.tabs.addTab(self.store_tab, "Store List")
        self.tabs.addTab(self.rep_tab, "Employees")
        self.tabs.addTab(self.dm_tab, "District DMs")
        self.tabs.addTab(self.exclusion_tab, "Excluded Devices")

        self._build_status_tab()
        self._build_audit_tab()
        self._build_store_tab()
        self._build_rep_tab()
        self._build_dm_tab()
        self._build_exclusion_tab()
        self.refresh_store_accounts_table()
        self.refresh_sales_reps_table()
        self.refresh_district_managers_table()
        self.refresh_whatsapp_groups_table()
        self.refresh_device_exclusions_table()

        # Status bar — fixed light bar with navy text, same colors as the Tk build.
        self.status_bar = QLabel(self.status_text)
        self.status_bar.setStyleSheet("background:#E9ECF5; color:#090d26; padding:6px;")
        body_layout.addWidget(self.status_bar)

    def _file_row(self, grid, row: int, label: str, edit, command) -> None:
        grid.addWidget(QLabel(label), row, 0)
        grid.addWidget(edit, row, 1)
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(command)
        grid.addWidget(browse_btn, row, 2)

    def _apply_theme(self, colors=None) -> None:
        """Apply theme colors to all widgets (Qt: one app-wide QSS;
        native title bars re-synced inside apply_theme_to_app)."""
        colors = self.theme_manager.get_colors()
        self.theme_manager.apply_theme_to_app(QApplication.instance())
        self.header_mgr.update_for_theme()
        # Re-apply native titlebar sync explicitly (apply_theme_to_app also
        # walks all top-level widgets; this pins the main window).
        try:
            sync_titlebar_qt(self, dark=(self.theme_manager.current_theme == "dark"))
        except Exception:
            pass
        self.COLOR_BG = colors["bg"]
        self.COLOR_CARD = colors["panel"]
        self.COLOR_TEXT = colors["text"]
        self.COLOR_MUTED = colors["text_dim"]
        self.COLOR_BORDER = colors["border"]
        self.COLOR_INPUT = colors.get("input", colors["panel"])
        self.COLOR_PANEL_ALT = colors.get("panel_alt", colors["panel"])
        # Row-highlight colors for the two audit tables MUST follow the
        # active theme (same palette as the Tk _apply_row_tag_colors).
        self._apply_row_tag_colors()
        self._apply_zoom_qss()
        # Recolor already-rendered rows with the new palette.
        try:
            if self.data_loaded:
                self.refresh_status_table()
                self.refresh_table()
        except Exception:
            pass

    def _apply_row_tag_colors(self) -> None:
        """Theme-aware row highlight colors for the two audit tables.

        Light theme keeps the classic pastel highlights:
            yellow = pending, blue = sent / completed-after-update, green = cleared.
        Dark theme swaps them for muted dark tints of the same families with
        light text (same values as the Tkinter tag_configure calls).
        """
        dark = getattr(self.theme_manager, "current_theme", "dark") == "dark"
        if dark:
            pending = ("#3a3117", "#ffd66e")   # dark yellow tint
            sent = ("#16324f", "#9ecbff")      # dark blue tint
            cleared = ("#17331f", "#8fe3a8")   # dark green tint
        else:
            pending = ("#FFF3CD", "#111827")
            sent = ("#D7ECFF", "#111827")
            cleared = ("#D9F7DF", "#111827")
        colors = self.theme_manager.get_colors()
        normal_bg = getattr(self, "COLOR_CARD", None) or colors.get("panel", ("#141b38" if dark else "#ffffff"))
        normal_fg = getattr(self, "COLOR_TEXT", None) or colors.get("text", ("#e8ecf7" if dark else "#111827"))
        self._row_tag_colors = {
            "status_pending": (pending[0], pending[1]),
            "status_completed_after_update": (sent[0], sent[1]),
            "status_completed_sent": (cleared[0], cleared[1]),
            "status_completed": (normal_bg, normal_fg),
            "variance_pending": (pending[0], pending[1]),
            "variance_sent": (sent[0], sent[1]),
            "variance_cleared": (cleared[0], cleared[1]),
        }

    def _paint_row(self, table, row: int, tag: str) -> None:
        """Apply the theme colors of a Tk row-tag to every item of one
        QTableWidget row."""
        colors = getattr(self, "_row_tag_colors", {})
        bg, fg = colors.get(tag, (None, None))
        if bg is None:
            return
        for col in range(table.columnCount()):
            item = table.item(row, col)
            if item is not None:
                item.setBackground(QColor(bg))
                item.setForeground(QColor(fg))

    def _make_table(self, headers, widths, multi_select=False) -> QTableWidget:
        """ttk.Treeview(show='headings') -> QTableWidget. Same headings,
        same column widths, read-only, row selection, Qt-native scrollbars."""
        table = QTableWidget(0, len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.verticalHeader().setVisible(False)
        table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        table.setSelectionBehavior(QAbstractItemView.SelectRows)
        table.setSelectionMode(QAbstractItemView.ExtendedSelection if multi_select
                               else QAbstractItemView.SingleSelection)
        table.setWordWrap(False)
        table.setTextElideMode(Qt.ElideRight)
        for col, width in enumerate(widths):
            table.setColumnWidth(col, width)
        return table

    def _wire_sorting(self, table) -> None:
        """Heading-click sort with reverse toggle on re-click (port of the
        Tkinter sort_any_tree heading command binding)."""
        header = table.horizontalHeader()
        header.setSectionsClickable(True)
        header.setSortIndicatorShown(True)

        def _on_section(logical_index: int) -> None:
            state = getattr(table, "_gfh_sort_reverse", {})
            reverse = not state.get(logical_index, False)
            state[logical_index] = reverse
            table._gfh_sort_reverse = state
            self.sort_any_tree(table, logical_index, reverse)

        header.sectionClicked.connect(_on_section)

    @staticmethod
    def sort_any_tree(table, col: int, reverse: bool) -> None:
        """Sort a QTableWidget by column value (ttk.Treeview heading-sort
        port; repeated clicks reverse the order)."""
        table.sortItems(col, Qt.DescendingOrder if reverse else Qt.AscendingOrder)
        table.horizontalHeader().setSortIndicator(
            col, Qt.DescendingOrder if reverse else Qt.AscendingOrder)

    @staticmethod
    def _combo_set(combo, value: str) -> None:
        """Select `value` in a combo (Tk combobox.set() port); a value not
        in the list clears the selection to blank, same as setting a Tk
        textvariable to "".  """
        combo.setCurrentIndex(combo.findText(value))

    # ── Status / dialog helpers (worker threads -> UIBridge) ────────────
    def set_status(self, text: str) -> None:
        self.status_text = text
        self._bridge.status.emit(text)

    def _set_status_label(self, text: str) -> None:
        self.status_text = text
        self.status_bar.setText(text)
        try:
            self.status_bar.repaint()
        except Exception:
            pass

    def _worker_dialog(self, kind: str, title: str, text: str) -> None:
        """Show a messagebox from a worker thread — marshaled to the GUI
        thread through the UIBridge (Tk messagebox calls in the send
        threads were cross-thread already; Qt requires the GUI thread)."""
        self._bridge.dialog.emit(kind, title, text)

    def _show_dialog(self, kind: str, title: str, text: str) -> None:
        if kind == "error":
            QMessageBox.critical(self, title, text)
        elif kind == "warning":
            QMessageBox.warning(self, title, text)
        else:
            QMessageBox.information(self, title, text)


    # ------------------------------------------------------------
    # TAB 1: Inventory Audit Status
    # ------------------------------------------------------------
    def _build_status_tab(self) -> None:
        tab_outer = QVBoxLayout(self.status_tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)

        controls = QGroupBox("Send Inventory Audit Status")
        crow = QHBoxLayout(controls)
        crow.addWidget(QLabel("Send by:"))
        self.status_send_mode_combo = QComboBox()
        self.status_send_mode_combo.addItems(["District", "Store", "Sales Rep"])
        crow.addWidget(self.status_send_mode_combo)
        crow.addSpacing(12)
        crow.addWidget(QLabel("District filter:"))
        self.status_district_combo = QComboBox()
        self.status_district_combo.addItem("All Districts")
        # activated = user selection only (same trigger as <<ComboboxSelected>>)
        self.status_district_combo.activated.connect(lambda _i: self.on_status_district_change())
        crow.addWidget(self.status_district_combo)
        crow.addSpacing(12)
        crow.addWidget(QLabel("Store filter:"))
        self.status_store_combo = QComboBox()
        self.status_store_combo.addItem("All Stores")
        self.status_store_combo.activated.connect(lambda _i: self.on_status_store_change())
        crow.addWidget(self.status_store_combo)
        crow.addSpacing(6)
        for text, fn in (("Check Current Filter", self.auto_check_status_rows),
                         ("Check Pending Only", self.auto_check_pending_status_rows),
                         ("Clear Checkmarks", self.clear_status_checkmarks),
                         ("Send Status Image", self.send_checked_status),
                         ("Add Store", self.add_store_prompt),
                         ("Open Folder", self.open_app_folder)):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            crow.addWidget(btn)
        crow.addStretch(1)
        tab_outer.addWidget(controls)

        search_box = QGroupBox("Search Inventory Audit Status")
        srow = QHBoxLayout(search_box)
        for label, attr, width in (("Any", "status_search_any", 100),
                                   ("District", "status_search_district", 85),
                                   ("Store", "status_search_store", 100),
                                   ("Status", "status_search_status", 85),
                                   ("Rep", "status_search_rep", 100)):
            srow.addWidget(QLabel(label + ":"))
            edit = QLineEdit()
            edit.setFixedWidth(width)
            edit.textChanged.connect(lambda _t: self.refresh_status_table())
            setattr(self, attr + "_edit", edit)
            srow.addWidget(edit)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self.clear_status_search)
        srow.addWidget(clear_btn)
        srow.addStretch(1)
        tab_outer.addWidget(search_box)

        legend = QLabel("Colors: Yellow = Pending, Blue = Completed after updated sheet load, Green = Completed and sent to WhatsApp.")
        tab_outer.addWidget(legend)

        self.status_tree = self._make_table(
            ["District", "Store", "Status", "Rep Name", "Checkbox"],
            [130, 200, 110, 170, 95])
        self._wire_sorting(self.status_tree)
        self.status_tree.itemClicked.connect(self.on_status_tree_click)
        tab_outer.addWidget(self.status_tree, 1)

    # ------------------------------------------------------------
    # TAB 2: Variance Audit
    # ------------------------------------------------------------
    def _build_audit_tab(self) -> None:
        tab_outer = QVBoxLayout(self.audit_tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)

        send_box = QGroupBox("Variance Audit Controls")
        grid = QGridLayout(send_box)
        # Row 0: filters
        grid.addWidget(QLabel("Send by:"), 0, 0)
        self.audit_send_mode_combo = QComboBox()
        self.audit_send_mode_combo.addItems(["District", "Store", "Sales Rep"])
        grid.addWidget(self.audit_send_mode_combo, 0, 1)
        grid.addWidget(QLabel("District filter:"), 0, 2)
        self.audit_district_combo = QComboBox()
        self.audit_district_combo.addItem("All Districts")
        self.audit_district_combo.activated.connect(lambda _i: self.on_audit_district_change())
        grid.addWidget(self.audit_district_combo, 0, 3)
        grid.addWidget(QLabel("Store filter:"), 0, 4)
        self.audit_store_combo = QComboBox()
        self.audit_store_combo.addItem("All Stores")
        self.audit_store_combo.activated.connect(lambda _i: self.on_audit_store_change())
        grid.addWidget(self.audit_store_combo, 0, 5)
        self.send_only_unsent_check = QCheckBox("Only unsent")
        self.send_only_unsent_check.setChecked(bool(SEND_ONLY_UNSENT_BY_DEFAULT))
        grid.addWidget(self.send_only_unsent_check, 0, 6)
        self.include_cleared_check = QCheckBox("Show cleared")
        self.include_cleared_check.toggled.connect(lambda _on: self.refresh_table())
        grid.addWidget(self.include_cleared_check, 0, 7)
        grid.setColumnStretch(8, 1)
        # Row 1: actions
        for col, (text, fn) in enumerate((
                ("Check Current Filter", self.auto_check_audit_rows),
                ("Clear Checkmarks", self.clear_audit_checkmarks),
                ("Send Checked Variance Image", self.send_checked_variances),
                ("Send Selected Image", self.send_selected),
                ("Send Pending", self.send_pending),
                ("Mark Cleared", lambda: self.mark_selected(True)),
                ("Mark Not Cleared", lambda: self.mark_selected(False)),
                ("Export Log", self.export_log),
                ("Open Folder", self.open_app_folder),
                ("Clear UI", self.clear_current_ui),
                ("Copy IMEI", self.copy_selected_imei)), start=0):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            grid.addWidget(btn, 1, col)
        tab_outer.addWidget(send_box)

        action_box = QGroupBox("Final Send Actions")
        arow = QHBoxLayout(action_box)
        arow.addWidget(QLabel("District:"))
        self.final_district_combo = QComboBox()
        self.final_district_combo.addItem("All Districts")
        arow.addWidget(self.final_district_combo)
        arow.addSpacing(12)
        start_btn = QPushButton("Send Starting Message")
        start_btn.clicked.connect(self.send_starting_message)
        arow.addWidget(start_btn)
        for n in (1, 2, 3):
            btn = QPushButton(f"Send Reminder {n}")
            btn.clicked.connect(lambda _c=False, num=n: self.send_reminder(num))
            arow.addWidget(btn)
        final_btn = QPushButton("Send Final District Result")
        final_btn.clicked.connect(self.send_final_district_result)
        arow.addWidget(final_btn)
        arow.addStretch(1)
        tab_outer.addWidget(action_box)

        search_box = QGroupBox("Search Variance Rows")
        srow = QHBoxLayout(search_box)
        for label, attr, width in (("Any", "audit_search_any", 100),
                                   ("District", "audit_search_district", 85),
                                   ("Store", "audit_search_store", 100),
                                   ("Product", "audit_search_product", 115),
                                   ("IMEI", "audit_search_imei", 100),
                                   ("Rep", "audit_search_rep", 85)):
            srow.addWidget(QLabel(label + ":"))
            edit = QLineEdit()
            edit.setFixedWidth(width)
            edit.textChanged.connect(lambda _t: self.refresh_table())
            setattr(self, attr + "_edit", edit)
            srow.addWidget(edit)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(self.clear_audit_search)
        srow.addWidget(clear_btn)
        srow.addStretch(1)
        tab_outer.addWidget(search_box)

        legend = QLabel("Colors: Yellow = Pending variance, Blue = Sent to WhatsApp, Green = Cleared.")
        tab_outer.addWidget(legend)

        self.audit_tree = self._make_table(
            ["District", "Store", "Product", "IMEI", "Status", "Rep Name", "Clearance", "Checkbox"],
            [110, 150, 280, 130, 100, 170, 110, 85], multi_select=True)
        self._wire_sorting(self.audit_tree)
        self.audit_tree.itemClicked.connect(self.on_audit_tree_click)
        self.audit_tree.itemDoubleClicked.connect(lambda _item: self.on_audit_double_click())
        self.audit_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.audit_tree.customContextMenuRequested.connect(lambda _pos: self.copy_selected_imei())
        sc_copy = QShortcut(QKeySequence("Ctrl+C"), self.audit_tree)
        sc_copy.activated.connect(self.copy_selected_imei)
        tab_outer.addWidget(self.audit_tree, 1)

    # ------------------------------------------------------------
    # TAB 3: Store List
    # ------------------------------------------------------------
    def _build_store_tab(self) -> None:
        tab_outer = QVBoxLayout(self.store_tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)

        form = QGroupBox("Add or Update Store")
        grid = QGridLayout(form)
        grid.addWidget(QLabel("District"), 0, 0)
        self._store_district_combo = QComboBox()
        self._store_district_combo.addItems(self.db.all_known_districts())
        grid.addWidget(self._store_district_combo, 0, 1)
        grid.addWidget(QLabel("Store"), 0, 2)
        self.store_name_edit = QLineEdit()
        self.store_name_edit.setFixedWidth(230)
        grid.addWidget(self.store_name_edit, 0, 3)
        for col, (text, fn) in enumerate((
                ("Save Store", self.save_store_account_from_form),
                ("Import XLSX", self.import_store_accounts_file),
                ("Delete Selected", self.delete_selected_store_account),
                ("Download Template", self.download_store_accounts_template),
                ("Clear Form", self.clear_store_form)), start=4):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            grid.addWidget(btn, 0, col)
        tab_outer.addWidget(form)

        store_search_box = QWidget()
        srow = QHBoxLayout(store_search_box)
        srow.setContentsMargins(0, 0, 0, 0)
        srow.addWidget(QLabel("Search:"))
        self.store_search_edit = QLineEdit()
        self.store_search_edit.setFixedWidth(290)
        self.store_search_edit.textChanged.connect(lambda _t: self.refresh_store_accounts_table())
        srow.addWidget(self.store_search_edit)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(lambda: (self.store_search_edit.clear(), self.refresh_store_accounts_table()))
        srow.addWidget(clear_btn)
        srow.addStretch(1)
        tab_outer.addWidget(store_search_box)

        self.store_tree = self._make_table(["District", "Store"], [220, 360])
        self._wire_sorting(self.store_tree)
        self.store_tree.itemSelectionChanged.connect(self.on_store_account_select)
        tab_outer.addWidget(self.store_tree, 1)

    # ------------------------------------------------------------
    # TAB 4: Employees
    # ------------------------------------------------------------
    def _build_rep_tab(self) -> None:
        tab_outer = QVBoxLayout(self.rep_tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)

        info = QGroupBox("About Employees")
        info_l = QVBoxLayout(info)
        info_lbl = QLabel(
            "Store each employee's full name and phone number here so WhatsApp reports can tag them.\n"
            "The 'Created By' field is optional — fill it in when the count file uses a short username "
            "(e.g. 'ArmanAli') that differs from the full Employee Name. "
            "The audit tries Created By first, then Employee Name, when resolving who to tag.\n"
            "Use 'Auto-detect Created By' to scan loaded variances and pre-fill the column automatically.")
        info_lbl.setWordWrap(True)
        info_l.addWidget(info_lbl)
        tab_outer.addWidget(info)

        form = QGroupBox("Add or Update Employee")
        grid = QGridLayout(form)
        grid.addWidget(QLabel("Employee Name:"), 0, 0)
        self.rep_name_edit = QLineEdit()
        self.rep_name_edit.setFixedWidth(190)
        grid.addWidget(self.rep_name_edit, 0, 1)
        grid.addWidget(QLabel("Phone:"), 0, 2)
        self.rep_phone_edit = QLineEdit()
        self.rep_phone_edit.setFixedWidth(150)
        grid.addWidget(self.rep_phone_edit, 0, 3)
        grid.addWidget(QLabel("Created By (optional):"), 0, 4)
        self.rep_created_by_edit = QLineEdit()
        self.rep_created_by_edit.setFixedWidth(150)
        grid.addWidget(self.rep_created_by_edit, 0, 5)
        grid.setColumnStretch(6, 1)
        btn_frame = QWidget()
        brow = QHBoxLayout(btn_frame)
        brow.setContentsMargins(0, 0, 0, 0)
        for text, fn in (("Save Employee", self.save_sales_rep_from_form),
                         ("Auto-detect Created By", self._rep_auto_detect_created_by),
                         ("Import XLSX", self.import_employees_file),
                         ("Download Template", self.download_employees_template),
                         ("Export Excel", self.export_employees_file),
                         ("Delete Selected", self.delete_selected_sales_rep),
                         ("Clear Form", self.clear_rep_form)):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            brow.addWidget(btn)
        brow.addStretch(1)
        grid.addWidget(btn_frame, 1, 0, 1, 7)
        tab_outer.addWidget(form)

        rep_search_box = QWidget()
        srow = QHBoxLayout(rep_search_box)
        srow.setContentsMargins(0, 0, 0, 0)
        srow.addWidget(QLabel("Search:"))
        self.rep_search_edit = QLineEdit()
        self.rep_search_edit.setFixedWidth(290)
        self.rep_search_edit.textChanged.connect(lambda _t: self.refresh_sales_reps_table())
        srow.addWidget(self.rep_search_edit)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(lambda: (self.rep_search_edit.clear(), self.refresh_sales_reps_table()))
        srow.addWidget(clear_btn)
        srow.addStretch(1)
        tab_outer.addWidget(rep_search_box)

        self.rep_tree = self._make_table(
            ["Employee Name", "Phone Number", "Created By (count username)"],
            [280, 180, 200])
        self._wire_sorting(self.rep_tree)
        self.rep_tree.itemSelectionChanged.connect(self.on_sales_rep_select)
        tab_outer.addWidget(self.rep_tree, 1)
    @staticmethod
    def _mix_hex(a: str, b: str, t: float) -> str:
        """Blend two #rrggbb colours (t=0 -> a, t=1 -> b)."""
        try:
            ar, ag, ab = (int(a.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
            br, bg, bb = (int(b.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4))
            return "#{:02x}{:02x}{:02x}".format(
                round(ar + (br - ar) * t),
                round(ag + (bg - ag) * t),
                round(ab + (bb - ab) * t))
        except Exception:
            return a

    # ------------------------------------------------------------
    # TAB 5: District DMs (WhatsApp send mode, group names, DM tags)
    # ------------------------------------------------------------
    def _build_dm_tab(self) -> None:
        tab_outer = QVBoxLayout(self.dm_tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)

        # WhatsApp send mode selector at top of tab
        mode_box = QGroupBox("WhatsApp Send Mode")
        mrow = QHBoxLayout(mode_box)
        mrow.addWidget(QLabel("Send reports via:"))
        self.wa_mode_var = self.db.get_setting("whatsapp_mode", "desktop")
        self.wa_desktop_radio = QRadioButton("WhatsApp Desktop App (pyautogui)")
        self.wa_web_radio = QRadioButton("WhatsApp Web (browser)")
        if self.wa_mode_var == "desktop":
            self.wa_desktop_radio.setChecked(True)
        elif self.wa_mode_var == "web":
            self.wa_web_radio.setChecked(True)
        self.wa_desktop_radio.toggled.connect(lambda on: on and setattr(self, "wa_mode_var", "desktop"))
        self.wa_web_radio.toggled.connect(lambda on: on and setattr(self, "wa_mode_var", "web"))
        mrow.addWidget(self.wa_desktop_radio)
        mrow.addSpacing(12)
        mrow.addWidget(self.wa_web_radio)
        mrow.addSpacing(12)
        save_mode_btn = QPushButton("Save Mode")
        save_mode_btn.clicked.connect(lambda: (self.db.save_setting("whatsapp_mode", self.wa_mode_var),
                                               self.set_status(f"WhatsApp mode saved: {self.wa_mode_var}")))
        mrow.addWidget(save_mode_btn)
        mrow.addStretch(1)
        tab_outer.addWidget(mode_box)

        # Side-by-side layout (Tk PanedWindow -> QSplitter)
        splitter = QSplitter(Qt.Horizontal)
        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(4, 4, 4, 4)
        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(4, 4, 4, 4)
        splitter.addWidget(left)
        splitter.addWidget(right)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 1)
        tab_outer.addWidget(splitter, 1)

        # ── Left: WhatsApp Group Names ──────────────────────────────────
        wg_form = QGroupBox("WhatsApp Group Names")
        wg_outer = QVBoxLayout(wg_form)
        wg_form_row = QWidget()
        wrow = QHBoxLayout(wg_form_row)
        wrow.setContentsMargins(0, 0, 0, 0)
        wrow.addWidget(QLabel("District:"))
        self._wg_district_combo = QComboBox()
        self._wg_district_combo.addItems(self.db.all_known_districts())
        wrow.addWidget(self._wg_district_combo)
        wrow.addSpacing(8)
        wrow.addWidget(QLabel("Group Name:"))
        self.wg_group_name_edit = QLineEdit()
        self.wg_group_name_edit.setFixedWidth(210)
        wrow.addWidget(self.wg_group_name_edit)
        wrow.addStretch(1)
        wg_outer.addWidget(wg_form_row)
        wg_btn_row = QWidget()
        wbrow = QHBoxLayout(wg_btn_row)
        wbrow.setContentsMargins(0, 0, 0, 0)
        for text, fn in (("Save", self.save_whatsapp_group_from_form),
                         ("Delete", self.delete_selected_whatsapp_group),
                         ("Clear", self.clear_wg_form)):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            wbrow.addWidget(btn)
        wbrow.addStretch(1)
        wg_outer.addWidget(wg_btn_row)
        lv.addWidget(wg_form)

        self.wg_tree = self._make_table(["District", "WhatsApp Group Name"], [160, 280])
        self._wire_sorting(self.wg_tree)
        self.wg_tree.itemSelectionChanged.connect(self.on_whatsapp_group_select)
        lv.addWidget(self.wg_tree, 1)

        # ── Right: District Manager WhatsApp Tags ───────────────────────
        form = QGroupBox("DM WhatsApp Tags")
        dm_outer = QVBoxLayout(form)
        dm_form_row = QWidget()
        drow = QHBoxLayout(dm_form_row)
        drow.setContentsMargins(0, 0, 0, 0)
        drow.addWidget(QLabel("District:"))
        self._dm_district_combo = QComboBox()
        self._dm_district_combo.addItems(self.db.all_known_districts())
        drow.addWidget(self._dm_district_combo)
        drow.addSpacing(8)
        drow.addWidget(QLabel("DM Name:"))
        self.dm_name_edit = QLineEdit()
        self.dm_name_edit.setFixedWidth(180)
        drow.addWidget(self.dm_name_edit)
        drow.addSpacing(8)
        drow.addWidget(QLabel("Phone:"))
        self.dm_phone_edit = QLineEdit()
        self.dm_phone_edit.setFixedWidth(150)
        drow.addWidget(self.dm_phone_edit)
        drow.addStretch(1)
        dm_outer.addWidget(dm_form_row)
        dm_btn_row = QWidget()
        dbrow = QHBoxLayout(dm_btn_row)
        dbrow.setContentsMargins(0, 0, 0, 0)
        for text, fn in (("Save", self.save_district_manager_from_form),
                         ("Delete", self.delete_selected_district_manager),
                         ("Clear", self.clear_dm_form)):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            dbrow.addWidget(btn)
        dbrow.addStretch(1)
        dm_outer.addWidget(dm_btn_row)
        rv.addWidget(form)

        self.dm_tree = self._make_table(["District", "DM Name", "Phone Number"], [160, 220, 180])
        self._wire_sorting(self.dm_tree)
        self.dm_tree.itemSelectionChanged.connect(self.on_district_manager_select)
        rv.addWidget(self.dm_tree, 1)

    def refresh_district_managers_table(self) -> None:
        if not hasattr(self, "dm_tree"):
            return
        self.dm_tree.setRowCount(0)
        for row in self.db.district_managers():
            r = self.dm_tree.rowCount()
            self.dm_tree.insertRow(r)
            values = (row.get("District", ""), row.get("DM Name", ""), row.get("Phone", ""))
            for col, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                self.dm_tree.setItem(r, col, item)
            self.dm_tree.item(r, 0).setData(Qt.UserRole, normalize_district(row.get("District", "")))

    def save_district_manager_from_form(self) -> None:
        try:
            self.db.save_district_manager(self._dm_district_combo.currentText(), self.dm_name_edit.text(), self.dm_phone_edit.text())
            self.refresh_district_managers_table()
            self.refresh_entry_form_districts()
            self.set_status(f"Saved district DM: {self._dm_district_combo.currentText()} | {self.dm_phone_edit.text()}")
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", str(exc))

    def delete_selected_district_manager(self) -> None:
        rows = {i.row() for i in self.dm_tree.selectedIndexes()} if hasattr(self, "dm_tree") else set()
        if not rows:
            QMessageBox.information(self, "No selection", "Select a district DM first.")
            return
        if QMessageBox.question(self, "Delete district DM", "Delete the selected district DM phone record?") != QMessageBox.Yes:
            return
        district = self.dm_tree.item(sorted(rows)[0], 0).data(Qt.UserRole)
        self.db.delete_district_manager(district)
        self.refresh_district_managers_table()
        self.refresh_entry_form_districts()
        self.clear_dm_form()
        self.set_status("Deleted selected district DM.")

    def clear_dm_form(self) -> None:
        self._combo_set(self._dm_district_combo, "")
        self.dm_name_edit.clear()
        self.dm_phone_edit.clear()

    def refresh_whatsapp_groups_table(self) -> None:
        if not hasattr(self, "wg_tree"):
            return
        self.wg_tree.setRowCount(0)
        for row in self.db.whatsapp_groups():
            r = self.wg_tree.rowCount()
            self.wg_tree.insertRow(r)
            values = (row.get("District", ""), row.get("Group Name", ""))
            for col, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                self.wg_tree.setItem(r, col, item)
            self.wg_tree.item(r, 0).setData(Qt.UserRole, normalize_district(row.get("District", "")))

    def save_whatsapp_group_from_form(self) -> None:
        try:
            self.db.save_whatsapp_group(self._wg_district_combo.currentText(), self.wg_group_name_edit.text())
            self.refresh_whatsapp_groups_table()
            self.refresh_entry_form_districts()
            self.set_status(f"Saved WhatsApp group: {self._wg_district_combo.currentText()} | {self.wg_group_name_edit.text()}")
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", str(exc))

    def delete_selected_whatsapp_group(self) -> None:
        rows = {i.row() for i in self.wg_tree.selectedIndexes()} if hasattr(self, "wg_tree") else set()
        if not rows:
            QMessageBox.information(self, "No selection", "Select a WhatsApp group first.")
            return
        if QMessageBox.question(self, "Delete WhatsApp group", "Delete the selected WhatsApp group record?") != QMessageBox.Yes:
            return
        district = self.wg_tree.item(sorted(rows)[0], 0).data(Qt.UserRole)
        self.db.delete_whatsapp_group(district)
        self.refresh_whatsapp_groups_table()
        self.refresh_entry_form_districts()
        self.clear_wg_form()
        self.set_status("Deleted selected WhatsApp group.")

    def clear_wg_form(self) -> None:
        self._combo_set(self._wg_district_combo, "")
        self.wg_group_name_edit.clear()

    def on_whatsapp_group_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.wg_tree.selectedIndexes()})
        if not rows:
            return
        row = rows[0]
        district = self.wg_tree.item(row, 0).text()
        for entry in self.db.whatsapp_groups():
            if normalize_district(entry.get("District", "")) == normalize_district(district):
                self._combo_set(self._wg_district_combo, entry.get("District", ""))
                self.wg_group_name_edit.setText(entry.get("Group Name", ""))
                return

    def on_district_manager_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.dm_tree.selectedIndexes()})
        if not rows:
            return
        row = rows[0]
        district = self.dm_tree.item(row, 0).text()
        for entry in self.db.district_managers():
            if normalize_district(entry.get("District", "")) == normalize_district(district):
                self._combo_set(self._dm_district_combo, entry.get("District", ""))
                self.dm_name_edit.setText(entry.get("DM Name", ""))
                self.dm_phone_edit.setText(entry.get("Phone", ""))
                return

    # ── Created By → Employee Mapping Tab ─────────────────────────────────
    # NOTE: exactly like the Tkinter build, this panel is implemented but
    # not wired into the notebook (self.mapping_tab is never created there
    # either). Ported verbatim-in-spirit for parity.
    def _build_mapping_tab(self) -> None:
        """Panel to map 'Created By' values from inventory count to
        Employee names + phone numbers from the new Timesheet file.

        This lets the audit correctly identify and WhatsApp the right
        person when the count's Created By username (e.g. 'ArmanAli')
        doesn't directly match the Timesheet Employee name
        (e.g. 'Arman Ali Mohammed').
        """
        # ── Instructions ───────────────────────────────────────────────
        info = QGroupBox("About This Panel")
        info_l = QVBoxLayout(info)
        info_lbl = QLabel(
            "The Inventory Count file uses a short 'Created By' username (e.g. 'ArmanAli'), "
            "while the Timesheet file uses the full 'Employee' name (e.g. 'Arman Ali Mohammed').\n"
            "Add a mapping here so the audit can look up the correct employee (and their phone) "
            "when sending WhatsApp reports.  Mappings are saved permanently in the database.")
        info_lbl.setWordWrap(True)
        info_l.addWidget(info_lbl)
        self.mapping_tab_layout = QVBoxLayout(self.mapping_tab)
        self.mapping_tab_layout.addWidget(info)

        # ── Auto-populate from loaded data ─────────────────────────────
        auto_frame = QWidget()
        arow = QHBoxLayout(auto_frame)
        arow.setContentsMargins(0, 0, 0, 0)
        auto_btn = QPushButton("Auto-detect Created By values from loaded variances")
        auto_btn.clicked.connect(self._mapping_auto_detect)
        arow.addWidget(auto_btn)
        refresh_btn = QPushButton("Refresh Table")
        refresh_btn.clicked.connect(self._mapping_refresh)
        arow.addWidget(refresh_btn)
        arow.addStretch(1)
        self.mapping_tab_layout.addWidget(auto_frame)

        # ── Add / Edit form ────────────────────────────────────────────
        form = QGroupBox("Add / Edit Mapping")
        grid = QGridLayout(form)
        grid.addWidget(QLabel("Created By (from count file):"), 0, 0)
        self.map_created_by_edit = QLineEdit()
        self.map_created_by_edit.setFixedWidth(180)
        grid.addWidget(self.map_created_by_edit, 0, 1)
        grid.addWidget(QLabel("Employee Name (from timesheet):"), 0, 2)
        self.map_employee_edit = QLineEdit()
        self.map_employee_edit.setFixedWidth(230)
        grid.addWidget(self.map_employee_edit, 0, 3)
        grid.addWidget(QLabel("Phone (WhatsApp):"), 0, 4)
        self.map_phone_edit = QLineEdit()
        self.map_phone_edit.setFixedWidth(150)
        grid.addWidget(self.map_phone_edit, 0, 5)
        btn_row = QWidget()
        brow = QHBoxLayout(btn_row)
        brow.setContentsMargins(0, 0, 0, 0)
        save_btn = QPushButton("Save Mapping")
        save_btn.clicked.connect(self._mapping_save)
        brow.addWidget(save_btn)
        del_btn = QPushButton("Delete Selected")
        del_btn.clicked.connect(self._mapping_delete)
        brow.addWidget(del_btn)
        clear_btn = QPushButton("Clear Form")
        clear_btn.clicked.connect(self._mapping_clear_form)
        brow.addWidget(clear_btn)
        brow.addStretch(1)
        grid.addWidget(btn_row, 1, 0, 1, 6)
        self.mapping_tab_layout.addWidget(form)

        # ── Table ──────────────────────────────────────────────────────
        tbl_frame = QGroupBox("Saved Mappings")
        tf = QVBoxLayout(tbl_frame)
        self.mapping_tree = self._make_table(
            ["Created By (count file)", "Employee Name (timesheet)", "Phone (WhatsApp)"],
            [200, 280, 160])
        self.mapping_tree.itemSelectionChanged.connect(self._mapping_on_select)
        tf.addWidget(self.mapping_tree)
        self.mapping_tab_layout.addWidget(tbl_frame, 1)
        self._mapping_refresh()

    def _mapping_refresh(self) -> None:
        self.mapping_tree.setRowCount(0)
        for m in self.db.get_created_by_mappings():
            r = self.mapping_tree.rowCount()
            self.mapping_tree.insertRow(r)
            for col, value in enumerate((m["created_by"], m["employee_name"], m["phone"])):
                self.mapping_tree.setItem(r, col, QTableWidgetItem(str(value)))

    def _mapping_on_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.mapping_tree.selectedIndexes()})
        if not rows:
            return
        row = rows[0]
        self.map_created_by_edit.setText(self.mapping_tree.item(row, 0).text())
        self.map_employee_edit.setText(self.mapping_tree.item(row, 1).text())
        phone_item = self.mapping_tree.item(row, 2)
        self.map_phone_edit.setText(phone_item.text() if phone_item is not None else "")

    def _mapping_save(self) -> None:
        cb = self.map_created_by_edit.text().strip()
        em = self.map_employee_edit.text().strip()
        ph = self.map_phone_edit.text().strip()
        if not cb:
            QMessageBox.warning(self, "Missing Value", "Enter the 'Created By' value from the count file.")
            return
        self.db.upsert_created_by_mapping(cb, em, ph)
        self._mapping_refresh()
        self._mapping_clear_form()
        self.set_status(f"Mapping saved: '{cb}' → '{em}'")

    def _mapping_delete(self) -> None:
        rows = sorted({i.row() for i in self.mapping_tree.selectedIndexes()})
        if not rows:
            QMessageBox.warning(self, "No Selection", "Select a row to delete.")
            return
        cb = self.mapping_tree.item(rows[0], 0).text()
        if QMessageBox.question(self, "Confirm Delete", f"Delete mapping for '{cb}'?") == QMessageBox.Yes:
            self.db.delete_created_by_mapping(cb)
            self._mapping_refresh()
            self.set_status(f"Deleted mapping for '{cb}'")

    def _mapping_clear_form(self) -> None:
        self.map_created_by_edit.clear()
        self.map_employee_edit.clear()
        self.map_phone_edit.clear()

    def _mapping_auto_detect(self) -> None:
        """Scan all loaded variance rows for unique Created By values
        and add any not already mapped (leaving employee/phone blank
        for the user to fill in)."""
        if not self.data_loaded:
            QMessageBox.warning(self, "No Data", "Load variances first, then auto-detect.")
            return
        existing = {m["created_by"] for m in self.db.get_created_by_mappings()}
        added = 0
        try:
            with self.db.connect() as con:
                rows = con.execute(
                    "SELECT DISTINCT created_by FROM variances WHERE created_by IS NOT NULL AND created_by != ''"
                ).fetchall()
            for (cb,) in rows:
                if cb and cb not in existing:
                    self.db.upsert_created_by_mapping(cb, "", "")
                    added += 1
            self._mapping_refresh()
            self.set_status(f"Auto-detected {added} new Created By value(s). Fill in Employee Name and Phone.")
            if added == 0:
                QMessageBox.information(self, "Auto-detect", "All Created By values already have mappings (or none found in loaded data).")
        except Exception as exc:
            QMessageBox.critical(self, "Error", str(exc))

    # ------------------------------------------------------------
    # TAB 6: Excluded Devices
    # ------------------------------------------------------------
    def _build_exclusion_tab(self) -> None:
        tab_outer = QVBoxLayout(self.exclusion_tab)
        tab_outer.setContentsMargins(10, 10, 10, 10)

        form = QGroupBox("Exclude Devices From Variance Images")
        grid = QGridLayout(form)
        grid.addWidget(QLabel("District"), 0, 0)
        self._exclusion_district_combo = QComboBox()
        self._exclusion_district_combo.addItems(self.db.all_known_districts())
        grid.addWidget(self._exclusion_district_combo, 0, 1)
        grid.addWidget(QLabel("Product"), 0, 2)
        self.exclusion_product_edit = QLineEdit()
        self.exclusion_product_edit.setFixedWidth(240)
        grid.addWidget(self.exclusion_product_edit, 0, 3)
        grid.addWidget(QLabel("IMEI"), 0, 4)
        self.exclusion_imei_edit = QLineEdit()
        self.exclusion_imei_edit.setFixedWidth(180)
        self.exclusion_imei_edit.editingFinished.connect(self.lookup_product_for_exclusion_imei)
        grid.addWidget(self.exclusion_imei_edit, 0, 5)
        grid.addWidget(QLabel("Comments"), 0, 6)
        self.exclusion_comments_edit = QLineEdit()
        self.exclusion_comments_edit.setFixedWidth(280)
        grid.addWidget(self.exclusion_comments_edit, 0, 7)
        btn_row = QWidget()
        brow = QHBoxLayout(btn_row)
        brow.setContentsMargins(0, 0, 0, 0)
        for text, fn in (("Find Product", self.lookup_product_for_exclusion_imei),
                         ("Save", self.save_device_exclusion_from_form),
                         ("Update Comment", self.update_selected_exclusion_comment),
                         ("Import XLSX", self.import_excluded_imeis_file),
                         ("Export Excel", self.export_excluded_imeis_file),
                         ("Download Template", self.download_excluded_imeis_template),
                         ("Delete", self.delete_selected_device_exclusion),
                         ("Delete All", self.delete_all_device_exclusions_from_ui),
                         ("Clear", self.clear_device_exclusion_form)):
            btn = QPushButton(text)
            btn.clicked.connect(fn)
            brow.addWidget(btn)
        brow.addStretch(1)
        grid.addWidget(btn_row, 1, 0, 1, 8)
        tab_outer.addWidget(form)

        # Search bar for Excluded Devices
        exclusion_search_box = QWidget()
        srow = QHBoxLayout(exclusion_search_box)
        srow.setContentsMargins(0, 0, 0, 0)
        srow.addWidget(QLabel("Search:"))
        self.exclusion_search_edit = QLineEdit()
        self.exclusion_search_edit.setFixedWidth(290)
        self.exclusion_search_edit.textChanged.connect(lambda _t: self.refresh_device_exclusions_table())
        srow.addWidget(self.exclusion_search_edit)
        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(lambda: (self.exclusion_search_edit.clear(), self.refresh_device_exclusions_table()))
        srow.addWidget(clear_btn)
        srow.addStretch(1)
        tab_outer.addWidget(exclusion_search_box)

        self.exclusion_tree = self._make_table(
            ["District", "Product", "IMEI", "Comments"], [160, 400, 220, 420])
        self._wire_sorting(self.exclusion_tree)
        self.exclusion_tree.itemSelectionChanged.connect(self.on_device_exclusion_select)
        self.exclusion_tree.itemDoubleClicked.connect(lambda _item: self.edit_selected_exclusion_comment())
        tab_outer.addWidget(self.exclusion_tree, 1)

    def delete_all_device_exclusions_from_ui(self) -> None:
        if QMessageBox.question(
                self, "Delete All Exclusions",
                "Are you sure you want to completely remove ALL device exclusions from the database? This cannot be undone.") != QMessageBox.Yes:
            return
        self.db.delete_all_device_exclusions()
        self.refresh_device_exclusions_table()
        self.refresh_entry_form_districts()
        self.refresh_table()
        self.clear_device_exclusion_form()
        self.set_status("Deleted all device exclusions.")

    def refresh_device_exclusions_table(self) -> None:
        if not hasattr(self, "exclusion_tree"):
            return
        self.exclusion_tree.setRowCount(0)
        search = self.exclusion_search_edit.text().strip().lower() if hasattr(self, "exclusion_search_edit") else ""
        for row in self.db.device_exclusions():
            district = row.get("District", "")
            product = row.get("Product", "")
            imei = row.get("IMEI", "")
            comments = row.get("Comments", "")
            if search and search not in district.lower() and search not in product.lower() and search not in imei.lower() and search not in comments.lower():
                continue
            r = self.exclusion_tree.rowCount()
            self.exclusion_tree.insertRow(r)
            values = (district, product, imei, comments)
            for col, value in enumerate(values):
                self.exclusion_tree.setItem(r, col, QTableWidgetItem(str(value)))
            iid = row.get("RuleKey", device_rule_key(district + "|" + product + "|" + imei))
            self.exclusion_tree.item(r, 0).setData(Qt.UserRole, iid)

    def product_for_imei_from_loaded_rows(self, imei: str) -> str:
        imei_key = device_rule_key(imei)
        if not imei_key:
            return ""

        # Prefer current loaded variance rows.
        for row in self.db.get_rows_by_keys(self.loaded_keys) if self.loaded_keys else []:
            if device_rule_key(row.imei) == imei_key:
                return safe_text(row.product)

        # Fallback to currently read inventory records if available.
        for rec in self.current_inventory_records:
            for key, value in rec.items():
                if device_rule_key(value) == imei_key:
                    product_col = find_column(rec, ["Product", "Product Name", "Device", "Model", "Item", "SKU Description"])
                    if product_col:
                        return safe_text(rec.get(product_col, ""))
        return ""

    def lookup_product_for_exclusion_imei(self, _event=None):
        imei = self.exclusion_imei_edit.text()
        if not safe_text(imei):
            return None
        product = self.product_for_imei_from_loaded_rows(imei)
        if product:
            self.exclusion_product_edit.setText(product)
            self.set_status(f"Product found for IMEI {imei}: {product}")
        else:
            self.set_status(f"No product found for IMEI {imei}. Product field left unchanged.")
        return None

    def save_device_exclusion_from_form(self) -> None:
        try:
            district = self._exclusion_district_combo.currentText()
            imei = self.exclusion_imei_edit.text()
            product = self.exclusion_product_edit.text()
            comments = self.exclusion_comments_edit.text()

            if safe_text(imei) and not safe_text(product):
                found_product = self.product_for_imei_from_loaded_rows(imei)
                if found_product:
                    product = found_product
                    self.exclusion_product_edit.setText(found_product)

            # IMEI-ONLY exclusions (user rule 2026-09-15): the IMEI is the
            # exclusion — whole products cannot be excluded. A product-only
            # entry is rejected and the app asks for the IMEI(s) instead.
            if not safe_text(imei):
                raise ValueError(
                    "IMEI is required — whole products cannot be excluded. "
                    "Enter the IMEI(s) of the specific device(s) to exclude."
                )
            # IMEI exclusions stay global and ignore district.
            district = ""
            self._combo_set(self._exclusion_district_combo, "")

            selected_key = self.selected_exclusion_key.strip()
            if selected_key:
                self.db.delete_device_exclusion(selected_key)

            self.db.save_device_exclusion(product, imei, comments, district)
            self.selected_exclusion_key = ""
            self.refresh_device_exclusions_table()
            self.refresh_entry_form_districts()
            self.refresh_table()
            self.set_status(
                f"Saved exclusion. District: {safe_text(district) or '-'} | Product: {safe_text(product) or '-'} | IMEI: {safe_text(imei) or '-'} | Comments: {safe_text(comments) or '-'}"
            )
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", str(exc))

    def update_selected_exclusion_comment(self) -> None:
        rows = sorted({i.row() for i in self.exclusion_tree.selectedIndexes()}) if hasattr(self, "exclusion_tree") else []
        if not rows:
            QMessageBox.information(self, "No selection", "Select an exclusion row first.")
            return
        rule_key = self.exclusion_tree.item(rows[0], 0).data(Qt.UserRole)
        try:
            self.db.update_device_exclusion_comment(rule_key, self.exclusion_comments_edit.text())
            self.selected_exclusion_key = rule_key
            self.refresh_device_exclusions_table()
            self.refresh_table()
            self.set_status("Updated selected exclusion comment.")
        except Exception as exc:
            QMessageBox.critical(self, "Update failed", str(exc))

    def edit_selected_exclusion_comment(self, _event=None):
        rows = sorted({i.row() for i in self.exclusion_tree.selectedIndexes()}) if hasattr(self, "exclusion_tree") else []
        if not rows:
            return None

        current_comment = self.exclusion_comments_edit.text()
        new_comment, ok = QInputDialog.getText(
            self,
            "Edit Comments",
            "Enter comments for selected exclusion.",
            text=current_comment,
        )
        if not ok:
            return None

        self.exclusion_comments_edit.setText(new_comment)
        self.update_selected_exclusion_comment()
        return None

    def import_excluded_imeis_file(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self,
            "Import Excluded Devices XLSX",
            "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)",
        )
        if not path:
            return

        try:
            records: List[Dict[str, str]] = []
            suffix = Path(path).suffix.lower()
            if suffix in {".xlsx", ".xlsm", ".xls"}:
                records = read_xlsx_records(path)
            else:
                QMessageBox.critical(self, "Unsupported file", "Please select an Excel file.")
                return

            if not records:
                QMessageBox.critical(self, "Empty file", "No exclusion rows found in the selected file.")
                return

            sample = records[0]
            district_col = find_column(sample, ["District", "Market", "Area"])
            product_col = find_column(sample, ["Product", "Product Name", "Device", "Model", "Item"])
            imei_col = find_column(sample, ["IMEI", "IMEI Number", "ESN", "ESN Number", "Serial", "Serial 1", "Device ID"])
            comments_col = find_column(sample, ["Comments", "Comment", "Notes", "Note", "Reason"])

            if not imei_col:
                QMessageBox.critical(
                    self,
                    "Missing columns",
                    "Import file needs an IMEI column — exclusions are IMEI-level only. Supported headers: District, Product, IMEI, Comments.",
                )
                return

            imported = 0
            skipped = 0
            for rec in records:
                district = normalize_district(rec.get(district_col, "")) if district_col else ""
                product = safe_text(rec.get(product_col, "")) if product_col else ""
                imei = safe_text(rec.get(imei_col, "")) if imei_col else ""
                comments = safe_text(rec.get(comments_col, "")) if comments_col else ""

                if imei and not product:
                    product = self.product_for_imei_from_loaded_rows(imei)

                # IMEI-ONLY exclusions (user rule 2026-09-15): rows without
                # an IMEI cannot be imported — whole products cannot be
                # excluded. The app asks for the IMEI(s) instead.
                if not imei:
                    skipped += 1
                    continue

                # IMEI exclusions ignore district.
                district = ""

                self.db.save_device_exclusion(product, imei, comments, district)
                imported += 1

            self.refresh_device_exclusions_table()
            self.refresh_table()
            QMessageBox.information(self, "Import complete", f"Imported {imported} exclusion row(s).\nSkipped {skipped} invalid row(s).\n\nExclusions are IMEI-level only — whole products cannot be excluded. Provide the IMEI(s) of any device that must be hidden.")
            self.set_status(f"Imported excluded devices. Imported: {imported}. Skipped: {skipped}.")
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(self, "Import failed", str(exc))
            self.set_status("Excluded devices import failed.")

    def export_excluded_imeis_file(self) -> None:
        rows = self.db.device_exclusions()
        if not rows:
            QMessageBox.information(self, "No exclusions", "No excluded devices are saved.")
            return
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"GFH_Excluded_Devices_{stamp}.xlsx"
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        path, _selected = QFileDialog.getSaveFileName(
            self, "Export Excluded Devices", str(EXPORT_DIR / default_name),
            "Excel Files (*.xlsx)")
        if not path:
            return
        output_path = Path(path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["District", "Product", "IMEI", "Comments"])
        for r_idx, row in enumerate(rows, start=2):
            ws.cell(row=r_idx, column=1, value=row.get("District", ""))
            ws.cell(row=r_idx, column=2, value=row.get("Product", ""))

            imei_cell = ws.cell(row=r_idx, column=3, value=row.get("IMEI", ""))
            imei_cell.data_type = 's'
            imei_cell.number_format = '@'

            ws.cell(row=r_idx, column=4, value=row.get("Comments", ""))
        wb.save(output_path)

        self.set_status(f"Exported excluded devices: {output_path}")
        QMessageBox.information(self, "Export complete", f"Excluded devices exported:\n{output_path}")

    def delete_selected_device_exclusion(self) -> None:
        rows = sorted({i.row() for i in self.exclusion_tree.selectedIndexes()}) if hasattr(self, "exclusion_tree") else []
        if not rows:
            QMessageBox.information(self, "No selection", "Select a device exclusion first.")
            return
        if QMessageBox.question(self, "Delete exclusion", "Delete the selected device exclusion from the database?") != QMessageBox.Yes:
            return
        rule_key = self.exclusion_tree.item(rows[0], 0).data(Qt.UserRole)
        self.db.delete_device_exclusion(rule_key)
        self.refresh_device_exclusions_table()
        self.refresh_entry_form_districts()
        self.refresh_table()
        self.clear_device_exclusion_form()
        self.set_status("Deleted selected device exclusion.")

    def clear_device_exclusion_form(self) -> None:
        self._combo_set(self._exclusion_district_combo, "")
        self.exclusion_product_edit.clear()
        self.exclusion_imei_edit.clear()
        self.exclusion_comments_edit.clear()
        self.selected_exclusion_key = ""

    def on_device_exclusion_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.exclusion_tree.selectedIndexes()}) if hasattr(self, "exclusion_tree") else []
        if not rows:
            return
        rule_key = self.exclusion_tree.item(rows[0], 0).data(Qt.UserRole)
        self.selected_exclusion_key = rule_key
        for row in self.db.device_exclusions():
            if row.get("RuleKey") == rule_key:
                self._combo_set(self._exclusion_district_combo, row.get("District", ""))
                self.exclusion_product_edit.setText(row.get("Product", ""))
                self.exclusion_imei_edit.setText(row.get("IMEI", ""))
                self.exclusion_comments_edit.setText(row.get("Comments", ""))
                return

    def _rep_auto_detect_created_by(self) -> None:
        """Match 'Created By' usernames from the count file to full Employee
        names from the timesheet by finding who worked at the same Store.

        Example: Count file has 'BabarAli' at 'Mt View Store'.
                 Timesheet has 'Babar Ali' at 'Mt View Store'.
                 → auto-maps 'BabarAli' → 'Babar Ali'.
        """
        inv_records = getattr(self, "current_inventory_records", [])
        ts_records = getattr(self, "current_time_sheet_records", [])

        if not inv_records:
            QMessageBox.information(self, "No Data", "Load an Inventory Count file first (Load Variances).")
            return

        # Build store → employee map from timesheet
        sample_ts = ts_records[0] if ts_records else {}
        store_col_ts = find_column(sample_ts, ["Store"])
        emp_col_ts = find_column(sample_ts, ["Employee", "Employee Name"])
        email_col_ts = find_column(sample_ts, ["Email", "Email Address"])

        store_to_employee: Dict[str, str] = {}  # norm_store → full employee name
        store_to_email: Dict[str, str] = {}

        for rec in ts_records:
            emp_raw = safe_text(rec.get(emp_col_ts, "")) if emp_col_ts else ""
            if not emp_raw or "TOTAL" in emp_raw.upper() or "— " in emp_raw:
                continue
            store_raw = rec.get(store_col_ts, "") if store_col_ts else ""
            norm = normalize_store(store_raw)
            if norm:
                store_to_employee[norm] = emp_raw.strip()
                if email_col_ts:
                    store_to_email[norm] = safe_text(rec.get(email_col_ts, ""))

        # Build store → created_by map from count file
        sample_inv = inv_records[0] if inv_records else {}
        store_col_inv = find_column(sample_inv, ["Store"])
        cb_col = find_column(sample_inv, ["Created By", "Count By", "User Login"])

        store_to_created_by: Dict[str, str] = {}
        for rec in inv_records:
            store_raw = rec.get(store_col_inv, "") if store_col_inv else ""
            norm = normalize_store(store_raw)
            cb = safe_text(rec.get(cb_col, "")).strip() if cb_col else ""
            if norm and cb:
                store_to_created_by[norm] = cb

        # Match: same store = same person
        existing = {m["created_by"]: m for m in self.db.get_created_by_mappings()}
        added = updated = 0

        for norm_store, cb in store_to_created_by.items():
            emp_name = store_to_employee.get(norm_store, "")
            email = store_to_email.get(norm_store, "")
            if cb in existing:
                # Update employee name if now known and not already set
                if emp_name and not existing[cb]["employee_name"]:
                    self.db.upsert_created_by_mapping(cb, emp_name, existing[cb]["phone"])
                    updated += 1
            else:
                self.db.upsert_created_by_mapping(cb, emp_name, "")
                added += 1

        self.refresh_sales_reps_table()

        if added + updated == 0 and not store_to_created_by:
            QMessageBox.information(self, "Auto-detect", "No 'Created By' values found in the count file.")
            return

        msg = []
        if added:
            msg.append(f"{added} new mapping(s) added")
        if updated:
            msg.append(f"{updated} mapping(s) updated with employee name")

        summary_lines = []
        for norm_store, cb in sorted(store_to_created_by.items()):
            emp = store_to_employee.get(norm_store, "— not found in timesheet")
            summary_lines.append(f"  {cb}  →  {emp}  ({display_store(norm_store)})")

        self.set_status(f"Auto-detect: {', '.join(msg) if msg else 'all already mapped'}.")
        QMessageBox.information(
            self,
            "Auto-detect Complete",
            (f"{', '.join(msg)}.\n\n" if msg else "All already mapped.\n\n") +
            "Matches found:\n" + "\n".join(summary_lines[:20]) +
            ("\n..." if len(summary_lines) > 20 else "") +
            "\n\nFill in Phone numbers in the Employees tab, then Save.",
        )

    # ── Download import templates (extractor pattern) ─────────────────────
    def _save_import_template_xlsx(self, sheet_title, initial_file, headers,
                                   example_rows, column_hint, done_note):
        """Builds and saves one import template workbook with the import
        columns already defined in the header row (LANCZOS-free, openpyxl)."""
        try:
            import openpyxl
            from openpyxl.styles import Font, PatternFill
            from openpyxl.utils import get_column_letter
        except ImportError:
            QMessageBox.critical(
                self,
                "Missing Dependency",
                "openpyxl is required to build the template.\n\nInstall it with:\n  pip install openpyxl",
            )
            return
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        save_path, _selected = QFileDialog.getSaveFileName(
            self,
            "Save import template",
            str(EXPORT_DIR / initial_file),
            "Excel file (*.xlsx)",
        )
        if not save_path:
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = sheet_title
            header_fill = PatternFill("solid", fgColor="090D26")
            for c, name in enumerate(headers, start=1):
                cell = ws.cell(row=1, column=c, value=name)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
                ws.column_dimensions[get_column_letter(c)].width = 28
            for row in example_rows:
                ws.append(row)
            wb.save(save_path)
        except Exception as exc:
            QMessageBox.critical(self, "Template Failed", f"Could not create the template:\n\n{exc}")
            return
        QMessageBox.information(
            self,
            "Template Ready",
            f"Template saved:\n{save_path}\n\n"
            f"Columns: {column_hint}\n\n"
            f"{done_note}",
        )
        try:
            self.log(f"Import template saved: {save_path}")
        except Exception:
            pass

    def import_store_accounts_file(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self, "Import Store List XLSX", "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)")
        if not path:
            return

        try:
            records: List[Dict[str, str]] = []
            suffix = Path(path).suffix.lower()

            if suffix in {".xlsx", ".xlsm", ".xls"}:
                records = read_xlsx_records(path)
            else:
                QMessageBox.critical(self, "Unsupported file", "Please select an Excel file.")
                return

            if not records:
                QMessageBox.critical(self, "Empty file", "No store rows found in the selected file.")
                return

            sample = records[0]
            district_col = find_column(sample, ["District", "Distrcit"])
            store_col = find_column(sample, ["Store", "Store Name", "Location"])

            missing = []
            if not district_col:
                missing.append("District")
            if not store_col:
                missing.append("Store")

            if missing:
                QMessageBox.critical(
                    self,
                    "Missing columns",
                    "Missing required column(s): " + ", ".join(missing) +
                    "\n\nRequired headers: District, Store",
                )
                return

            parsed_records: List[Dict[str, str]] = []
            skipped = 0
            for rec in records:
                district = normalize_district(rec.get(district_col, ""))
                store = display_store(rec.get(store_col, ""))

                if not store or not district or district == "Unknown":
                    skipped += 1
                    continue

                parsed_records.append({"District": district, "Store": store})

            imported = self.db.replace_store_accounts(parsed_records)

            self.master_store_records = self.db.store_master_records()
            self.refresh_store_accounts_table()
            self.refresh_entry_form_districts()
            self.populate_status_filters()
            self.populate_audit_filters()

            if self.data_loaded:
                source_name = os.path.basename(self.inventory_edit.text().strip()) if self.inventory_edit.text().strip() else ""
                self.status_rows, status_summary = build_inventory_status_rows(
                    self.current_inventory_records,
                    self.current_time_sheet_records,
                    master_store_records=self.master_store_records,
                    source_file=source_name,
                )
                self.status_row_by_key = {r.key: r for r in self.status_rows}
                self.db.upsert_inventory_status_rows(self.status_rows)
                self.refresh_status_table()

            QMessageBox.information(
                self,
                "Import complete",
                f"Replaced saved store list with {imported} row(s).\nSkipped {skipped} invalid row(s).\n\nThis list will stay saved until the next upload.",
            )
            self.set_status(f"Store list replaced from file. Saved: {imported}. Skipped: {skipped}.")
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(self, "Import failed", str(exc))
            self.set_status("Store list import failed.")

    def refresh_store_accounts_table(self) -> None:
        if not hasattr(self, "store_tree"):
            return
        self.store_tree.setRowCount(0)
        search = self.store_search_edit.text().strip().lower() if hasattr(self, "store_search_edit") else ""
        for row in self.db.store_accounts():
            district = row.get("District", "")
            store = row.get("Store", "")
            if search and search not in district.lower() and search not in store.lower():
                continue
            iid = row.get("StoreKey", normalize_store(store))
            r = self.store_tree.rowCount()
            self.store_tree.insertRow(r)
            for col, value in enumerate((district, store)):
                self.store_tree.setItem(r, col, QTableWidgetItem(str(value)))
            self.store_tree.item(r, 0).setData(Qt.UserRole, iid)

    def refresh_entry_form_districts(self) -> None:
        """Refresh all entry-form district comboboxes from the DB."""
        districts = self.db.all_known_districts()
        for combo in (self._store_district_combo, self._wg_district_combo, self._dm_district_combo, self._exclusion_district_combo):
            try:
                current = combo.currentText()
                combo.blockSignals(True)
                combo.clear()
                combo.addItems(districts)
                self._combo_set(combo, current)
                combo.blockSignals(False)
            except Exception:
                pass
        # Also refresh the audit and status filter combos
        if hasattr(self, 'audit_district_combo'):
            self.populate_audit_filters()
        if hasattr(self, 'status_district_combo'):
            self.populate_status_filters()

    def import_employees_file(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self, "Import Employees XLSX", "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)")
        if not path:
            return

        try:
            records: List[Dict[str, str]] = []
            suffix = Path(path).suffix.lower()

            if suffix in {".xlsx", ".xlsm", ".xls"}:
                records = read_xlsx_records(path)
            else:
                QMessageBox.critical(self, "Unsupported file", "Please select an Excel file.")
                return

            if not records:
                QMessageBox.critical(self, "Empty file", "No employee rows found in the selected file.")
                return

            sample = records[0]
            name_col = find_column(sample, [
                "Employee Name",
                "Employee",
                "Sales Rep Name",
                "Sales Rep",
                "Rep Name",
                "Salesperson",
                "Name",
                "Full Name",
            ])
            phone_col = find_column(sample, [
                "Phone Number",
                "Phone",
                "Mobile",
                "Mobile Number",
                "Contact",
                "Contact Number",
                "Number",
                "Whatsapp",
                "WhatsApp Number",
            ])

            missing = []
            if not name_col:
                missing.append("Employee Name")
            if not phone_col:
                missing.append("Phone Number")

            if missing:
                QMessageBox.critical(
                    self,
                    "Missing columns",
                    "Missing required column(s): " + ", ".join(missing) +
                    "\n\nRequired headers: Employee Name, Phone Number",
                )
                return

            imported = 0
            skipped = 0
            for rec in records:
                name = safe_text(rec.get(name_col, ""))
                phone = normalize_phone(rec.get(phone_col, ""))

                if not name or not phone:
                    skipped += 1
                    continue

                self.db.save_sales_rep(name, phone)
                imported += 1

            self.refresh_sales_reps_table()
            QMessageBox.information(
                self,
                "Import complete",
                f"Imported or updated {imported} employee phone row(s).\nSkipped {skipped} invalid row(s).\n\nEmployees stay saved until you delete them.",
            )
            self.set_status(f"Imported employees from file. Imported: {imported}. Skipped: {skipped}.")
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(self, "Import failed", str(exc))
            self.set_status("Employee import failed.")

    def export_employees_file(self) -> None:
        rows = self.db.sales_reps()
        if not rows:
            QMessageBox.information(self, "No employees", "No employees are saved in the database yet.")
            return

        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"GFH_Employees_{stamp}.xlsx"
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        path, _selected = QFileDialog.getSaveFileName(
            self, "Export Employees Excel", str(EXPORT_DIR / default_name),
            "Excel Files (*.xlsx)")
        if not path:
            return

        output_path = Path(path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Employee Name", "Phone Number"])
        for r_idx, row in enumerate(rows, start=2):
            ws.cell(row=r_idx, column=1, value=row.get("Rep Name", ""))
            ws.cell(row=r_idx, column=2, value=row.get("Phone", ""))
        wb.save(output_path)

        self.set_status(f"Exported employee list: {output_path}")
        QMessageBox.information(self, "Export complete", f"Employee list exported:\n{output_path}")

    def refresh_sales_reps_table(self) -> None:
        if not hasattr(self, "rep_tree"):
            return
        self.rep_tree.setRowCount(0)
        search = self.rep_search_edit.text().strip().lower() if hasattr(self, "rep_search_edit") else ""
        # Show standard sales_reps rows
        for row in self.db.sales_reps():
            rep_name = row.get("Rep Name", "")
            phone = row.get("Phone", "")
            created_by = ""
            if search and search not in rep_name.lower() and search not in phone.lower():
                continue
            iid = row.get("RepKey", person_name_key(rep_name))
            r = self.rep_tree.rowCount()
            self.rep_tree.insertRow(r)
            for col, value in enumerate((rep_name, phone, created_by)):
                self.rep_tree.setItem(r, col, QTableWidgetItem(str(value)))
            self.rep_tree.item(r, 0).setData(Qt.UserRole, iid)
        # Also show created_by mappings that don't have a matching sales_rep entry
        existing_names = {row.get("Rep Name", "").lower() for row in self.db.sales_reps()}
        for m in self.db.get_created_by_mappings():
            if m["employee_name"].lower() in existing_names:
                # Already shown above — just update its created_by column
                for r in range(self.rep_tree.rowCount()):
                    name_item = self.rep_tree.item(r, 0)
                    if name_item is not None and name_item.text().lower() == m["employee_name"].lower():
                        self.rep_tree.setItem(r, 2, QTableWidgetItem(m["created_by"]))
                        break
                continue
            if search and search not in m["created_by"].lower() and search not in m["employee_name"].lower():
                continue
            r = self.rep_tree.rowCount()
            self.rep_tree.insertRow(r)
            for col, value in enumerate((m["employee_name"], m["phone"], m["created_by"])):
                self.rep_tree.setItem(r, col, QTableWidgetItem(str(value)))

    def save_store_account_from_form(self) -> None:
        try:
            self.db.save_store_account(
                self._store_district_combo.currentText(),
                self.store_name_edit.text(),
                "",
                "",
                "",
            )
            self.master_store_records = self.db.store_master_records()
            self.refresh_store_accounts_table()
            self.refresh_entry_form_districts()
            self.populate_status_filters()
            self.populate_audit_filters()
            if self.data_loaded:
                source_name = os.path.basename(self.inventory_edit.text().strip()) if self.inventory_edit.text().strip() else ""
                self.status_rows, status_summary = build_inventory_status_rows(
                    self.current_inventory_records,
                    self.current_time_sheet_records,
                    master_store_records=self.master_store_records,
                    source_file=source_name,
                )
                self.status_row_by_key = {r.key: r for r in self.status_rows}
                self.db.upsert_inventory_status_rows(self.status_rows)
                self.refresh_status_table()
            self.set_status(f"Saved store: {self._store_district_combo.currentText()} | {self.store_name_edit.text()}")
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", str(exc))

    def delete_selected_store_account(self) -> None:
        rows = sorted({i.row() for i in self.store_tree.selectedIndexes()}) if hasattr(self, "store_tree") else []
        if not rows:
            QMessageBox.information(self, "No selection", "Select a store first.")
            return
        if QMessageBox.question(self, "Delete store", "Delete the selected store?") != QMessageBox.Yes:
            return
        store_key = self.store_tree.item(rows[0], 0).data(Qt.UserRole)
        self.db.delete_store_account(store_key)
        self.master_store_records = self.db.store_master_records()
        self.refresh_store_accounts_table()
        self.refresh_entry_form_districts()
        self.clear_store_form()
        self.set_status("Deleted selected store.")

    def clear_store_form(self) -> None:
        self._combo_set(self._store_district_combo, "")
        self.store_name_edit.clear()

    def on_store_account_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.store_tree.selectedIndexes()}) if hasattr(self, "store_tree") else []
        if not rows:
            return
        store_key = self.store_tree.item(rows[0], 0).data(Qt.UserRole)
        for row in self.db.store_accounts():
            if row.get("StoreKey") == store_key:
                self._combo_set(self._store_district_combo, row.get("District", ""))
                self.store_name_edit.setText(row.get("Store", ""))
                return

    def save_sales_rep_from_form(self) -> None:
        try:
            name = self.rep_name_edit.text().strip()
            phone = self.rep_phone_edit.text().strip()
            created_by = self.rep_created_by_edit.text().strip() if hasattr(self, "rep_created_by_edit") else ""
            self.db.save_sales_rep(name, phone)
            # Also save the created_by mapping if provided
            if created_by:
                self.db.upsert_created_by_mapping(created_by, name, phone)
            self.refresh_sales_reps_table()
            self.set_status(f"Saved: {name} | {phone}" + (f" | Created By: {created_by}" if created_by else ""))
        except Exception as exc:
            QMessageBox.critical(self, "Save failed", str(exc))

    def delete_selected_sales_rep(self) -> None:
        rows = sorted({i.row() for i in self.rep_tree.selectedIndexes()}) if hasattr(self, "rep_tree") else []
        if not rows:
            QMessageBox.information(self, "No selection", "Select an employee first.")
            return
        if QMessageBox.question(self, "Delete employee", "Delete the selected employee record?") != QMessageBox.Yes:
            return
        rep_key = self.rep_tree.item(rows[0], 0).data(Qt.UserRole)
        created_by_item = self.rep_tree.item(rows[0], 2)
        vals_created_by = created_by_item.text() if created_by_item is not None else ""
        if rep_key:
            self.db.delete_sales_rep(rep_key)
        # Also remove the created_by mapping if present
        if vals_created_by:
            self.db.delete_created_by_mapping(vals_created_by)
        self.refresh_sales_reps_table()
        self.clear_rep_form()
        self.set_status("Deleted selected employee.")

    def clear_rep_form(self) -> None:
        self.rep_name_edit.clear()
        self.rep_phone_edit.clear()
        if hasattr(self, "rep_created_by_edit"):
            self.rep_created_by_edit.clear()

    def on_sales_rep_select(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.rep_tree.selectedIndexes()}) if hasattr(self, "rep_tree") else []
        if not rows:
            return
        row = rows[0]
        name_item = self.rep_tree.item(row, 0)
        phone_item = self.rep_tree.item(row, 1)
        created_by_item = self.rep_tree.item(row, 2)
        if name_item is not None:
            self.rep_name_edit.setText(name_item.text() if name_item is not None else "")
            self.rep_phone_edit.setText(phone_item.text() if phone_item is not None else "")
            if hasattr(self, "rep_created_by_edit"):
                self.rep_created_by_edit.setText(created_by_item.text() if created_by_item is not None else "")
    def download_store_accounts_template(self):
        """Store list import template — District + Store columns."""
        self._save_import_template_xlsx(
            sheet_title="Stores",
            initial_file="Store_List_Template.xlsx",
            headers=["District", "Store"],
            example_rows=[
                ["Arizona", "Store 1"],
                ["Houston", "Store 2"],
            ],
            column_hint="District | Store",
            done_note="Replace the two example rows with the real district store list, then use Import XLSX.",
        )

    def download_employees_template(self):
        """Employees import template — Employee Name + Phone Number."""
        self._save_import_template_xlsx(
            sheet_title="Employees",
            initial_file="Employees_Template.xlsx",
            headers=["Employee Name", "Phone Number"],
            example_rows=[
                ["John Doe", "346-555-0101"],
                ["Jane Smith", "346-555-0102"],
            ],
            column_hint="Employee Name | Phone Number",
            done_note="Replace the two example rows with the real employees, then use Import XLSX.",
        )

    def download_excluded_imeis_template(self):
        """Excluded-IMEIs import template — District | Product | IMEI | Comments."""
        self._save_import_template_xlsx(
            sheet_title="Excluded IMEIs",
            initial_file="Excluded_IMEIs_Template.xlsx",
            headers=["District", "Product", "IMEI", "Comments"],
            example_rows=[
                ["Arizona", "Device Model A", "350000000000123", "Demo unit"],
                ["Houston", "Device Model B", "350000000000124", "Damaged"],
            ],
            column_hint="District | Product | IMEI | Comments",
            done_note="IMEI is required; the other columns are optional. Then use Import XLSX.",
        )

    def pick_inventory(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self, "Select Inventory_Count_Result_Details", "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)")
        if path:
            self.inventory_edit.setText(path)

    def pick_time_sheet(self) -> None:
        path, _selected = QFileDialog.getOpenFileName(
            self, "Select Timesheet (timesheets_*.xlsx)", "",
            "Excel Files (*.xlsx *.xlsm *.xls);;All Files (*.*)")
        if path:
            self.time_sheet_edit.setText(path)

    def load_variances(self) -> None:
        inventory = self.inventory_edit.text().strip()
        time_sheet = self.time_sheet_edit.text().strip()
        if not inventory or not os.path.exists(inventory):
            QMessageBox.critical(self, "Missing file", "Select Inventory_Count_Result_Details.xlsx first.")
            return
        if not time_sheet or not os.path.exists(time_sheet):
            QMessageBox.critical(self, "Missing file", "Select Timesheet file (timesheets_*.xlsx) first.")
            return
        try:
            self.clear_current_ui(silent=True)
            self.set_status("Reading Excel files...")
            inv_records = read_xlsx_records(inventory)
            ts_records = read_xlsx_records(time_sheet)
            self.current_inventory_records = inv_records
            self.current_time_sheet_records = ts_records
            self.master_store_records = self.db.store_master_records()
            self.set_status("Building Inventory Status and Variance Audit rows...")

            self.status_rows, status_summary = build_inventory_status_rows(inv_records, ts_records, master_store_records=self.master_store_records, source_file=os.path.basename(inventory))
            self.status_row_by_key = {r.key: r for r in self.status_rows}
            self.db.upsert_inventory_status_rows(self.status_rows)
            rows, variance_summary = extract_variances(inv_records, ts_records, master_store_records=self.master_store_records, source_file=os.path.basename(inventory))
            self.db.upsert_rows(rows)
            self.loaded_keys = {row.key for row in rows}
            self.data_loaded = True

            self.populate_status_filters()
            self.populate_audit_filters()
            self.refresh_status_table()
            self.refresh_table()

            self.summary_text = (
                f"Inventory Status. Completed: {status_summary['completed']}   Pending: {status_summary['pending']}   "
                f"Variance Audit: {len(rows)}   SIMs skipped: {variance_summary['skipped_sims']}"
            )
            self.set_status(
                f"Loaded current session. Inventory Status rows: {len(self.status_rows)}. Variances: {len(rows)}. "
                f"Inventory rows read: {variance_summary['raw_inventory_rows']}. Latest rows used: {variance_summary['latest_inventory_rows']}. "
                f"Stale rows skipped: {variance_summary['stale_inventory_rows']}. SIM rows skipped: {variance_summary['skipped_sims']}."
            )
        except Exception as exc:
            traceback.print_exc()
            QMessageBox.critical(self, "Load failed", str(exc))
            self.set_status("Load failed. Check the Excel columns and try again.")

    def add_store_prompt(self) -> None:
        district, ok = QInputDialog.getText(self, "Add Store", "Enter district name.")
        if not ok:
            return
        district = normalize_district(district)
        if district == "Unknown":
            QMessageBox.critical(self, "Invalid district", "Please enter a valid district name.")
            return

        store, ok2 = QInputDialog.getText(self, "Add Store", "Enter store name.")
        if not ok2:
            return
        store = display_store(store)
        if not store:
            QMessageBox.critical(self, "Invalid store", "Store name is required.")
            return

        self.db.save_store_account(district, store, "", "", "")
        self.master_store_records = self.db.store_master_records()
        self.refresh_store_accounts_table()
        self.refresh_entry_form_districts()

        if self.data_loaded:
            source_name = os.path.basename(self.inventory_edit.text().strip()) if self.inventory_edit.text().strip() else ""
            self.status_rows, status_summary = build_inventory_status_rows(
                self.current_inventory_records,
                self.current_time_sheet_records,
                master_store_records=self.master_store_records,
                source_file=source_name,
            )
            self.status_row_by_key = {r.key: r for r in self.status_rows}
            self.populate_status_filters()
            self.refresh_status_table()
            self.summary_text = (
                f"Inventory Audit Status. Completed: {status_summary['completed']}   Pending: {status_summary['pending']}   Variance Audit: {len(self.loaded_keys)}"
            )
        else:
            self.populate_status_filters()

        QMessageBox.information(self, "Store saved", f"Saved store.\nDistrict: {district}\nStore: {store}")
        self.set_status(f"Added store: {district} | {store}")

    def populate_status_filters(self) -> None:
        districts = set()
        for r in self.status_rows:
            d = normalize_district(r.district or "Unknown")
            if d and d != "Unknown":
                districts.add(d)
        for d in self.db.all_known_districts():
            if d and d != "Unknown":
                districts.add(d)
        district_values = ["All Districts"] + sorted(districts)
        current = self.status_district_combo.currentText()
        self.status_district_combo.blockSignals(True)
        self.status_district_combo.clear()
        self.status_district_combo.addItems(district_values)
        self._combo_set(self.status_district_combo, current if current in district_values else "All Districts")
        self.status_district_combo.blockSignals(False)
        self.update_status_store_filter_values()

    def update_status_store_filter_values(self) -> None:
        selected_district = self.status_district_combo.currentText().strip()
        rows = self.status_rows
        if selected_district and selected_district != "All Districts":
            rows = [r for r in rows if normalize_district(r.district) == normalize_district(selected_district)]
        stores = sorted({r.store for r in rows if r.store})
        values = ["All Stores"] + stores
        current = self.status_store_combo.currentText()
        self.status_store_combo.blockSignals(True)
        self.status_store_combo.clear()
        self.status_store_combo.addItems(values)
        self._combo_set(self.status_store_combo, current if current in values else "All Stores")
        self.status_store_combo.blockSignals(False)

    def populate_audit_filters(self) -> None:
        districts = set()
        rows = self.db.get_rows_by_keys(self.loaded_keys) if self.loaded_keys else []
        for r in rows:
            d = normalize_district(r.district or "Unknown")
            if d and d != "Unknown":
                districts.add(d)
        for d in self.db.all_known_districts():
            if d and d != "Unknown":
                districts.add(d)
        district_values = ["All Districts"] + sorted(districts)
        current = self.audit_district_combo.currentText()
        self.audit_district_combo.blockSignals(True)
        self.audit_district_combo.clear()
        self.audit_district_combo.addItems(district_values)
        self._combo_set(self.audit_district_combo, current if current in district_values else "All Districts")
        self.audit_district_combo.blockSignals(False)
        # Also populate the final district combo
        if hasattr(self, "final_district_combo"):
            final_current = self.final_district_combo.currentText()
            self.final_district_combo.blockSignals(True)
            self.final_district_combo.clear()
            self.final_district_combo.addItems(district_values)
            self._combo_set(self.final_district_combo, final_current if final_current in district_values else "All Districts")
            self.final_district_combo.blockSignals(False)
        self.update_audit_store_filter_values()

    def update_audit_store_filter_values(self) -> None:
        rows = self.db.get_rows_by_keys(self.loaded_keys) if self.loaded_keys else []
        district = self.audit_district_combo.currentText().strip()
        if district and district != "All Districts":
            rows = [r for r in rows if normalize_district(r.district) == normalize_district(district)]
        stores = sorted({r.store for r in rows if r.store})
        values = ["All Stores"] + stores
        current = self.audit_store_combo.currentText()
        self.audit_store_combo.blockSignals(True)
        self.audit_store_combo.clear()
        self.audit_store_combo.addItems(values)
        self._combo_set(self.audit_store_combo, current if current in values else "All Stores")
        self.audit_store_combo.blockSignals(False)

    def on_status_district_change(self, _event=None) -> None:
        self.update_status_store_filter_values()
        self.auto_check_status_rows()
        self.refresh_status_table()

    def on_status_store_change(self, _event=None) -> None:
        self.auto_check_status_rows()
        self.refresh_status_table()

    def on_audit_district_change(self, _event=None) -> None:
        self.update_audit_store_filter_values()
        self.auto_check_audit_rows()
        self.refresh_table()

    def on_audit_store_change(self, _event=None) -> None:
        self.auto_check_audit_rows()
        self.refresh_table()

    def matching_status_rows(self) -> List[InventoryStatusRow]:
        rows = list(self.status_rows)
        district = self.status_district_combo.currentText().strip()
        store = self.status_store_combo.currentText().strip()
        if district and district != "All Districts":
            rows = [r for r in rows if normalize_district(r.district) == normalize_district(district)]
        if store and store != "All Stores":
            rows = [r for r in rows if r.store == store]

        any_q = safe_text(self.status_search_any_edit.text()).lower()
        district_q = safe_text(self.status_search_district_edit.text()).lower()
        store_q = safe_text(self.status_search_store_edit.text()).lower()
        status_q = safe_text(self.status_search_status_edit.text()).lower()
        rep_q = safe_text(self.status_search_rep_edit.text()).lower()

        if district_q:
            rows = [r for r in rows if district_q in safe_text(r.district).lower()]
        if store_q:
            rows = [r for r in rows if store_q in safe_text(r.store).lower()]
        if status_q:
            rows = [r for r in rows if status_q in safe_text(r.status).lower()]
        if rep_q:
            rows = [r for r in rows if rep_q in safe_text(r.rep_name).lower()]
        if any_q:
            rows = [
                r for r in rows
                if any_q in " ".join([
                    safe_text(r.district),
                    safe_text(r.store),
                    safe_text(r.status),
                    safe_text(r.rep_name),
                ]).lower()
            ]
        return rows

    def clear_status_search(self) -> None:
        self.status_search_any_edit.clear()
        self.status_search_district_edit.clear()
        self.status_search_store_edit.clear()
        self.status_search_status_edit.clear()
        self.status_search_rep_edit.clear()
        self.refresh_status_table()

    def refresh_status_table(self) -> None:
        self.status_tree.setRowCount(0)
        self.status_key_by_iid.clear()
        if not self.data_loaded:
            return
        rows = self.matching_status_rows()
        rows.sort(key=lambda r: (r.district, r.store, r.rep_name, r.status))
        state_map = self.db.inventory_status_state_map()
        for row in rows:
            checkbox = "☑ Send" if row.key in self.status_checked_keys else "☐ Skip"
            tag = self.status_row_tag(row, state_map)
            iid = self.status_tree.rowCount()
            self.status_tree.insertRow(iid)
            values = (row.district, row.store, row.status, row.rep_name, checkbox)
            for col, value in enumerate(values):
                self.status_tree.setItem(iid, col, QTableWidgetItem(str(value)))
            self._paint_row(self.status_tree, iid, tag)
            self.status_tree.item(iid, 0).setData(Qt.UserRole, row.key)
            self.status_key_by_iid[str(iid)] = row.key
        checked = len([r for r in rows if r.key in self.status_checked_keys])
        self.set_status(f"Inventory Status rows shown: {len(rows)}. Checked for send: {checked}.")

    def on_status_tree_click(self, item) -> None:
        if item.column() != 4:  # checkbox column
            return
        row = item.row()
        key = self.status_tree.item(row, 0).data(Qt.UserRole)
        if not key:
            return
        if key in self.status_checked_keys:
            self.status_checked_keys.remove(key)
        else:
            self.status_checked_keys.add(key)
        self.refresh_status_table()

    def matching_audit_rows(self) -> List[VarianceRow]:
        rows = self.db.get_rows_by_keys(self.loaded_keys) if self.loaded_keys else []
        if not self.include_cleared_check.isChecked():
            rows = [row for row in rows if not row.cleared]
        district = self.audit_district_combo.currentText().strip()
        store = self.audit_store_combo.currentText().strip()
        if district and district != "All Districts":
            rows = [row for row in rows if normalize_district(row.district) == normalize_district(district)]
        if store and store != "All Stores":
            rows = [row for row in rows if row.store == store]

        any_q = safe_text(self.audit_search_any_edit.text()).lower()
        district_q = safe_text(self.audit_search_district_edit.text()).lower()
        store_q = safe_text(self.audit_search_store_edit.text()).lower()
        product_q = safe_text(self.audit_search_product_edit.text()).lower()
        imei_q = safe_text(self.audit_search_imei_edit.text()).lower()
        rep_q = safe_text(self.audit_search_rep_edit.text()).lower()

        if district_q:
            rows = [row for row in rows if district_q in safe_text(row.district).lower()]
        if store_q:
            rows = [row for row in rows if store_q in safe_text(row.store).lower()]
        if product_q:
            rows = [row for row in rows if product_q in safe_text(row.product).lower()]
        if imei_q:
            rows = [row for row in rows if imei_q in safe_text(row.imei).lower()]
        if rep_q:
            rows = [row for row in rows if rep_q in safe_text(row.rep_name).lower()]
        if any_q:
            rows = [
                row for row in rows
                if any_q in " ".join([
                    safe_text(row.district),
                    safe_text(row.store),
                    safe_text(row.product),
                    safe_text(row.imei),
                    safe_text(row.status),
                    safe_text(row.rep_name),
                ]).lower()
            ]

        rows = self.filter_excluded_variance_rows(rows)
        return rows

    def clear_audit_search(self) -> None:
        self.audit_search_any_edit.clear()
        self.audit_search_district_edit.clear()
        self.audit_search_store_edit.clear()
        self.audit_search_product_edit.clear()
        self.audit_search_imei_edit.clear()
        self.audit_search_rep_edit.clear()
        self.refresh_table()

    def copy_text_to_clipboard(self, text_value: str, status_message: str = "Copied to clipboard.") -> None:
        text_value = safe_text(text_value)
        if not text_value:
            return
        try:
            QApplication.clipboard().setText(text_value)
            self.set_status(status_message)
        except Exception as exc:
            QMessageBox.critical(self, "Copy failed", str(exc))

    def copy_selected_imei(self, event=None) -> None:
        rows = sorted({i.row() for i in self.audit_tree.selectedIndexes()}) if hasattr(self, "audit_tree") else []
        if not rows:
            QMessageBox.information(self, "No selection", "Select a variance row first.")
            return None
        imeis: List[str] = []
        for row in rows:
            imei_item = self.audit_tree.item(row, 3)
            if imei_item is not None:
                imei = safe_text(imei_item.text())
                if imei:
                    imeis.append(imei)
        if not imeis:
            QMessageBox.information(self, "No IMEI", "No IMEI found in selected row.")
            return None
        self.copy_text_to_clipboard("\n".join(imeis), f"Copied {len(imeis)} IMEI(s) to clipboard.")
        return None

    def refresh_table(self) -> None:
        self.audit_tree.setRowCount(0)
        self.key_by_iid.clear()

        if not self.data_loaded:
            self.summary_text = "No data loaded"
            self.set_status("UI cleared. Select both file locations, then click Load Variances.")
            return

        rows = self.matching_audit_rows()
        rows.sort(key=lambda r: (r.district, r.store, r.rep_name, r.status, r.product, r.imei))

        for row in rows:
            clearance = "Cleared" if row.cleared else ("Sent" if row.sent_count else "Not cleared")
            checkbox = "☑ Send" if row.key in self.audit_checked_keys else "☐ Skip"
            iid = self.audit_tree.rowCount()
            self.audit_tree.insertRow(iid)
            values = (row.district, row.store, row.product, row.imei, row.status, row.rep_name, clearance, checkbox)
            for col, value in enumerate(values):
                self.audit_tree.setItem(iid, col, QTableWidgetItem(str(value)))
            self._paint_row(self.audit_tree, iid, self.audit_row_tag(row))
            self.audit_tree.item(iid, 0).setData(Qt.UserRole, row.key)
            self.key_by_iid[str(iid)] = row.key
        cleared_count = sum(1 for r in rows if r.cleared)
        pending_count = sum(1 for r in rows if not r.cleared)
        checked = sum(1 for r in rows if r.key in self.audit_checked_keys)
        self.set_status(f"Variance Audit rows shown: {len(rows)}. Cleared: {cleared_count}. Not cleared: {pending_count}. Checked for send: {checked}.")

    def on_audit_double_click(self, _event=None) -> None:
        rows = sorted({i.row() for i in self.audit_tree.selectedIndexes()})
        if not rows:
            return
        key = self.audit_tree.item(rows[0], 0).data(Qt.UserRole)
        if not key or key not in self.key_by_iid.values():
            return
        self.send_selected()

    def on_audit_tree_click(self, item) -> None:
        row = item.row()
        key = self.audit_tree.item(row, 0).data(Qt.UserRole)
        if item.column() == 3:  # IMEI column: click copies the IMEI
            imei_item = self.audit_tree.item(row, 3)
            if imei_item is not None:
                self.copy_text_to_clipboard(imei_item.text(), f"Copied IMEI: {imei_item.text()}")
            return
        if item.column() != 7:  # checkbox column
            return
        if not key:
            return
        if key in self.audit_checked_keys:
            self.audit_checked_keys.remove(key)
        else:
            self.audit_checked_keys.add(key)
        self.refresh_table()

    def selected_rows(self) -> List[VarianceRow]:
        rows = sorted({i.row() for i in self.audit_tree.selectedIndexes()})
        keys = [self.audit_tree.item(row, 0).data(Qt.UserRole) for row in rows]
        keys = [k for k in keys if k]
        return self.db.get_rows_by_keys(keys)

    def current_mode_key(self) -> str:
        label = self.audit_send_mode_combo.currentText().strip().lower()
        if label == "store":
            return "store"
        if label in {"sales rep", "rep", "salesperson"}:
            return "rep"
        return "district"

    def current_status_mode_key(self) -> str:
        label = self.status_send_mode_combo.currentText().strip().lower()
        if label == "store":
            return "store"
        if label in {"sales rep", "rep", "salesperson"}:
            return "rep"
        return "district"
    def filter_excluded_variance_rows(self, rows: List[VarianceRow]) -> List[VarianceRow]:
        if not rows:
            return []
        return [row for row in rows if not self.db.is_device_excluded(row.district, row.product, row.imei)]
    def status_row_tag(self, row: InventoryStatusRow, state_map: Optional[Dict[str, Dict[str, str]]] = None) -> str:
        state_map = state_map if state_map is not None else self.db.inventory_status_state_map()
        state = state_map.get(normalize_store(row.store), {})
        status_norm = normalize_header(row.status)
        previous_norm = normalize_header(state.get("Previous Status", ""))
        sent_at = safe_text(state.get("Last Sent At", ""))

        if status_norm != "completed":
            return "status_pending"
        if sent_at:
            return "status_completed_sent"
        if previous_norm == "pending":
            return "status_completed_after_update"
        return "status_completed"

    def auto_check_pending_status_rows(self) -> None:
        rows = [r for r in self.matching_status_rows() if normalize_header(r.status) != "completed"]
        self.status_checked_keys = {r.key for r in rows}
        self.refresh_status_table()
        self.set_status(f"Prepared {len(rows)} pending inventory status row(s) for sending.")

    def auto_check_status_rows(self) -> None:
        rows = self.matching_status_rows()
        self.status_checked_keys = {r.key for r in rows}
        self.refresh_status_table()
        self.set_status(f"Prepared {len(rows)} inventory status row(s) for sending. Uncheck any row you want to skip.")

    def clear_status_checkmarks(self) -> None:
        self.status_checked_keys.clear()
        self.refresh_status_table()
        self.set_status("Inventory status checkmarks cleared.")
    def auto_check_audit_rows(self) -> None:
        rows = self.matching_audit_rows()
        self.audit_checked_keys = {r.key for r in rows}
        self.refresh_table()
        self.set_status(f"Prepared {len(rows)} variance row(s) for sending. Uncheck any row you want to skip.")

    def clear_audit_checkmarks(self) -> None:
        self.audit_checked_keys.clear()
        self.refresh_table()
        self.set_status("Variance checkmarks cleared.")
    @staticmethod
    def audit_row_tag(row: VarianceRow) -> str:
        if row.cleared:
            return "variance_cleared"
        if row.sent_count:
            return "variance_sent"
        return "variance_pending"
    @staticmethod
    def grouped_batches(rows: List[VarianceRow], mode: str) -> List[Tuple[str, str, List[VarianceRow]]]:
        grouped: Dict[Tuple[str, str], List[VarianceRow]] = {}
        for row in rows:
            district = normalize_district(row.district or "Unknown")
            if mode == "store":
                key = (district, row.store or "Unknown Store")
            elif mode == "rep":
                key = (district, row.rep_name or "Unknown Rep")
            else:
                key = (district, district)
            grouped.setdefault(key, []).append(row)
        batches: List[Tuple[str, str, List[VarianceRow]]] = []
        for (district, value), group_rows in sorted(grouped.items(), key=lambda item: (item[0][0], item[0][1])):
            title = district if mode == "district" else f"{value} | {district}"
            batches.append((title, district, group_rows))
        return batches

    @staticmethod
    def grouped_status_batches(rows: List[InventoryStatusRow], mode: str) -> List[Tuple[str, str, List[InventoryStatusRow]]]:
        grouped: Dict[Tuple[str, str], List[InventoryStatusRow]] = {}
        for row in rows:
            district = normalize_district(row.district or "Unknown")
            if mode == "store":
                key = (district, row.store or "Unknown Store")
            elif mode == "rep":
                key = (district, row.rep_name or "Unknown Rep")
            else:
                key = (district, district)
            grouped.setdefault(key, []).append(row)
        batches: List[Tuple[str, str, List[InventoryStatusRow]]] = []
        for (district, value), group_rows in sorted(grouped.items(), key=lambda item: (item[0][0], item[0][1])):
            title = district if mode == "district" else f"{value} | {district}"
            batches.append((title, district, group_rows))
        return batches

    @staticmethod
    def chunks(rows: List, size: int) -> Iterable[List]:
        for i in range(0, len(rows), size):
            yield rows[i:i + size]

    def send_checked_variances(self) -> None:
        if not self.data_loaded:
            QMessageBox.information(self, "Load files first", "Select both file locations and click Load Variances first.")
            return
        rows = [row for row in self.db.get_rows_by_keys(self.loaded_keys) if row.key in self.audit_checked_keys]
        rows = [row for row in rows if self.include_cleared_check.isChecked() or not row.cleared]
        rows = self.filter_excluded_variance_rows(rows)
        if not rows:
            QMessageBox.information(self, "Nothing checked", "Select a district or store so variance rows get checked, then send them.")
            return
        mode = self.current_mode_key()
        batches = self.grouped_batches(rows, mode)
        if QMessageBox.question(self, "Send checked variances", f"Send {len(rows)} checked variance row(s) in {len(batches)} image batch(es)?") != QMessageBox.Yes:
            return
        self._send_rows_thread(rows, mode=mode, manual=True)

    def send_selected(self) -> None:
        rows = self.filter_excluded_variance_rows(self.selected_rows())
        if not rows:
            QMessageBox.information(self, "No selection", "Select one or more non-excluded variance rows first.")
            return
        mode = self.current_mode_key()
        self._send_rows_thread(rows, mode=mode, manual=True)

    def send_pending(self) -> None:
        if not self.data_loaded:
            QMessageBox.information(self, "Load files first", "Select both file locations and click Load Variances first.")
            return
        mode = self.current_mode_key()
        rows = [row for row in self.db.get_rows_by_keys(self.loaded_keys) if not row.cleared]
        district = self.audit_district_combo.currentText().strip()
        store = self.audit_store_combo.currentText().strip()
        if district and district != "All Districts":
            rows = [row for row in rows if normalize_district(row.district) == normalize_district(district)]
        if store and store != "All Stores":
            rows = [row for row in rows if row.store == store]
        rows = self.filter_excluded_variance_rows(rows)
        if self.send_only_unsent_check.isChecked():
            rows = [row for row in rows if not row.last_sent_at]
        if not rows:
            QMessageBox.information(self, "Nothing to send", "No pending unsent variances found in the current filter.")
            return
        label = SEND_MODE_LABELS.get(mode, mode.title())
        batches = self.grouped_batches(rows, mode)
        if QMessageBox.question(self, "Send pending", f"Send {len(rows)} pending variance rows in {len(batches)} image batch(es) by {label}?\n\nImages will route to district WhatsApp groups.") != QMessageBox.Yes:
            return
        self._send_rows_thread(rows, mode=mode, manual=False)

    def send_checked_status(self) -> None:
        if not self.data_loaded:
            QMessageBox.information(self, "Load files first", "Select both file locations and click Load Variances first.")
            return
        rows = [self.status_row_by_key[k] for k in self.status_checked_keys if k in self.status_row_by_key]
        if not rows:
            QMessageBox.information(self, "Nothing checked", "Select a district or store so Inventory Audit Status rows get checked, then send them.")
            return
        mode = self.current_status_mode_key()
        batches = self.grouped_status_batches(rows, mode)
        if QMessageBox.question(self, "Send Inventory Audit Status", f"Send {len(rows)} inventory status row(s) in {len(batches)} image batch(es)?") != QMessageBox.Yes:
            return
        thread = threading.Thread(target=self._send_status_rows, args=(rows, mode), daemon=True)
        thread.start()

    def _send_rows(self, rows: List[VarianceRow], mode: str, manual: bool) -> None:
        try:
            renderer = ImageRenderer()
            sender = WhatsAppSender(status_callback=self.set_status,
                                    mode=getattr(self, "wa_mode_var", None) and self.wa_mode_var or "desktop")
            batches = self.grouped_batches(rows, mode)
            send_mode_for_file = mode
            for batch_title, district, batch_rows in batches:
                group_name = group_name_for_district(district, self.db)
                for chunk_no, chunk_rows in enumerate(self.chunks(batch_rows, AUTO_SEND_CHUNK_SIZE), start=1):
                    title = batch_title if len(batch_rows) <= AUTO_SEND_CHUNK_SIZE else f"{batch_title} | Part {chunk_no}"
                    image_path = renderer.render_rows(title, chunk_rows, mode=send_mode_for_file)
                    try:
                        self.set_status(f"Sending {len(chunk_rows)} variance row(s) to {group_name}. {title}.")
                        _win_state = self._save_window_state()
                        sender.send_image(group_name, image_path, text_message=self.variance_request_message(chunk_rows))
                        self._restore_window_state(_win_state)
                        self.db.mark_sent(chunk_rows, group_name, title, str(image_path), mode=send_mode_for_file)
                    except Exception as exc:
                        self.db.mark_sent(chunk_rows, group_name, title, str(image_path), mode=send_mode_for_file, error=str(exc))
                        self._worker_dialog("error", "WhatsApp send failed", f"Image was created but WhatsApp send failed.\n\nGroup: {group_name}\nBatch: {title}\nImage: {image_path}\n\nError: {exc}")
                        self.set_status(f"Send failed. Image saved: {image_path}")
                        return
            self._bridge.refresh_audit.emit()
            self.set_status("Variance image sending completed.")
        except Exception as exc:
            traceback.print_exc()
            self._worker_dialog("error", "Send failed", str(exc))
            self.set_status("Send failed.")

    def _send_status_rows(self, rows: List[InventoryStatusRow], mode: str) -> None:
        try:
            sender = WhatsAppSender(status_callback=self.set_status,
                                    mode=getattr(self, "wa_mode_var", None) and self.wa_mode_var or "desktop")
            batches = self.grouped_status_batches(rows, mode)
            for batch_title, district, batch_rows in batches:
                group_name = group_name_for_district(district, self.db)
                for chunk_no, chunk_rows in enumerate(self.chunks(batch_rows, 28), start=1):
                    title = batch_title if len(batch_rows) <= 28 else f"{batch_title} | Part {chunk_no}"
                    image_path = self._render_status_rows(title, chunk_rows, mode="inventory status")
                    self.set_status(f"Sending Inventory Audit Status with {len(chunk_rows)} row(s) to {group_name}. {title}.")
                    _win_state = self._save_window_state()
                    sender.send_image(group_name, image_path, text_message=self.inventory_status_message())
                    self._restore_window_state(_win_state)
                    self.db.mark_status_sent(chunk_rows)
                    pending_message = self.pending_inventory_count_message(chunk_rows)
                    if pending_message:
                        self.set_status(f"Sending incomplete inventory count reminder to {group_name}.")
                        _win_state2 = self._save_window_state()
                        sender.send_text(group_name, pending_message)
                        self._restore_window_state(_win_state2)
            self._bridge.refresh_status.emit()
            self.set_status("Inventory Audit Status image sending completed.")
        except Exception as exc:
            traceback.print_exc()
            self._worker_dialog("error", "Send failed", str(exc))
            self.set_status("Inventory Audit Status send failed.")

    def send_starting_message(self) -> None:
        """Send 'Please complete an Inventory count in 15 minutes.' to district WhatsApp groups."""
        if not self.data_loaded:
            QMessageBox.information(self, "Load files first", "Select both file locations and click Load Variances first.")
            return
        all_districts = sorted({normalize_district(row.district) for row in self.status_rows if safe_text(row.district)})
        if not all_districts:
            QMessageBox.information(self, "No districts", "No district data is available.")
            return
        selected = self.final_district_combo.currentText().strip()
        if selected and selected != "All Districts":
            districts = [normalize_district(selected)]
        else:
            districts = all_districts
        if QMessageBox.question(self, "Send starting message", f"Send starting message to {len(districts)} district group(s)?\n\nMessage: Please complete an Inventory count in 15 minutes.") != QMessageBox.Yes:
            return
        thread = threading.Thread(target=self._send_starting_message_thread, args=(districts,), daemon=True)
        thread.start()

    def _send_starting_message_thread(self, districts: List[str]) -> None:
        try:
            sender = WhatsAppSender(status_callback=self.set_status,
                                    mode=getattr(self, "wa_mode_var", None) and self.wa_mode_var or "desktop")
            message = "Please complete an Inventory count in 15 minutes."
            for district in districts:
                group_name = group_name_for_district(district, self.db)
                self.set_status(f"Sending starting message to {group_name}...")
                _win_state = self._save_window_state()
                sender.send_text(group_name, message)
                self._restore_window_state(_win_state)
            self.set_status("Starting message sent to all selected districts.")
        except Exception as exc:
            traceback.print_exc()
            self._worker_dialog("error", "Send failed", str(exc))
            self.set_status(f"Starting message send failed: {exc}")

    def send_reminder(self, reminder_number: int) -> None:
        """Send a single reminder message for uncleared variances to district WhatsApp groups."""
        reminder_messages = {
            1: "Please clear the pending variances.",
            2: "This is the second reminder. Please clear the pending variances immediately.",
            3: "Final reminder. Uncleared variances will be reported.",
        }
        if reminder_number not in reminder_messages:
            return
        if not self.data_loaded:
            QMessageBox.information(self, "Load files first", "Select both file locations and click Load Variances first.")
            return
        all_rows = self.db.get_rows_by_keys(self.loaded_keys)
        uncleared_rows = [row for row in all_rows if not row.cleared]
        uncleared_rows = self.filter_excluded_variance_rows(uncleared_rows)
        if not uncleared_rows:
            QMessageBox.information(self, "No uncleared variances", "All variances have been cleared. No reminder needed.")
            return
        all_districts = sorted({normalize_district(row.district) for row in uncleared_rows if safe_text(row.district)})
        if not all_districts:
            QMessageBox.information(self, "No districts", "No district data available for reminders.")
            return
        selected = self.final_district_combo.currentText().strip()
        if selected and selected != "All Districts":
            districts = [normalize_district(selected)]
        else:
            districts = all_districts
        message = reminder_messages[reminder_number]
        if QMessageBox.question(
                self,
                f"Send Reminder {reminder_number}",
                f"Send reminder {reminder_number} to {len(districts)} district group(s)?\n\n"
                f"Message: {message}\n\n"
                f"Total {len(uncleared_rows)} uncleared variance row(s).",
        ) != QMessageBox.Yes:
            return
        thread = threading.Thread(target=self._send_single_reminder_thread, args=(districts, uncleared_rows, reminder_number, message), daemon=True)
        thread.start()

    def _send_single_reminder_thread(self, districts: List[str], uncleared_rows: List[VarianceRow], reminder_number: int, message: str) -> None:
        try:
            sender = WhatsAppSender(status_callback=self.set_status,
                                    mode=getattr(self, "wa_mode_var", None) and self.wa_mode_var or "desktop")
            for district in districts:
                district_rows = [row for row in uncleared_rows if normalize_district(row.district) == normalize_district(district)]
                if not district_rows:
                    continue
                group_name = group_name_for_district(district, self.db)
                rep_mentions = self.variance_request_message(district_rows)
                full_message = f"Reminder {reminder_number}/3: {message}"
                if rep_mentions:
                    full_message = f"{rep_mentions} {message}"
                self.set_status(f"Sending reminder {reminder_number}/3 to {group_name}...")
                _win_state = self._save_window_state()
                sender.send_text(group_name, full_message)
                self._restore_window_state(_win_state)
                time.sleep(2)
            self.set_status(f"Reminder {reminder_number} sent successfully.")
            try:
                self._worker_dialog("info", "Reminder sent", f"Reminder {reminder_number} has been sent to the selected district group(s).")
            except Exception:
                pass
        except Exception as exc:
            traceback.print_exc()
            self._worker_dialog("error", "Send failed", str(exc))
            self.set_status(f"Reminder {reminder_number} send failed: {exc}")

    def send_final_district_result(self) -> None:
        if not self.data_loaded:
            QMessageBox.information(self, "Load files first", "Select both file locations and click Load Variances first.")
            return
        all_districts = sorted({normalize_district(row.district) for row in self.status_rows if safe_text(row.district)})
        if not all_districts:
            all_districts = sorted({normalize_district(row.district) for row in self.db.get_rows_by_keys(self.loaded_keys) if safe_text(row.district)})
        if not all_districts:
            QMessageBox.information(self, "No districts", "No district data is available for final sending.")
            return
        selected = self.final_district_combo.currentText().strip()
        if selected and selected != "All Districts":
            districts = [normalize_district(selected)]
        else:
            districts = all_districts
        if QMessageBox.question(self, "Send final district result", f"Send final audit result to {len(districts)} district group(s)?") != QMessageBox.Yes:
            return
        thread = threading.Thread(target=self._send_final_district_result_thread, args=(districts,), daemon=True)
        thread.start()

    def _send_final_district_result_thread(self, districts: List[str]) -> None:
        final_results: List[Dict[str, str]] = []
        try:
            sender = WhatsAppSender(self.set_status, mode=getattr(self, "wa_mode_var", None) and self.wa_mode_var or "desktop")
            renderer = ImageRenderer()
            all_rows = self.db.get_rows_by_keys(self.loaded_keys)
            self.set_status("Sending final district audit results...")
            for district in districts:
                district_rows = [row for row in all_rows if normalize_district(row.district) == normalize_district(district)]
                district_rows = [row for row in district_rows if not row.cleared]
                district_rows = self.filter_excluded_variance_rows(district_rows)
                group_name = group_name_for_district(district, self.db)
                pending_status_rows = [
                    row for row in self.status_rows
                    if normalize_district(row.district) == normalize_district(district)
                    and normalize_header(row.status) != "completed"
                ]
                caption = self.final_district_caption(district, district_rows, pending_status_rows)
                if district_rows:
                    image_path = renderer.render_rows(district, district_rows, mode="district")
                    status = "Variance Found"
                else:
                    image_path = self._render_no_variance_image(district)
                    status = "No Variance Found"
                self.set_status(f"Sending final district result to {group_name}.")
                _win_state = self._save_window_state()
                sender.send_image(group_name, image_path, text_message=caption)
                self._restore_window_state(_win_state)
                final_results.append({
                    "district": district,
                    "stores": self.format_store_list_caption(sorted({safe_text(r.store) for r in district_rows if safe_text(r.store)})),
                    "pending_inventory_stores": self.format_store_list_caption(sorted({safe_text(r.store) for r in pending_status_rows if safe_text(r.store)})),
                    "status": status,
                    "caption": caption,
                    "image_path": str(image_path),
                    "group_name": group_name,
                    "sent_at": now_text(),
                })

            log_path = self.create_full_inventory_audit_log(final_results)
            self.set_status(f"Final district audit result sending completed. Full audit log created: {log_path}")
            try:
                self._worker_dialog("info", "Audit log created", f"Full inventory audit log created:\n{log_path}")
            except Exception:
                pass
        except Exception as exc:
            traceback.print_exc()
            self._worker_dialog("error", "Send failed", str(exc))
            self.set_status(f"Final district sending failed: {exc}")

    def export_log(self) -> None:
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_name = f"GFH_Variance_Clearance_Log_{stamp}.xlsx"
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        path, _selected = QFileDialog.getSaveFileName(
            self, "Export Clearance Log", str(EXPORT_DIR / default_name),
            "Excel Files (*.xlsx)")
        if not path:
            return
        keys = self.loaded_keys if self.data_loaded else None
        self.db.export_xlsx(Path(path), keys=keys)
        self.set_status(f"Exported clearance log: {path}")
        QMessageBox.information(self, "Export complete", f"Log exported:\n{path}")

    # ── Cross-instance DB sync ───────────────────────────────────────────
    def _start_db_sync_poll(self) -> None:
        """Begin the 2-second mtime-polling loop (Tk after() -> single-shot
        QTimer re-armed by _check_db_sync)."""
        self._schedule_db_sync()

    def _schedule_db_sync(self) -> None:
        if getattr(self, "_db_sync_timer", None) is None:
            self._db_sync_timer = QTimer(self)
            self._db_sync_timer.setSingleShot(True)
            self._db_sync_timer.timeout.connect(self._check_db_sync)
        self._db_sync_timer.start(2000)

    def closeEvent(self, event) -> None:
        """Cancel the background poll then close (Tk _on_close port)."""
        try:
            timer = getattr(self, "_db_sync_timer", None)
            if timer is not None:
                timer.stop()
            if self._db_sync_after_id is not None:
                self._db_sync_after_id = None
        except Exception:
            pass
        event.accept()
    def inventory_status_message(self) -> str:
        today = dt.datetime.now()
        return f"Inventory Audit Status as of {today.month}/{today.day}/{today.year}."

    def variance_request_message(self, rows: List[VarianceRow]) -> str:
        """Build WhatsApp @mention message for variance request.

        Uses resolve_phone_for_rep() which tries:
        1. created_by_mappings by 'Created By' username (ArmanAli → phone)
        2. created_by_mappings by employee_name (Abdullah Hussain → phone)
        3. sales_reps table by rep_name
        """
        seen_phones: set[str] = set()
        phones: List[str] = []
        missing: List[str] = []

        for row in rows:
            rep_name = safe_text(row.rep_name)
            created_by = safe_text(row.created_by)
            phone = normalize_phone(
                self.db.resolve_phone_for_rep(rep_name, created_by)
            )
            if not phone:
                missing.append(created_by or rep_name or "Unknown")
                continue
            if phone in seen_phones:
                continue
            seen_phones.add(phone)
            phones.append(whatsapp_mention(phone))

        if missing:
            unique_missing = sorted(set(missing))
            self.set_status(
                "No phone found for: " + ", ".join(unique_missing[:5]) +
                (f"... (+{len(unique_missing)-5} more)" if len(unique_missing) > 5 else "") +
                " — add phone in the Employees tab (Created By column)."
            )

        if not phones:
            return ""

        return " ".join(phones) + " please share the images of the variances."

    def pending_inventory_count_message(self, rows: List[InventoryStatusRow]) -> str:
        """Build the per-store reminder sent after the Inventory Audit Status
        image.

        Each pending store is addressed by the PERSON responsible whenever
        the row has one (the Salesperson column — which also carries the
        timesheet-matched employee for stores absent from the count file):

        1. Rep has a phone on file (Employees tab / sales reps): tag it —
               @<phone>, please complete the inventory count ASAP.
        2. Rep known but no phone anywhere: address the person by name —
               <Rep Name>, please complete the inventory count ASAP.
        3. No rep on the row: fall back to the store name —
               <Store Name>, please complete the inventory count ASAP.

        One line per pending store, blocks separated by a blank line.
        """
        messages: List[str] = []
        seen_stores: set[str] = set()

        for row in rows:
            if safe_text(row.status).lower() == "completed":
                continue

            store_name = safe_text(row.store)
            if not store_name or store_name in seen_stores:
                continue
            seen_stores.add(store_name)

            rep_name = safe_text(row.rep_name)
            if rep_name:
                phone = normalize_phone(self.db.resolve_phone_for_rep(rep_name))
                if phone:
                    messages.append(
                        f"{whatsapp_mention(phone)}, please complete the inventory count ASAP."
                    )
                    continue
                messages.append(
                    f"{rep_name}, please complete the inventory count ASAP."
                )
                continue

            messages.append(f"{store_name}, please complete the inventory count ASAP.")

        return "\n\n".join(messages)

    def create_full_inventory_audit_log(self, final_results: List[Dict[str, str]]) -> Path:
        EXPORT_DIR.mkdir(parents=True, exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = EXPORT_DIR / f"GFH_Full_Inventory_Audit_Log_{stamp}.xlsx"
        variance_rows = self.db.get_rows_by_keys(self.loaded_keys) if self.loaded_keys else []
        # Same exclusion filter as the Variance Audit tab — the log, the tab
        # and the WhatsApp sends must all show the SAME devices: every device
        # except IMEI-excluded ones.
        variance_rows = self.filter_excluded_variance_rows(variance_rows)
        final_results = final_results or []

        wb = openpyxl.Workbook()
        ws = wb.active
        headers = [
            "Section", "District", "Store", "Product", "IMEI", "Status", "Rep Name",
            "Clearance Status", "Sent Count", "Last Sent At", "Final Caption",
            "Pending Inventory Stores", "Final Image Path", "WhatsApp Group", "Generated At"
        ]
        ws.append(headers)
        
        imei_col_idx = headers.index("IMEI") + 1
        
        current_row = 2

        for row in sorted(self.status_rows, key=lambda r: (r.district, r.store, r.rep_name)):
            data = [
                "Inventory Status", row.district, row.store, "", "", row.status, row.rep_name,
                "", "", "", "", "", "", "", now_text()
            ]
            for c_idx, val in enumerate(data, start=1):
                cell = ws.cell(row=current_row, column=c_idx, value=val)
                if c_idx == imei_col_idx:
                    cell.data_type = 's'
                    cell.number_format = '@'
            current_row += 1

        for row in sorted(variance_rows, key=lambda r: (r.district, r.store, r.rep_name, r.product, r.imei)):
            data = [
                "Variance Audit", row.district, row.store, row.product, row.imei, row.status, row.rep_name,
                "Cleared" if row.cleared else "Not Cleared", row.sent_count, row.last_sent_at,
                "", "", "", "", now_text()
            ]
            for c_idx, val in enumerate(data, start=1):
                cell = ws.cell(row=current_row, column=c_idx, value=val)
                if c_idx == imei_col_idx:
                    cell.data_type = 's'
                    cell.number_format = '@'
            current_row += 1

        for item in final_results:
            data = [
                "Final District Send", item.get("district", ""), item.get("stores", ""), "", "",
                item.get("status", ""), "", "", "", "", item.get("caption", ""),
                item.get("pending_inventory_stores", ""), item.get("image_path", ""),
                item.get("group_name", ""), item.get("sent_at", "")
            ]
            for c_idx, val in enumerate(data, start=1):
                cell = ws.cell(row=current_row, column=c_idx, value=val)
                if c_idx == imei_col_idx:
                    cell.data_type = 's'
                    cell.number_format = '@'
            current_row += 1

        wb.save(output_path)
        return output_path

    def _send_rows_thread(self, rows: List[VarianceRow], mode: str, manual: bool) -> None:
        thread = threading.Thread(target=self._send_rows, args=(rows, mode, manual), daemon=True)
        thread.start()
    def _render_status_rows(self, batch_title: str, rows: List[InventoryStatusRow], mode: str = "status") -> Path:
        if Image is None:
            raise RuntimeError("Pillow is required. Install with: py -m pip install pillow")

        renderer = ImageRenderer()
        width = 1320
        margin = 28
        title_font = renderer._font(28, True)
        header_font = renderer._font(16, True)
        cell_font = renderer._font(15, False)
        cell_bold = renderer._font(15, True)
        sub_font = renderer._font(15, False)
        row_h = 34
        black = (18, 20, 43)
        border = (0, 0, 0)
        light_green = (210, 240, 214)
        pending_fill = (238, 230, 177)
        white = (255, 255, 255)
        district_colors = {
            "Arizona": (196, 138, 230),
            "Atlanta": (189, 49, 120),
            "Colorado": (96, 194, 236),
            "Colorado East": (96, 194, 236),
            "Colorado West": (78, 186, 228),
            "Houston": (244, 128, 128),
            "Louisiana": (247, 199, 34),
            "Tennessee": (74, 145, 62),
        }

        logo_img = None
        logo_w = 0
        logo_h_used = 0
        if STATUS_LOGO_PATH.exists():
            try:
                logo_img = Image.open(STATUS_LOGO_PATH).convert("RGBA")
                scale = min(560 / logo_img.width, 135 / logo_img.height)
                size = (max(1, int(logo_img.width * scale)), max(1, int(logo_img.height * scale)))
                logo_img = logo_img.resize(size)
                logo_w, logo_h_used = size
            except Exception:
                logo_img = None
                logo_w = 0
                logo_h_used = 0

        header_area_h = max(logo_h_used, 80) + 18
        table_top = margin + header_area_h + 18
        height = table_top + 26 + row_h * len(rows) + 24

        img = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(img)

        title_text = "GFH Inventory Audit"
        today = dt.datetime.now()
        date_text = f"{today.month}/{today.day}/{today.year}"

        # Measure each string so we can center it horizontally
        title_bbox = draw.textbbox((0, 0), title_text, font=title_font)
        title_w = title_bbox[2] - title_bbox[0]
        date_bbox = draw.textbbox((0, 0), date_text, font=cell_bold)
        date_w = date_bbox[2] - date_bbox[0]

        title_x = (width - title_w) // 2
        title_y = margin + 6
        draw.text((title_x, title_y), title_text, fill=black, font=title_font)

        date_x = (width - date_w) // 2
        draw.text((date_x, title_y + 40), date_text, fill=black, font=cell_bold)

        if safe_text(batch_title):
            bt_text = safe_text(batch_title)
            bt_bbox = draw.textbbox((0, 0), bt_text, font=sub_font)
            bt_w = bt_bbox[2] - bt_bbox[0]
            bt_x = (width - bt_w) // 2
            draw.text((bt_x, title_y + 40 + 26), bt_text, fill=black, font=sub_font)

        if logo_img is not None:
            logo_x = width - margin - logo_w
            logo_y = margin
            img.paste(logo_img, (logo_x, logo_y), logo_img)

        y = table_top
        x = margin
        col_widths = [160, 280, 145, 679]
        headers = ["District", "Store", "Status", "Salesperson"]
        draw.rectangle((x, y, width - margin, y + 26), outline=border, fill=white)
        cx = x
        for idx, h in enumerate(headers):
            draw.line((cx, y, cx, y + 26), fill=border, width=1)
            draw.text((cx + 4, y + 4), h, fill=black, font=header_font)
            cx += col_widths[idx]
        draw.line((width - margin, y, width - margin, y + 26), fill=border, width=1)
        draw.line((x, y + 26, width - margin, y + 26), fill=border, width=1)
        y += 26

        rows_sorted = sorted(rows, key=lambda r: (normalize_district(r.district), r.store))
        for row in rows_sorted:
            district = normalize_district(row.district)
            row_fill = district_colors.get(district, (245, 245, 245))
            draw.rectangle((x, y, width - margin, y + row_h), fill=row_fill, outline=border)
            vals = [district, row.store, row.status, row.rep_name or "-"]
            cx = x
            for cidx, val in enumerate(vals):
                if cidx == 2:
                    status_fill = light_green if safe_text(row.status).lower() == "completed" else pending_fill
                    draw.rectangle((cx, y, cx + col_widths[cidx], y + row_h), fill=status_fill, outline=border)
                draw.line((cx, y, cx, y + row_h), fill=border, width=1)
                text_val = safe_text(val) or "-"
                if len(text_val) > 42 and cidx == 3:
                    text_val = text_val[:39] + "..."
                if len(text_val) > 22 and cidx == 1:
                    text_val = text_val[:19] + "..."
                font = cell_bold if cidx in (0, 1, 3) else cell_font
                fill = (210, 0, 0) if (cidx == 2 and safe_text(row.status).lower() != "completed") else black
                draw.text((cx + 4, y + 6), text_val, fill=fill, font=font)
                cx += col_widths[cidx]
            draw.line((width - margin, y, width - margin, y + row_h), fill=border, width=1)
            y += row_h

        safe_title = re.sub(r"[^A-Za-z0-9_-]+", "_", batch_title)[:60] or "Inventory_Audit_Status"
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = IMAGE_DIR / f"GFH_Inventory_Audit_Status_{safe_title}_{stamp}.png"
        img.save(path)
        return path
    @staticmethod
    def format_store_list_caption(stores: List[str]) -> str:
        clean = [safe_text(s) for s in stores if safe_text(s)]
        if not clean:
            return ""
        if len(clean) == 1:
            return clean[0]
        if len(clean) == 2:
            return f"{clean[0]} and {clean[1]}"
        return ", ".join(clean[:-1]) + f", and {clean[-1]}"

    def final_district_caption(self, district: str, rows: List[VarianceRow], pending_status_rows: Optional[List[InventoryStatusRow]] = None) -> str:
        parts: List[str] = []
        dm_phone = self.db.find_district_manager_phone(district)
        if dm_phone:
            parts.append(whatsapp_mention(dm_phone))

        stores = sorted({safe_text(r.store) for r in rows if safe_text(r.store)})
        if stores:
            store_text = self.format_store_list_caption(stores)
            parts.append(f"An Inventory audit has been completed, variance found in {store_text}.")
        else:
            parts.append("An Inventory audit has been completed and no variance found.")

        pending_stores = sorted({safe_text(r.store) for r in (pending_status_rows or []) if safe_text(r.store)})
        if pending_stores:
            pending_text = self.format_store_list_caption(pending_stores)
            verb = "did not complete" if len(pending_stores) == 1 else "did not complete"
            parts.append(f"{pending_text} {verb} an inventory count.")

        return " ".join(parts)

    def _render_no_variance_image(self, district: str) -> Path:
        if Image is None:
            raise RuntimeError("Pillow is required. Install with: py -m pip install pillow")

        renderer = ImageRenderer()
        width = 1360
        margin = 32
        title_font = renderer._font(34, True)
        sub_font = renderer._font(20, False)
        body_font = renderer._font(28, True)
        note_font = renderer._font(18, False)

        dark = (18, 20, 43)
        gray = (95, 95, 102)
        red = (233, 27, 47)
        border = (215, 218, 223)

        logo_img = None
        logo_w = 0
        logo_h_used = 0
        if STATUS_LOGO_PATH.exists():
            try:
                logo_img = Image.open(STATUS_LOGO_PATH).convert("RGBA")
                scale = min(620 / logo_img.width, 150 / logo_img.height)
                size = (max(1, int(logo_img.width * scale)), max(1, int(logo_img.height * scale)))
                logo_img = logo_img.resize(size)
                logo_w, logo_h_used = size
            except Exception:
                logo_img = None
                logo_w = 0
                logo_h_used = 0

        header_area_h = max(logo_h_used, 74) + 18
        box_top = margin + header_area_h + 16
        height = box_top + 190 + 40
        img = Image.new("RGB", (width, height), "white")
        draw = ImageDraw.Draw(img)

        title_text = "GFH Inventory Variance"
        sub_text = f"{district}   Generated: {now_text()}"
        title_bbox = draw.textbbox((0, 0), title_text, font=title_font)
        title_w = title_bbox[2] - title_bbox[0]
        sub_bbox = draw.textbbox((0, 0), sub_text, font=sub_font)
        sub_w = sub_bbox[2] - sub_bbox[0]
        draw.text(((width - title_w) // 2, margin + 4), title_text, fill=dark, font=title_font)
        draw.text(((width - sub_w) // 2, margin + 48), sub_text, fill=gray, font=sub_font)

        if logo_img is not None:
            logo_x = margin
            logo_y = margin + (max(logo_h_used, 74) - logo_h_used) // 2
            img.paste(logo_img, (logo_x, logo_y), logo_img)
            divider_x = margin + logo_w + 14
            draw.rectangle(
                (divider_x, margin, divider_x + 3, margin + max(logo_h_used, 74)), fill=red)

        draw.rectangle((margin, box_top, width - margin, box_top + 160), outline=border, fill=(255, 255, 255))
        draw.rectangle((margin, box_top, width - margin, box_top + 46), fill=red)
        draw.text((margin + 14, box_top + 11), "Final Audit Result", fill="white", font=sub_font)
        draw.text((margin + 18, box_top + 78), "No variance found in this district.", fill=dark, font=body_font)
        draw.text((margin + 18, box_top + 118), "All variances were cleared or none were found.", fill=gray, font=note_font)

        safe_title = re.sub(r"[^A-Za-z0-9_-]+", "_", district)[:60] or "District"
        stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        path = IMAGE_DIR / f"GFH_Final_Audit_{safe_title}_{stamp}.png"
        img.save(path)
        return path
    def mark_selected(self, cleared: bool) -> None:
        rows = self.selected_rows()
        if not rows:
            QMessageBox.information(self, "No selection", "Select one or more variance rows first.")
            return
        def _do():
            for row in rows:
                self.db.set_cleared(row.key, cleared)
        self._db_write(_do)
        self.refresh_table()
        action = "cleared" if cleared else "not cleared"
        self.set_status(f"Marked {len(rows)} variance row(s) as {action}.")

    def clear_current_ui(self, silent: bool = False) -> None:
        self.loaded_keys.clear()
        self.data_loaded = False
        self.status_rows = []
        self.status_row_by_key = {}
        self.status_checked_keys.clear()
        self.audit_checked_keys.clear()
        self.status_key_by_iid.clear()
        self.populate_status_filters()
        self.populate_audit_filters()
        self.refresh_status_table()
        self.refresh_table()
        if not silent:
            self.set_status("UI cleared. Select both file locations, then click Load Variances.")
    def _check_db_sync(self) -> None:
        """Called every 2 s.  If another instance modified the DB, refresh GUI."""
        try:
            if not self._db_sync_paused and DB_PATH.exists():
                mtime = DB_PATH.stat().st_mtime
                if mtime != self._db_mtime:
                    self._db_mtime = mtime
                    # Reload master data in case store/rep lists changed too
                    try:
                        self.master_store_records = self.db.store_master_records()
                    except Exception:
                        pass
                    # Refresh both visible tables silently
                    try:
                        self.refresh_table()
                    except Exception:
                        pass
                    try:
                        self.refresh_status_table()
                    except Exception:
                        pass
        except Exception:
            pass
        self._schedule_db_sync()
    def _db_write(self, fn, *args, **kwargs):
        """Wrap any DB write so the sync poller does not re-trigger on our own write."""
        self._db_sync_paused = True
        try:
            result = fn(*args, **kwargs)
        finally:
            # Update our own mtime baseline after writing, then re-enable polling
            try:
                self._db_mtime = DB_PATH.stat().st_mtime
            except Exception:
                pass
            self._db_sync_paused = False
        return result

    def open_app_folder(self) -> None:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(APP_DIR))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(APP_DIR)])
            else:
                subprocess.Popen(["xdg-open", str(APP_DIR)])
        except Exception as exc:
            QMessageBox.critical(self, "Open folder failed", str(exc))


def _maximize_window(window):
    """Maximize after the window is shown; fall back safely on other window managers."""
    try:
        window.showMaximized()
    except Exception:
        try:
            window.setWindowState(Qt.WindowMaximized)
        except Exception:
            pass


_SINGLE_INSTANCE_MUTEX_HANDLE = None  # module-level ref keeps the OS handle alive for process lifetime


def _acquire_single_instance(mutex_name: str) -> bool:
    """Windows named-mutex single-instance guard.

    Returns True when this process may start (first instance / non-Windows /
    guard failed open); False when another copy is already running — in that
    case main() focuses the existing window and exits quietly.

    Why this exists: the one-file EXE unpacks silently for several seconds,
    so users double-click the icon again while nothing is on screen. Without
    the guard a second identical window opens at the SAME geometry (see
    _apply_dynamic_geometry) and stacks pixel-perfect on top of the first —
    the UI appears to have "double buttons".
    """
    global _SINGLE_INSTANCE_MUTEX_HANDLE
    if not sys.platform.startswith("win"):
        return True
    try:
        import ctypes
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateMutexW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_wchar_p]
        kernel32.CreateMutexW.restype = ctypes.c_void_p
        ERROR_ALREADY_EXISTS = 183
        handle = kernel32.CreateMutexW(None, 0, "Global\\" + mutex_name)
        if not handle:
            return True  # guard could not be created -> fail open, never brick the app
        _SINGLE_INSTANCE_MUTEX_HANDLE = handle
        return ctypes.get_last_error() != ERROR_ALREADY_EXISTS
    except Exception:
        return True


def _focus_existing_window(title_substring: str, exclude: str = "") -> None:
    """Best-effort: restore + foreground the already-running instance's window."""
    if not sys.platform.startswith("win"):
        return
    try:
        import ctypes
        user32 = ctypes.windll.user32
        matches = []

        EnumProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

        def _on_window(hwnd, _lparam):
            try:
                if not user32.IsWindowVisible(hwnd):
                    return True
                buf = ctypes.create_unicode_buffer(512)
                user32.GetWindowTextW(hwnd, buf, 512)
                title = buf.value
                if title_substring and title_substring.lower() in title.lower():
                    if exclude and exclude.lower() in title.lower():
                        return True
                    matches.append(hwnd)
            except Exception:
                pass
            return True

        user32.EnumWindows(EnumProc(_on_window), None)
        if not matches:
            return
        hwnd = matches[0]
        SW_RESTORE = 9
        if user32.IsIconic(hwnd):
            user32.ShowWindow(hwnd, SW_RESTORE)
        user32.BringWindowToTop(hwnd)
        user32.SetForegroundWindow(hwnd)
    except Exception:
        pass



def _enable_dpi_awareness() -> None:
    """AppUserModelID + DPI awareness (Windows).

    Windows groups windows in the taskbar by process/AppUserModelID and
    needs physical-pixel reporting on high-DPI displays; without the
    shcore call the UI would render blurry on DPI-scaled laptops."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("GFH.InventoryAuditTimesheet")
        except Exception:
            pass
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # system DPI aware
        except Exception:
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def _vp_hide(win):
    """ANTI-FLASH: keep the window fully transparent while the UI is built, so a
    half-built window never flashes behind the boot splash."""
    try:
        win.setWindowOpacity(0.0)
    except Exception:
        pass
    try:
        import pyi_splash  # type: ignore
        pyi_splash.update_text("Loading interface...")
    except Exception:
        pass


def _vp_reveal(win):
    """Show the finished window in one step (called just before the splash closes)."""
    try:
        win.setWindowOpacity(1.0)
        win.raise_()
        win.activateWindow()
    except Exception:
        pass


def _close_pyi_splash():
    """Close the PyInstaller splash screen (frozen onefile builds only).

    Called once the main window is fully built and shown; harmless no-op
    in dev runs (pyi_splash only exists inside PyInstaller builds)."""
    try:
        import pyi_splash  # type: ignore
        pyi_splash.close()
    except Exception:
        pass


def show_startup_error(exc: BaseException) -> None:
    error_text = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    try:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        (APP_DIR / "startup_error.log").write_text(error_text, encoding="utf-8")
    except Exception:
        try:
            Path(tempfile.gettempdir(), "gfh_inventory_audit_startup_error.log").write_text(error_text, encoding="utf-8")
        except Exception:
            pass
    try:
        from PySide6.QtWidgets import QApplication, QMessageBox
        _app = QApplication.instance() or QApplication(sys.argv)
        QMessageBox.critical(None, "GFH Inventory Audit Error", str(exc))
    except Exception:
        print("GFH Inventory Audit Error:", exc)
        print(error_text)


def main() -> None:
    # ── Single-instance guard ───────────────────────────────────────────────
    # Onefile EXE unpacks silently for seconds; a second double-click used to
    # open a second identical window at the same geometry, stacked pixel-perfect
    # on top of the first — the user saw this as "double buttons". Instead of
    # launching a duplicate, focus the already-running window and exit quietly.
    if not _acquire_single_instance("GFHTelecom.TimesheetAudit.SingleInstance"):
        _focus_existing_window("Timesheet Edition")
        return

    _enable_dpi_awareness()
    app = QApplication(sys.argv)
    _check_required_packages(app)
    win = MainWindow()
    win.show()
    _vp_reveal(win)
    _close_pyi_splash()
    sys.exit(app.exec())


if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    try:
        main()
    except Exception as exc:
        traceback.print_exc()
        show_startup_error(exc)
        if sys.stdin and sys.stdin.isatty():
            input("Press Enter to close...")
