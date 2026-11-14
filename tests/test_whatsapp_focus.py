"""The bot must never un-maximize the WhatsApp Desktop window when focusing it.

Regression test for: "why whatsapp window layout changed when bot focus it to
send text" — the WhatsApp focus path used to call ShowWindow(hwnd, SW_RESTORE)
unconditionally. On Windows, SW_RESTORE applied to a MAXIMIZED window
un-maximizes it, so WhatsApp visibly re-flowed its whole layout on every send.

The invariant (2026-11, after the focus logic moved from the
WhatsAppSender._force_focus_whatsapp method into the module-level helpers
_wa_ensure_focused / _wa_focus_hwnd):

- A layout-breaking ShowWindow (SW_RESTORE=9 / SW_SHOWNORMAL=1) is allowed
  ONLY in one of two safe positions:
    1. behind an IsIconic(hwnd) gate (restore a MINIMIZED window only), or
    2. as the second half of the last-resort minimise->restore escalation
       (SW_MINIMIZE=6 ... SW_RESTORE=9 restores the window's PREVIOUS state,
       so a maximized window comes back maximized).
- State-preserving calls (SW_SHOW=5 etc.) are always fine.
- SetForegroundWindow must still be called.
"""
import ast
import pathlib

import pytest

REPO = pathlib.Path(__file__).resolve().parent.parent
APP_FILES = [
    REPO / "GFH_Inventory_Audit_Timesheet.py",
    REPO / "GFH_Inventory_Audit.py",
]

FOCUS_SCOPES = ("_force_focus_whatsapp", "_wa_ensure_focused", "_wa_focus_hwnd")

UNSAFE_VALUES = {1, 9}   # SW_SHOWNORMAL, SW_RESTORE — both un-maximize
UNSAFE_NAMES = {"SW_SHOWNORMAL", "SW_RESTORE", "SW_NORMAL"}
MINIMIZE_VALUE = 6       # SW_MINIMIZE
MINIMIZE_NAMES = {"SW_MINIMIZE"}


def _focus_scopes(tree: ast.AST):
    """Every function in the WhatsApp focus chain: the wrapper method and
    the module-level helpers it delegates to."""
    out = []
    for node in ast.walk(tree):
        if isinstance(node, ast.FunctionDef) and node.name in FOCUS_SCOPES:
            out.append(node)
        if isinstance(node, ast.ClassDef) and node.name == "WhatsAppSender":
            for item in node.body:
                if isinstance(item, ast.FunctionDef) and item.name in FOCUS_SCOPES:
                    out.append(item)
    return out


def _calls(node: ast.AST, name: str):
    out = []
    for sub in ast.walk(node):
        if isinstance(sub, ast.Call):
            fn = sub.func
            if (isinstance(fn, ast.Attribute) and fn.attr == name) or (
                isinstance(fn, ast.Name) and fn.id == name
            ):
                out.append(sub)
    return out


def _contains_call(node: ast.AST, name: str) -> bool:
    return bool(_calls(node, name))


def _showwindow_arg(call) -> object:
    """Second argument of a ShowWindow(hwnd, X) call as (value, name)."""
    if len(call.args) < 2:
        return None
    arg = call.args[1]
    if isinstance(arg, ast.Constant):
        return ("value", arg.value)
    if isinstance(arg, ast.Attribute):
        return ("name", arg.attr)
    return ("other", None)


def _showwindow_events(fn: ast.FunctionDef):
    """(value, name, is_gated) for every ShowWindow call in source order.

    Source order matters: the escalation restore is only valid if a
    SW_MINIMIZE precedes it. ast.NodeVisitor visits children in field
    order, which is source order.
    """
    gated_ids = set()
    for n in ast.walk(fn):
        if isinstance(n, ast.If) and _contains_call(n.test, "IsIconic"):
            for sub in ast.walk(n):
                gated_ids.add(id(sub))

    events = []

    class V(ast.NodeVisitor):
        def visit_Call(self, node):
            fn_attr = node.func
            is_show = (
                (isinstance(fn_attr, ast.Attribute) and fn_attr.attr == "ShowWindow")
                or (isinstance(fn_attr, ast.Name) and fn_attr.id == "ShowWindow")
            )
            if is_show:
                kind, value = _showwindow_arg(node) or ("other", None)
                events.append((kind, value, id(node) in gated_ids))
            self.generic_visit(node)

    V().visit(fn)
    return events


@pytest.mark.parametrize("app", APP_FILES, ids=lambda p: p.name)
def test_showwindow_restore_only_when_minimized(app):
    tree = ast.parse(app.read_text(encoding="utf-8"))
    scopes = _focus_scopes(tree)
    assert scopes, f"no WhatsApp focus function found in {app.name}"

    iconic_gates = 0
    setfg_calls = 0
    for fn in scopes:
        # 1) An IsIconic gate must exist somewhere in the focus chain.
        iconic_gates += sum(
            1 for n in ast.walk(fn)
            if isinstance(n, ast.If) and _contains_call(n.test, "IsIconic")
        )
        setfg_calls += len(_calls(fn, "SetForegroundWindow"))

        events = _showwindow_events(fn)
        minimized_seen = False
        for kind, value, gated in events:
            unsafe = (kind == "value" and value in UNSAFE_VALUES) or (
                kind == "name" and value in UNSAFE_NAMES
            )
            if kind == "value" and value == MINIMIZE_VALUE:
                minimized_seen = True
            if kind == "name" and value in MINIMIZE_NAMES:
                minimized_seen = True
            if unsafe:
                assert gated or minimized_seen, (
                    "layout-breaking ShowWindow(SW_RESTORE/SW_SHOWNORMAL) found "
                    "outside an IsIconic gate and outside the "
                    "minimise->restore escalation — it un-maximizes a maximized "
                    "WhatsApp window and re-flows its layout"
                )

    assert iconic_gates, "ShowWindow(SW_RESTORE) must be gated behind IsIconic(hwnd)"
    assert setfg_calls, "SetForegroundWindow must still be called"
