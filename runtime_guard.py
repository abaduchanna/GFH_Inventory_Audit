#!/usr/bin/env python3
"""3SVerse runtime tamper guard (FREE apps) - t197.

Sirf frozen EXE builds main active. Debugger attach detect hote hi app
silently band (koi popup nahi). Yeh licensing NAHI hai - sirf dynamic
analysis (live debuggers) ko rokta hai. Static protection native-compile
(_protect_build.py) se aati hai: pyinstxtractor + decompiler ko sirf
machine code milta hai, source kabhi nahi.

AI-copy ke against combined defense:
  1. native compile (_protect_build.py) -> EXE main Python source/bytecode hi nahi
  2. runtime guard (yeh file)           -> live debugger attach = silent exit
Dev runs (python app.py) main yeh guard NO-OP hai.
"""
import os
import sys


def debugger_present():
    """Interactive debugger attached? Windows only."""
    if os.name != "nt":
        return False
    try:
        import ctypes
        k32 = ctypes.windll.kernel32
        if k32.IsDebuggerPresent():
            return True
        out = ctypes.c_ulong(0)
        k32.CheckRemoteDebuggerPresent(k32.GetCurrentProcess(),
                                       ctypes.byref(out))
        return bool(out.value)
    except Exception:
        return False


def enforce():
    """Launch-time gate - frozen EXE only, silent exit on debugger."""
    if getattr(sys, "frozen", False) and os.name == "nt" and debugger_present():
        os._exit(0)
