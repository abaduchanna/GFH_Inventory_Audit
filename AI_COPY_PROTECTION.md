# 3SVerse AI-Copy Protection (t192 / t195 update — FREE app)

Yeh app FREE hai (owner ka internal GFH tool) — **trial/license gate NAHI hai**.
App kholte hi seedha kaam shuru, koi activation/trial popup nahi.

Copy protection ab sirf build-level layer pe hai:

| Layer | Kya hai | AI ke against kya block karta hai |
|---|---|---|
| 2. Runtime guard | `runtime_guard.enforce()` launch pe (sirf frozen EXE): debugger attach = silent exit, koi popup nahi | Live dynamic analysis (x64dbg/x32dbg/unpacking debuggers) = app band |
| 1. Native compile | `_protect_build.py` build se PEHLE saare modules ko Cython se machine code (.pyd) me compile karta hai, phir .py sources delete. EXE me Python source/bytecode NahiN hota | pyinstxtractor + decompiler = sirf machine code, source kabhi nahi |

(t195: license gate + trial + heartbeat layers is FREE app se REMOVE ho gaye —
trial system sirf sell-line apps (VidaPay_* / VDS_*) ke liye reserved hai.
Owner directive 2026-10-24.)

## Build pe kya chalana hai (koi change nahi)
```
python _protect_build.py     # native compile (pehle se wired)
<existing build .bat>        # PyInstaller jaisa pehle chalta tha
```

## APK note (future)
Android side ke liye: R8/ProGuard full obfuscation + Play Integrity API.
Template jab APK banao tab add hoga.
