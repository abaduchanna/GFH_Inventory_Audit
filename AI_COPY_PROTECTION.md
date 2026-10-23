# 3SVerse AI-Copy Protection (t192)

Is app me ab 4-layer protection hai — koi AI (reverse-engineering skill ke sath bhi)
iska working copy NAHI bana sakta:

| Layer | Kya hai | AI ke against kya block karta hai |
|---|---|---|
| 1. Native compile | `_protect_build.py` build se PEHLE saare modules ko Cython se machine code (.pyd) me compile karta hai, phir .py sources delete. EXE me Python source/bytecode NahiN hota | pyinstxtractor + decompiler = sirf machine code, source kabhi nahi |
| 2. License gate | `license_core.ensure_licensed()` launch pe: HWID one-PC activation, 7-day trial (anti-rollback dual-store), Ed25519 signed .lic | EXE copy karke doosre PC pe chalana = activation fail |
| 3. Heartbeat + revocation | Ledger (GitHub private repo) se check; refunded/revoked license = app dead | Leak hone par bhi kill-switch kaam karta hai |
| 4. Anti-debug / anti-VM | Frozen build me debugger attach ya VM detect = silent exit | Dynamic analysis / sandbox cloning |

## Build pe kya chalana hai (koi change nahi)
```
python _protect_build.py     # native compile (pehle se wired)
<existing build .bat>        # PyInstaller jaisa pehle chalta tha
```

## License issue karna (customer ko)
- License Studio / make_license_token.py se product `PRODUCT_ID` ke liye .lic banao
- Trial: har naye PC pe 7 din automatic — phir .lic zaroori

## APK note (future)
Android side ke liye: R8/ProGuard full obfuscation + Play Integrity API + license
server par wahi HWID scheme. Template jab APK banao tab add hoga.
