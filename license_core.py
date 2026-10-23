"""
3SVerse Licensing Core
Developed by www.3SVerse.com

Machine-locked activation for VidaPay desktop tools.

  * Offline license verification  -> Ed25519 signature embedded public key
  * Online activation binding     -> one .lic file activates on ONE PC only
  * 7-day trial per machine       -> dual-store anti-rollback + server sync
  * Heartbeat                     -> revoked/refunded licenses stop working
  * Anti-debug / anti-VM          -> active in frozen (exe) builds

Self-contained: python stdlib only. This module is copied into every app
repo; the placeholders below are filled per-product by the patch script.
"""

import base64
import datetime as _dt
import hashlib
import hmac as _hmac
import json
import os
import struct
import sys
import threading
import time
import urllib.error
import urllib.request

try:
    import winreg as _winreg
except Exception:
    _winreg = None



def _exit_note(reason):
    """Leave a one-line trace before a deliberate silent exit (t133/t134).
    Writes to %LOCALAPPDATA%\\VidaPay\\logs\\<app>\\crash-<stamp>-exit.txt so a
    vanished app can always be told apart from a real crash. Folder carries
    the app name so reports never mix between tools."""
    try:
        from pathlib import Path as _P
        _folders = {"extractor": "Incentive_Extractor", "rebate": "Rebate_Filing", "ordering": "Device_Ordering", "transferbot": "Transfer_Bot", "vds-rebate": "VDS_Rebate", "vds-ups": "VDS_UPS", "vds-xls": "VDS_Xls", "gfh-audit": "GFH_Audit", "gfh-inv": "GFH_Inv", "gfh-rebate": "GFH_Rebate", "gfh-acc": "GFH_Acc", "gfh-ups": "GFH_UPS"}
        _pid = globals().get("PRODUCT_ID", "") or "license"
        folder = _folders.get(_pid, _pid)
        base = os.environ.get("LOCALAPPDATA") or str(_P.home())
        d = _P(base) / "VidaPay" / "logs" / folder
        d.mkdir(parents=True, exist_ok=True)
        stamp = time.strftime("%Y%m%d-%H%M%S")
        with open(d / ("crash-%s-exit.txt" % stamp), "a",
                  encoding="utf-8", errors="replace") as f:
            f.write("[%s] DELIBERATE EXIT (license_core): %s\n"
                    % (time.strftime("%Y-%m-%d %H:%M:%S"), reason))
    except Exception:
        pass

# ---------------------------------------------------------------------------
# Per-product configuration (filled by patch script)
# ---------------------------------------------------------------------------
APP_NAME = "GFH Inventory Audit TS"
PRODUCT_ID = "gfh-inv"

# Ed25519 public keys (hex). The PRIVATE signing keys NEVER ship with apps:
#   license key  -> held by the seller's keygen, signs .lic files
#   server key   -> held by the activation server, signs activation receipts
LICENSE_PUB_HEX = "e6dce9860f5e9c995dcba674da052617e9d37c00e238c8174c5d3470de866937"
SERVER_PUB_HEX = "2923d11cf704853407bfa625d3ec7fb9686b9911e651538687a3635f5203347b"

# License ledger (GitHub-hosted, PRIVATE repo). All online licensing -
# one-PC binding, revocation, trial anti-reset - runs through the GitHub
# Contents API with a scoped token embedded at build time:
#   make_license_token.py -> license_token.py (generated, gitignored).
LEDGER_REPO = "abaduchanna/3SVerse_License_Server"
_TOKEN_XK = b"3SVerse-Ledger-Token-v1"

TRIAL_DAYS = 7
HEARTBEAT_DAYS = 7           # legacy constant; launch now re-verifies every start (see ensure_licensed)
ACTIVATION_TIMEOUT = 8       # seconds
TRIAL_SYNC_TIMEOUT = 4       # seconds - trial sync is best-effort, keep startup snappy
_BRAND = "3S Verse (www.3SVerse.com)"
_SUPPORT_EMAIL = "Connect@3SVerse.com"


# ===========================================================================
# Ed25519 (RFC 8032) - pure python, stdlib only
# ===========================================================================
_q = 2 ** 255 - 19
_l = 2 ** 252 + 27742317777372353535851937790883648493


def _H(m):
    return hashlib.sha512(m).digest()


def _inv(x):
    return pow(x, _q - 2, _q)


_d = (-121665 * _inv(121666)) % _q
_I = pow(2, (_q - 1) // 4, _q)


def _xrecover(y):
    xx = (y * y - 1) * _inv(_d * y * y + 1)
    x = pow(xx, (_q + 3) // 8, _q)
    if (x * x - xx) % _q != 0:
        x = (x * _I) % _q
    if x % 2 != 0:
        x = _q - x
    return x


_By = (4 * _inv(5)) % _q
_Bx = _xrecover(_By)
_B = (_Bx % _q, _By % _q, 1, (_Bx * _By) % _q)


def _edwards_add(P, Q):
    (X1, Y1, Z1, T1) = P
    (X2, Y2, Z2, T2) = Q
    A = ((Y1 - X1) * (Y2 - X2)) % _q
    B = ((Y1 + X1) * (Y2 + X2)) % _q
    C = (T1 * 2 * _d * T2) % _q
    D = (Z1 * 2 * Z2) % _q
    E = B - A
    F = D - C
    G = D + C
    Hh = B + A
    return ((E * F) % _q, (G * Hh) % _q, (F * G) % _q, (E * Hh) % _q)


def _scalarmult(P, e):
    Q = (0, 1, 1, 0)
    while e > 0:
        if e & 1:
            Q = _edwards_add(Q, P)
        P = _edwards_add(P, P)
        e >>= 1
    return Q


def _encodepoint(P):
    (X, Y, Z, T) = P
    zi = _inv(Z)
    x = X * zi % _q
    y = Y * zi % _q
    return (y | ((x & 1) << 255)).to_bytes(32, "little")


def _decodepoint(s):
    yint = int.from_bytes(s, "little")
    x_sign = yint >> 255
    y = yint & ((1 << 255) - 1)
    if y >= _q:
        raise ValueError("point out of field")
    x = _xrecover(y)
    if x & 1 != x_sign:
        x = _q - x
    P = (x, y, 1, (x * y) % _q)
    (X, Y, Z, T) = P
    if (Y * Y - X * X - Z * Z - _d * T * T) % _q != 0:
        raise ValueError("point not on curve")
    return P


def _hint(m):
    return int.from_bytes(_H(m), "little")


def _secret_expand(sk):
    h = _H(sk)
    a = int.from_bytes(h[:32], "little")
    a &= (1 << 254) - 8
    a |= (1 << 254)
    return a, h[32:]


def ed25519_publickey(sk):
    a, _prefix = _secret_expand(sk)
    return _encodepoint(_scalarmult(_B, a))


def ed25519_sign(msg, sk):
    pk = ed25519_publickey(sk)
    a, prefix = _secret_expand(sk)
    r = _hint(prefix + msg) % _l
    R = _encodepoint(_scalarmult(_B, r))
    k = _hint(R + pk + msg) % _l
    S = (r + k * a) % _l
    return R + S.to_bytes(32, "little")


def ed25519_verify(sig, msg, pk):
    if len(sig) != 64 or len(pk) != 32:
        return False
    try:
        R = _decodepoint(sig[:32])
        A = _decodepoint(pk)
        S = int.from_bytes(sig[32:], "little")
        if S >= _l:
            return False
        k = _hint(sig[:32] + pk + msg) % _l
        left = _scalarmult(_B, S)
        right = _edwards_add(R, _scalarmult(A, k))
        return _encodepoint(left) == _encodepoint(right)
    except Exception:
        return False


def _unhex(s):
    return bytes.fromhex("".join(s.split()))


# ===========================================================================
# Machine identity
# ===========================================================================
def raw_machine_id():
    """Stable per-OS-install identifier."""
    if _winreg is not None:
        try:
            k = _winreg.OpenKey(_winreg.HKEY_LOCAL_MACHINE,
                                r"SOFTWARE\Microsoft\Cryptography")
            mid, _ = _winreg.QueryValueEx(k, "MachineGuid")
            _winreg.CloseKey(k)
            if mid:
                return str(mid)
        except Exception:
            pass
    for p in ("/etc/machine-id", "/var/lib/dbus/machine-id"):
        try:
            if os.path.exists(p):
                return open(p).read().strip()
        except Exception:
            pass
    import platform
    return platform.node() or "unknown-machine"


def machine_binding_id():
    """Displayable machine fingerprint: 3SV-XXXX-XXXX-XXXX-XXXX"""
    digest = hashlib.sha256(("vplic:" + raw_machine_id()).encode("utf-8")).digest()
    code = base64.b32encode(digest[:10]).decode("ascii").rstrip("=")
    groups = "-".join(code[i:i + 4] for i in range(0, len(code), 4))
    return "3SV-" + groups


def machine_evidence():
    """Human-checkable evidence for the seller (hostname + primary MAC).
    NOT part of the binding - the 3SV fingerprint is the binding identity;
    this is stored in the ledger so the seller can see WHICH PC (and MAC)
    holds the license."""
    host = ""
    try:
        import socket
        host = (socket.gethostname() or "").strip()[:40]
    except Exception:
        pass
    mac = ""
    try:
        import uuid as _uuid
        node = _uuid.getnode()
        if node and not (node & 0x010000000000):
            mac = "%012X" % (node & 0xFFFFFFFFFFFF)
    except Exception:
        pass
    mac_txt = "-".join(mac[i:i + 4] for i in range(0, len(mac), 4)) or "n/a"
    return "%s | MAC %s" % (host or "unknown-pc", mac_txt)


# ===========================================================================
# Machine-keyed local encryption (activation records / trial stamps)
# ===========================================================================
_SALT = b"3SVerse::License::v1::" + PRODUCT_ID.encode("utf-8")


def _stor_key():
    return hashlib.sha256(_SALT + raw_machine_id().encode("utf-8")).digest()


def _keystream(nonce, length):
    out = bytearray()
    counter = 0
    key = _stor_key()
    while len(out) < length:
        block = _hmac.new(key, nonce + struct.pack(">I", counter),
                          hashlib.sha256).digest()
        out.extend(block)
        counter += 1
    return bytes(out[:length])


def _encrypt(data_bytes):
    nonce = hashlib.sha256(raw_machine_id().encode() + _SALT).digest()[:16]
    ks = _keystream(nonce, len(data_bytes))
    enc = bytes(a ^ b for a, b in zip(data_bytes, ks))
    mac = _hmac.new(_stor_key(), nonce + enc, hashlib.sha256).digest()
    return base64.b64encode(nonce + mac + enc)


def _decrypt(blob):
    try:
        raw = base64.b64decode(blob)
        nonce, mac, enc = raw[:16], raw[16:48], raw[48:]
        if not _hmac.compare_digest(
                mac, _hmac.new(_stor_key(), nonce + enc,
                               hashlib.sha256).digest()):
            return None
        ks = _keystream(nonce, len(enc))
        return bytes(a ^ b for a, b in zip(enc, ks))
    except Exception:
        return None


def _base_dir():
    if os.name == "nt":
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
        d = os.path.join(base, "VidaPay", "License")
    else:
        d = os.path.join(os.path.expanduser("~"), ".vidapay", "license")
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass
    return d


def _hide_file(path):
    if os.name != "nt":
        return
    try:
        import ctypes
        ctypes.windll.kernel32.SetFileAttributesW(str(path), 0x02)  # hidden
    except Exception:
        pass

def _lic_log(line):
    """Best-effort one-line trace for activation support (same folder as
    the crash reports). Never raises."""
    try:
        from pathlib import Path as _P
        _folders = {"extractor": "Incentive_Extractor",
                    "rebate": "Rebate_Filing",
                    "ordering": "Device_Ordering"}
        _pid = globals().get("PRODUCT_ID", "") or "license"
        base = os.environ.get("LOCALAPPDATA") or str(_P.home())
        d = _P(base) / "VidaPay" / "logs" / _folders.get(_pid, _pid)
        d.mkdir(parents=True, exist_ok=True)
        with open(d / "license.log", "a",
                  encoding="utf-8", errors="replace") as f:
            f.write("[%s] %s\n"
                    % (time.strftime("%Y-%m-%d %H:%M:%S"), line))
    except Exception:
        pass


# ===========================================================================
# License files: load + validate (Ed25519 signed by the seller's keygen)
# ===========================================================================
def license_payload_bytes(lic):
    core = {k: v for k, v in lic.items() if k != "sig"}
    return json.dumps(core, sort_keys=True, separators=(",", ":")).encode("utf-8")


# ---------------------------------------------------------------------------
# Seller-published signing keys (ledger/config.json) - key rotation WITHOUT
# app rebuilds. A license signature is accepted when it verifies against the
# BUILT-IN public key OR any key the seller published from License Studio
# (Settings -> Publish Public Key to Ledger). Apps refresh the published-key
# list online (cached on disk with a TTL, last-known list kept when the
# ledger is unreachable), so rotating the master key never requires a new
# build: publish from Studio and every installed app accepts the new keys.
# ---------------------------------------------------------------------------
PUB_CACHE_TTL = 6 * 3600  # seconds between online refreshes


def _pub_cache_path():
    return os.path.join(_base_dir(), "seller_pub_keys.json")


def _seller_pub_keys_remote(timeout=ACTIVATION_TIMEOUT):
    """-> list of 64-char pub hexes published in ledger/config.json
    ([] when the file does not exist yet)."""
    st, body = _gh("GET",
                   "/repos/%s/contents/ledger/config.json" % LEDGER_REPO,
                   timeout=timeout)
    if st != 200:
        return []
    cfg = json.loads(base64.b64decode(body["content"]).decode("utf-8")) or {}
    out = []
    for k in (cfg.get("pub_keys") or []):
        h = "".join(str(k).split()).lower()
        if len(h) != 64:
            continue
        try:
            bytes.fromhex(h)
        except Exception:
            continue
        if h not in out:
            out.append(h)
    return out


def accepted_pub_keys():
    """[LICENSE_PUB_HEX] + seller-published keys (disk cache + online)."""
    pubs = [LICENSE_PUB_HEX]
    cache = {"at": 0, "pubs": []}
    try:
        with open(_pub_cache_path(), "r", encoding="utf-8") as f:
            cache = json.load(f) or cache
    except Exception:
        pass
    cached = [p for p in cache.get("pubs", []) if p not in pubs]
    if time.time() - float(cache.get("at") or 0) < PUB_CACHE_TTL:
        return pubs + cached
    if not _ledger_ready():
        return pubs + cached          # trial/dev build - embedded key only
    try:
        remote = _seller_pub_keys_remote()
    except Exception:
        return pubs + cached          # ledger unreachable - last known list
    try:
        with open(_pub_cache_path(), "w", encoding="utf-8") as f:
            json.dump({"at": time.time(), "pubs": remote}, f)
    except Exception:
        pass
    return pubs + [p for p in remote if p not in pubs]


def license_signature_ok(lic):
    try:
        sig = base64.b64decode(lic.get("sig", ""))
        msg = license_payload_bytes(lic)
        for pub_hex in accepted_pub_keys():
            try:
                if ed25519_verify(sig, msg, _unhex(pub_hex)):
                    return True
            except Exception:
                continue
        return False
    except Exception:
        return False


def load_license_file(path):
    with open(path, "r", encoding="utf-8-sig") as f:
        lic = json.load(f)
    if not isinstance(lic, dict):
        raise ValueError("license file is not valid")
    return lic


def validate_license(lic):
    """Return (ok, reason). reason is None on success."""
    if not isinstance(lic, dict):
        return False, "License file is not valid."
    if not license_signature_ok(lic):
        return False, "License signature is INVALID (file edited or fake)."
    products = lic.get("products") or []
    if "bundle" not in products and PRODUCT_ID not in products:
        return False, ("This license is for a different product "
                       "(needed: %s)." % PRODUCT_ID)
    bound = str(lic.get("machine") or "").strip()
    if bound and bound != machine_binding_id():
        return False, ("This license was issued for another PC "
                       "(Machine ID mismatch).")
    if lic.get("model") == "expiry":
        exp = str(lic.get("expires") or "").strip()
        try:
            exp_date = _dt.date.fromisoformat(exp)
        except Exception:
            return False, "License has an invalid expiry date."
        if _dt.date.today() > exp_date:
            return False, ("License EXPIRED on %s. Renew it with your "
                           "seller." % exp)
    return True, None


# ===========================================================================
# Trial: 7 days per machine, dual store, anti clock-rollback
# ===========================================================================
def _trial_reg_read():
    if _winreg is None:
        return None, None
    try:
        k = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER,
                            r"Software\VidaPay\License\%s" % PRODUCT_ID)
        fr, _ = _winreg.QueryValueEx(k, "tfr")
        ls, _ = _winreg.QueryValueEx(k, "tls")
        _winreg.CloseKey(k)
        return _deobf(fr), _deobf(ls)
    except Exception:
        return None, None


def _trial_reg_write(fr, ls):
    if _winreg is None:
        return
    try:
        k = _winreg.CreateKey(_winreg.HKEY_CURRENT_USER,
                              r"Software\VidaPay\License\%s" % PRODUCT_ID)
        _winreg.SetValueEx(k, "tfr", 0, _winreg.REG_SZ, _obf(fr))
        _winreg.SetValueEx(k, "tls", 0, _winreg.REG_SZ, _obf(ls))
        _winreg.CloseKey(k)
    except Exception:
        pass


def _trial_file_read():
    try:
        p = os.path.join(_base_dir(), ".vptrial_%s" % PRODUCT_ID)
        if os.path.exists(p):
            obj = json.loads(open(p, "r").read())
            return _deobf(obj.get("fr", "")), _deobf(obj.get("ls", ""))
    except Exception:
        pass
    return None, None


def _trial_file_write(fr, ls):
    try:
        p = os.path.join(_base_dir(), ".vptrial_%s" % PRODUCT_ID)
        with open(p, "w") as f:
            f.write(json.dumps({"fr": _obf(fr), "ls": _obf(ls)}))
        _hide_file(p)
    except Exception:
        pass


def _obf(num):
    return base64.b64encode(_encrypt(str(int(num)).encode())).decode()


def _deobf(text):
    try:
        dec = _decrypt(base64.b64decode(text.encode()))
        return int(dec.decode()) if dec else None
    except Exception:
        return None


def trial_status():
    """Return (expired, days_left_float). Anti-rollback enforced.

    Reads the earliest first-run stamp across ALL independent stores
    (registry, hidden user file, hidden ProgramData file, decoy registry
    value) so wiping one or two of them cannot reset or extend the 7-day
    clock. The server ledger record (server_trial_sync) is the outer
    guard on top of these.
    """
    now = int(time.time())
    firsts, lasts = _trial_all_read()
    if not firsts:
        _trial_all_write(now, now)
        return False, float(TRIAL_DAYS)
    first_run = min(firsts)
    last_seen = max(lasts) if lasts else first_run
    effective_now = max(now, last_seen)
    _trial_all_write(first_run, effective_now)
    remaining_secs = (TRIAL_DAYS * 86400.0) - (effective_now - first_run)
    if remaining_secs <= 0:
        return True, 0.0
    return False, remaining_secs / 86400.0


def _trial_label(days_left):
    """Human label for the remaining trial time. Never shows '0 days
    left' while the trial is still running - the final day is expressed
    in hours, so a running trial is always visibly distinguishable from
    an expired one."""
    try:
        secs = float(days_left) * 86400.0
    except Exception:
        secs = 0.0
    if secs <= 0:
        return "0 - trial ended"
    if secs >= 2 * 86400:
        return "%d days left" % int(secs / 86400.0 + 0.999)
    if secs >= 86400:
        return "1 day left"
    hours = int(secs / 3600.0 + 0.999)
    if hours >= 2:
        return "%d hours left" % hours
    minutes = max(1, int(secs / 60.0 + 0.999))
    return "less than 1 hour left (%d min)" % minutes


def adopt_trial_start(first_seen):
    """Adopt an earlier (server-provided) trial start, healing all stores."""
    if not first_seen:
        return
    firsts, lasts = _trial_all_read()
    firsts.append(int(first_seen))
    first_run = min(firsts)
    last_seen = max(lasts) if lasts else first_run
    _trial_all_write(first_run, last_seen)


def _trial_all_read():
    """Merge every independent trial store. Returns (firsts, lasts) lists
    of successfully decoded stamps; unreadable/tampered stores simply
    contribute nothing instead of failing the trial."""
    firsts, lasts = [], []
    fr_r, ls_r = _trial_reg_read()
    if fr_r:
        firsts.append(int(fr_r))
        if ls_r:
            lasts.append(int(ls_r))
    fr_f, ls_f = _trial_file_read()
    if fr_f:
        firsts.append(int(fr_f))
        if ls_f:
            lasts.append(int(ls_f))
    fr_p, ls_p = _trial_progdata_read()
    if fr_p:
        firsts.append(int(fr_p))
        if ls_p:
            lasts.append(int(ls_p))
    fr_d, ls_d = _trial_decoy_read()
    if fr_d:
        firsts.append(int(fr_d))
        if ls_d:
            lasts.append(int(ls_d))
    return firsts, lasts


def _trial_all_write(fr, ls):
    """Write the merged stamps to EVERY store; individual failures are
    ignored (a store that cannot be written also cannot be read back)."""
    _trial_reg_write(fr, ls)
    _trial_file_write(fr, ls)
    _trial_progdata_write(fr, ls)
    _trial_decoy_write(fr, ls)


def _trial_progdata_path():
    """All-users location (survives HKCU cleanup tools). Same per-product
    naming as the user store."""
    if os.name == "nt":
        base = os.environ.get("PROGRAMDATA") or os.environ.get("LOCALAPPDATA")
        d = os.path.join(base or _base_dir(), "VidaPay", "License")
    else:
        d = os.path.join(_base_dir(), "pd")
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass
    return os.path.join(d, ".vptrial_%s" % PRODUCT_ID)


def _trial_progdata_read():
    try:
        p = _trial_progdata_path()
        if os.path.exists(p):
            obj = json.loads(open(p, "r").read())
            return _deobf(obj.get("fr", "")), _deobf(obj.get("ls", ""))
    except Exception:
        pass
    return None, None


def _trial_progdata_write(fr, ls):
    try:
        p = _trial_progdata_path()
        with open(p, "w") as f:
            f.write(json.dumps({"fr": _obf(fr), "ls": _obf(ls)}))
        _hide_file(p)
    except Exception:
        pass


def _trial_decoy_read():
    """Decoy registry value under an unremarkable key - survives registry
    cleaners that target the obvious Software\\VidaPay branch."""
    if _winreg is None:
        return None, None
    try:
        k = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER,
                            r"Software\AppDatal\Persistence\%s" % PRODUCT_ID)
        fr, _ = _winreg.QueryValueEx(k, "mfr")
        ls, _ = _winreg.QueryValueEx(k, "mls")
        _winreg.CloseKey(k)
        return _deobf(fr), _deobf(ls)
    except Exception:
        return None, None


def _trial_decoy_write(fr, ls):
    if _winreg is None:
        return
    try:
        k = _winreg.CreateKey(_winreg.HKEY_CURRENT_USER,
                              r"Software\AppDatal\Persistence\%s" % PRODUCT_ID)
        _winreg.SetValueEx(k, "mfr", 0, _winreg.REG_SZ, _obf(fr))
        _winreg.SetValueEx(k, "mls", 0, _winreg.REG_SZ, _obf(ls))
        _winreg.CloseKey(k)
    except Exception:
        pass


def _ledger_put_path(path, rec, sha=None, timeout=TRIAL_SYNC_TIMEOUT):
    """Compare-and-swap write to an explicit ledger path (trial sync)."""
    body = json.dumps(rec, sort_keys=True,
                      separators=(",", ":")).encode("utf-8")
    payload = {"message": "trial sync",
               "content": base64.b64encode(body).decode("ascii")}
    if sha:
        payload["sha"] = sha
    return _gh("PUT", path, payload, timeout=timeout)


def trial_server_sync():
    """Sync the trial window with the ledger (anti reinstall-reset).

    Best effort by design - the trial must work fully offline:
      - ledger not configured / offline / any error -> local state stands
      - ledger HAS a record for this machine -> adopt the earliest
        first_seen (heals a reinstall that wiped local state)
      - ledger has NO record -> create one with the local first run so
        future reinstalls cannot extend the trial

    Never raises.
    """
    try:
        if not _ledger_ready():
            return
        path = ("/repos/%s/contents/ledger/trials/%s-%s.json"
                % (LEDGER_REPO, PRODUCT_ID, machine_binding_id()))
        fr_r, _ls_r = _trial_reg_read()
        fr_f, _ls_f = _trial_file_read()
        firsts = [int(x) for x in (fr_r, fr_f) if x]
        local_first = min(firsts) if firsts else int(time.time())
        st, body = _gh("GET", path, timeout=TRIAL_SYNC_TIMEOUT)
        if st == 200:
            rec = json.loads(base64.b64decode(body["content"])
                             .decode("utf-8"))
            first_seen = int(rec.get("first_seen") or 0)
            if first_seen > 0:
                adopt_trial_start(first_seen)
        elif st == 404:
            rec = {"product": PRODUCT_ID, "mid": machine_binding_id(),
                   "first_seen": local_first,
                   "updated_at": int(time.time())}
            try:
                _ledger_put_path(path, rec)
            except Exception:
                pass
    except Exception:
        pass


# ===========================================================================
# Activation record (machine-encrypted, hidden, dual meaningful fields)
# ===========================================================================
def _activation_path():
    return os.path.join(_base_dir(), "act_%s.dat" % PRODUCT_ID)


def _backup_dir():
    """Machine-wide fallback store (survives user-profile file wipes)."""
    if os.name == "nt":
        base = (os.environ.get("PROGRAMDATA")
                or os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"))
        d = os.path.join(base, "VidaPay", "License")
    else:
        d = os.path.join(os.path.expanduser("~"), ".vidapay", "license")
    try:
        os.makedirs(d, exist_ok=True)
    except Exception:
        pass
    return d


def _backup_path():
    return os.path.join(_backup_dir(), "act_%s.dat" % PRODUCT_ID)


def _reg_activation_name():
    return "act_%s" % PRODUCT_ID



def _activation_store_blobs():
    """Candidate raw activation blobs, best store first."""
    blobs = []
    for path in (_activation_path(), _backup_path()):
        try:
            if os.path.exists(path):
                blobs.append(open(path, "rb").read())
        except Exception:
            pass
    if _winreg is not None:
        try:
            k = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER,
                                r"Software\VidaPay\License")
            try:
                data, _t = _winreg.QueryValueEx(k, _reg_activation_name())
                blobs.append(bytes(data))
            finally:
                _winreg.CloseKey(k)
        except Exception:
            pass
    return blobs


def _load_activation_raw():
    """Decrypt the record without full re-validation (for diagnostics).
    Tries the user file, the ProgramData backup, then the registry."""
    for blob in _activation_store_blobs():
        try:
            dec = _decrypt(blob)
            if not dec:
                continue
            return json.loads(dec.decode("utf-8"))
        except Exception:
            continue
    return None


def load_activation():
    """Return a validated activation record dict or None.

    EVERY store is tried (user file, ProgramData backup, HKCU registry)
    and the first record that fully validates wins. A stale or invalid
    record in one store no longer shadows a valid copy that lives in
    another store (field deadlock: the user file held an expired record
    and could not be overwritten - Permission denied - while the fresh
    activation sat in the registry/ProgramData backup, never read)."""
    bad = []
    winner = None
    for idx, blob in enumerate(_activation_store_blobs()):
        try:
            dec = _decrypt(blob)
            if not dec:
                continue
            rec = json.loads(dec.decode("utf-8"))
        except Exception:
            continue
        lic = rec.get("license") or {}
        ok, reason = validate_license(lic)
        if not ok:
            bad.append(str(reason))
            continue
        if str(rec.get("machine_id")) != machine_binding_id():
            bad.append("machine mismatch")
            continue
        receipt = str(rec.get("receipt") or "")
        if receipt:
            if not receipt_valid(receipt, rec.get("license_id"),
                                 rec.get("machine_id"),
                                 int(rec.get("activated_at") or 0)):
                bad.append("receipt invalid")
                continue
        winner = (idx, rec)
        break
    if winner is None:
        if bad:
            uniq = []
            for r in bad:
                if r not in uniq:
                    uniq.append(r)
            _lic_log("load_activation: no valid record in any store (%s)"
                     % "; ".join(uniq))
        return None
    idx, rec = winner
    if idx > 0:
        # The valid copy came from the backup/registry while an earlier
        # store still holds a stale record - rewrite every store (best
        # effort) so all copies converge on the valid record.
        _lic_log("load_activation: valid record from store #%d - "
                 "repairing earlier stores" % idx)
        save_activation(rec)
    return rec


def save_activation(rec):
    """Persist the activation record to EVERY store available (user file,
    ProgramData backup, HKCU registry). One surviving copy keeps this PC
    activated even if an antivirus or cleaner wipes the others."""
    try:
        ok, reason = validate_license(rec.get("license") or {})
    except Exception as _e:
        ok, reason = False, "validate crashed (%s)" % _e
    if not ok:
        # Never persist a record that would fail the next load_activation
        # (e.g. an expired license) - it would shadow any valid copy.
        _lic_log("save_activation: REFUSED invalid record (%s)" % reason)
        return False
    try:
        data = _encrypt(json.dumps(rec).encode("utf-8"))
    except Exception:
        _lic_log("save_activation: encrypt FAILED")
        return False
    saved = []
    try:
        p = _activation_path()
        try:
            os.chmod(p, 0o600)  # S_IWRITE|S_IREAD: clear read-only attr, stay readable
        except Exception:
            pass
        with open(p, "wb") as f:
            f.write(data)
        _hide_file(p)
        saved.append("user")
    except Exception as _e:
        _lic_log("save_activation: user file FAILED (%s)" % _e)
    try:
        with open(_backup_path(), "wb") as f:
            f.write(data)
        saved.append("backup")
    except Exception as _e:
        _lic_log("save_activation: backup file FAILED (%s)" % _e)
    if _winreg is not None:
        try:
            k = _winreg.CreateKey(_winreg.HKEY_CURRENT_USER,
                                  r"Software\VidaPay\License")
            _winreg.SetValueEx(k, _reg_activation_name(), 0,
                               _winreg.REG_BINARY, data)
            _winreg.CloseKey(k)
            saved.append("registry")
        except Exception as _e:
            _lic_log("save_activation: registry FAILED (%s)" % _e)
    _lic_log("save_activation: %s"
             % ("+".join(saved) if saved else "ALL FAILED"))
    return bool(saved)


def clear_activation():
    removed = []
    try:
        p = _activation_path()
        if os.path.exists(p):
            os.remove(p)
            removed.append("user")
    except Exception:
        pass
    try:
        bp = _backup_path()
        if os.path.exists(bp):
            os.remove(bp)
            removed.append("backup")
    except Exception:
        pass
    if _winreg is not None:
        try:
            k = _winreg.OpenKey(_winreg.HKEY_CURRENT_USER,
                                r"Software\VidaPay\License", 0,
                                _winreg.KEY_SET_VALUE)
            _winreg.DeleteValue(k, _reg_activation_name())
            _winreg.CloseKey(k)
            removed.append("registry")
        except Exception:
            pass
    _lic_log("clear_activation: %s"
             % ("+".join(removed) if removed else "nothing to clear"))


def activated_license_id():
    rec = load_activation()
    return rec.get("license_id") if rec else None


# ===========================================================================
# License ledger client (GitHub-hosted: one-PC binding + revocation)
# ===========================================================================
def _xor_cyc(data, key):
    reps = len(data) // len(key) + 1
    return bytes(a ^ b for a, b in zip(data, key * reps))


def _ledger_token():
    """Ledger access token. Embedded at build time via the generated
    license_token module (make_license_token.py). The env var is a
    developer/testing fallback only."""
    try:
        import license_token
        obf = str(getattr(license_token, "TOKEN_OBF", "") or "")
        if obf:
            return _xor_cyc(base64.b64decode(obf), _TOKEN_XK).decode("utf-8")
    except Exception:
        pass
    return os.environ.get("VIDAPAY_LICENSE_TOKEN", "")


def _ledger_ready():
    tok = _ledger_token()
    return bool(tok) and "PASTE" not in tok


class _LedgerDown(Exception):
    pass


def _gh(method, path, payload=None, timeout=ACTIVATION_TIMEOUT):
    url = "https://api.github.com" + path
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={
        "Authorization": "Bearer " + _ledger_token(),
        "Accept": "application/vnd.github+json",
        "User-Agent": "VidaPay-Licensing/2.0"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = resp.read().decode("utf-8", "replace")
            return resp.status, (json.loads(body) if body.strip() else {})
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return 404, {}
        if e.code in (409, 422):
            return e.code, {}
        raise _LedgerDown("http %d" % e.code)
    except _LedgerDown:
        raise
    except Exception:
        raise _LedgerDown("network")


def _ledger_rec_path(license_id):
    return ("/repos/%s/contents/ledger/activations/%s.json"
            % (LEDGER_REPO, license_id))


def _ledger_get(license_id, timeout=ACTIVATION_TIMEOUT):
    """-> (record, sha) or (None, None) when the record does not exist."""
    st, body = _gh("GET", _ledger_rec_path(license_id), timeout=timeout)
    if st == 404:
        return None, None
    if st != 200:
        raise _LedgerDown("read %d" % st)
    rec = json.loads(base64.b64decode(body["content"]).decode("utf-8"))
    return rec, body.get("sha")


def _ledger_put(license_id, rec, sha, timeout=ACTIVATION_TIMEOUT):
    """Compare-and-swap write. Returns (new_sha, conflict)."""
    body = json.dumps(rec, sort_keys=True,
                      separators=(",", ":")).encode("utf-8")
    payload = {"message": "license %s" % license_id,
               "content": base64.b64encode(body).decode("ascii")}
    if sha:
        payload["sha"] = sha
    st, resp = _gh("PUT", _ledger_rec_path(license_id), payload,
                   timeout=timeout)
    if st in (200, 201):
        return (resp.get("commit", {}) or {}).get("sha", ""), False
    return None, True


def receipt_valid(receipt_b64, license_id, machine_id, activated_at):
    """Binding-evidence check (the ledger itself is the real authority;
    the heartbeat re-checks it every HEARTBEAT_DAYS)."""
    try:
        obj = json.loads(base64.b64decode(receipt_b64))
        return (str(obj.get("mid")) == str(machine_id)
                and int(obj.get("at") or -1) == int(activated_at)
                and str(obj.get("lid")) == str(license_id))
    except Exception:
        return False


def _max_seats(rec):
    """How many PCs this license may be active on (default 1)."""
    try:
        n = int(rec.get("max_seats") or 1)
    except Exception:
        n = 1
    return max(1, n)


# ---------------------------------------------------------------------------
# Anti key-sharing move cooldown
# After a seat is released, the key may self-service move to ONE new PC
# inside TRANSFER_COOLDOWN_DAYS. Any further move inside the window is
# refused and falls back to the seller (email) - License Studio's seller-
# approved release always works and refreshes the one-move grace.
# ---------------------------------------------------------------------------
TRANSFER_COOLDOWN_DAYS = 14


def _cooldown_ok(rec, mid):
    """True when PC `mid` may take a NEW seat of this license.
    Single-seat licenses only: one free move per cooldown window; seller
    releases refresh it. Re-binding an existing seat is never a move."""
    if _max_seats(rec) > 1:
        return True  # per-PC volume licenses - no move gate
    if int(rec.get("cooldown_until") or 0) <= time.time():
        return True  # no active window
    if str(mid) in _seat_map(rec):
        return True  # same PC re-activating is not a move
    if str(mid) == str(rec.get("mid_code") or ""):
        return True  # re-binding the pinned / last-bound PC
    return int(rec.get("cooldown_moves_left") or 0) > 0


def _seat_map(rec):
    """fingerprint -> evidence. Migrates the old single mid_code field."""
    seats = rec.get("seats")
    if isinstance(seats, dict):
        return {str(k): str(v) for k, v in seats.items()}
    mid = str(rec.get("mid_code") or "")
    return {mid: str(rec.get("host") or "")} if mid else {}


def server_activate(lic):
    """Bind this license to THIS PC in the ledger. First writer wins, so
    a license can never be activated on two different PCs. A license that
    was machine-pinned at issue time may activate fully offline."""
    pinned_here = (str(lic.get("machine") or "").strip()
                   == machine_binding_id())
    if not _ledger_ready():
        if pinned_here:
            return True, "Activated (offline).", ""
        return False, ("Activation service is not configured. Internet is "
                       "required for the first activation."), ""
    lid = str(lic.get("license_id") or "").strip()
    if not lid:
        return False, "License file has no license id.", ""
    my_mid = machine_binding_id()
    for _attempt in range(5):
        try:
            rec, sha = _ledger_get(lid)
        except _LedgerDown:
            if pinned_here:
                return True, "Activated (offline).", ""
            return False, ("Could not reach the activation service "
                           "(internet required for first activation)."), ""
        if rec is None:
            return False, ("License not found in the registry - contact "
                           "%s with your key." % _BRAND), ""
        if str(rec.get("status")) == "revoked":
            return False, ("This license was revoked by the seller. "
                           "Contact %s." % _BRAND), ""
        seats = _seat_map(rec)
        max_seats = _max_seats(rec)
        if my_mid in seats:
            at = int(rec.get("activated_at") or time.time())
            if "seats" not in rec:
                try:  # one-time legacy migration write
                    rec2 = dict(rec)
                    rec2["seats"] = seats
                    _ledger_put(lid, rec2, sha)
                except Exception:
                    pass
        elif len(seats) >= max_seats:
            return False, ("This license is already active on %d PC(s). "
                           "Contact %s to move or upgrade it."
                           % (max_seats, _BRAND)), ""
        elif not _cooldown_ok(rec, my_mid):
            unlock = time.strftime("%b %d, %Y",
                                   time.gmtime(int(rec.get("cooldown_until") or 0)))
            return False, ("This license was moved to another PC within the "
                           "last %d days, so it cannot move again by itself "
                           "yet. Email %s with your order reference for a "
                           "manual move - self-service moves unlock again "
                           "on %s." % (TRANSFER_COOLDOWN_DAYS,
                                       _SUPPORT_EMAIL, unlock)), ""
        else:
            rec = dict(rec)
            rec["status"] = "active"
            seats[my_mid] = machine_evidence()
            rec["seats"] = seats
            # consume the one free self-service move inside a cooldown window
            if (int(rec.get("cooldown_until") or 0) > time.time()
                    and int(rec.get("cooldown_moves_left") or 0) > 0):
                rec["cooldown_moves_left"] = \
                    int(rec["cooldown_moves_left"]) - 1
            if "machine" not in lic:
                # Unpinned key: keep the legacy last-seat pointer.
                # A machine-pinned key must keep its ISSUE pin -
                # that field is required to rebuild the signed
                # payload byte-for-byte; overwriting it permanently
                # breaks re-activation ("failed verification").
                rec["mid_code"] = sorted(seats)[0]
            if not rec.get("activated_at"):
                rec["activated_at"] = int(time.time())
            rec["lic_sig"] = str(lic.get("sig") or rec.get("lic_sig") or "")
            _new_sha, conflict = _ledger_put(lid, rec, sha)
            if conflict:
                time.sleep(0.5)
                continue
            at = int(rec["activated_at"])
        ev = base64.b64encode(json.dumps(
            {"mid": my_mid, "at": at, "lid": lid},
            sort_keys=True, separators=(",", ":")).encode("utf-8")
        ).decode("ascii")
        return True, "Activated.", ev
    return False, "Activation service is busy, please try again.", ""


# ---------------------------------------------------------------------------
# Self-service deactivation ("I am moving to a new PC") + seller transfer log
# ---------------------------------------------------------------------------
TRANSFER_LOG_PATH = "/repos/%s/contents/ledger/transfer_log.json" % LEDGER_REPO


def _log_transfer(entry, timeout=ACTIVATION_TIMEOUT):
    """Best-effort append to the shared seller transfer log
    (ledger/transfer_log.json - a JSON array, newest first, capped at 500).
    Never raises: the seat is already released by the time this runs, so a
    failed log write must not fail the deactivation itself."""
    try:
        st, body = _gh("GET", TRANSFER_LOG_PATH, timeout=timeout)
        log, sha = [], None
        if st == 200:
            log = json.loads(base64.b64decode(body["content"]).decode("utf-8"))
            sha = body.get("sha")
        elif st != 404:
            return
        if not isinstance(log, list):
            log = []
        log.insert(0, dict(entry))
        payload = {"message": "transfer log %s" % str(entry.get("key") or "?"),
                   "content": base64.b64encode(
                       json.dumps(log[:500], indent=1).encode("utf-8")
                   ).decode("ascii")}
        if sha:
            payload["sha"] = sha
        _gh("PUT", TRANSFER_LOG_PATH, payload, timeout=timeout)
    except Exception:
        pass


def deactivate_this_pc(timeout=ACTIVATION_TIMEOUT):
    """Release THIS PC's seat in the ledger and forget the local activation.
    The license key itself is NOT consumed or changed - the customer can
    activate the same key on any PC (normal seat limits still apply).
    Returns (ok, message)."""
    local = load_activation()
    lid = str((local or {}).get("license_id") or "").strip()
    if not lid:
        return False, ("This PC is not activated right now, so there is "
                       "nothing to deactivate.")
    if not _ledger_ready():
        return False, ("An internet connection is required to deactivate "
                       "- the seat must be released server-side.")
    my_mid = machine_binding_id()
    released = False
    for _attempt in range(5):
        try:
            rec, sha = _ledger_get(lid)
        except _LedgerDown:
            return False, ("Could not reach the activation service. "
                           "Check the internet connection and try again.")
        if rec is None:
            released = True  # server-side record gone: local state is stale
            break
        seats = _seat_map(rec)
        if my_mid not in seats:
            released = True  # seat already free (seller moved it earlier)
            break
        rec = dict(rec)
        seats = dict(seats)
        seats.pop(my_mid, None)
        rec["seats"] = seats
        lic_local = (local or {}).get("license") or {}
        if ("machine" not in lic_local
                and str(rec.get("mid_code") or "") == my_mid):
            # Pinned keys must never lose mid_code (issue pin).
            rec["mid_code"] = sorted(seats)[0] if seats else ""
        if (_max_seats(rec) <= 1
                and int(rec.get("cooldown_until") or 0) <= time.time()):
            # arm the anti-sharing window (only when none is active - a
            # deactivate inside an active window must NOT refresh it, or
            # key sharers could chain A->B->C every few days):
            # the next self-service move is free once; the move after that
            # needs the seller inside 14 days
            rec["cooldown_until"] = (int(time.time())
                                     + TRANSFER_COOLDOWN_DAYS * 86400)
            rec["cooldown_moves_left"] = 1
        _new_sha, conflict = _ledger_put(lid, rec, sha)
        if not conflict:
            released = True
            _log_transfer({"at": int(time.time()), "key": lid,
                           "action": "app_deactivate", "app": APP_NAME,
                           "pc": machine_evidence()})
            break
        time.sleep(0.6)
    if not released:
        return False, ("The activation service is busy - please try again "
                       "in a few seconds.")
    clear_activation()
    return True, ("This PC has been deactivated. Your license key is NOT "
                  "consumed - run the app on the new PC and activate with "
                  "the same key.")


def _lic_from_ledger_record(rec, include_machine):
    """Rebuild the keygen-signed license payload from a ledger record.
    The payload must match the signed JSON byte-for-byte, so the machine
    pin is only included when requested (issue-time pins exist only when
    the seller machine-pinned the key at issue time)."""
    lic = {
        "license_id": str(rec.get("license_id") or "").strip(),
        "customer": str(rec.get("customer") or ""),
        "products": list(rec.get("products") or []),
        "model": str(rec.get("model") or "lifetime"),
        "max_seats": max(1, int(rec.get("max_seats") or 1)),
        "issued_at": str(rec.get("issued_at") or ""),
    }
    if include_machine and rec.get("mid_code"):
        lic["machine"] = str(rec.get("mid_code"))
    if lic["model"] == "expiry":
        lic["expires"] = str(rec.get("expires") or "")
    return lic


def activate_with_license_key(key_text):
    """Activate by typing the license key (VP3S-XXXX-XXXX-XXXX-XXXX) -
    no .lic file needed. The keygen-signed record is verified against
    the embedded public key, then bound to THIS PC (one-PC rule).
    Returns (ok, message, lic, receipt)."""
    key = "".join(str(key_text or "").split()).upper()
    if not key:
        return False, "Type or paste your license key first.", None, ""
    if not _ledger_ready():
        return False, ("Activation service is not configured. Internet "
                       "is required for the first activation."), None, ""
    try:
        rec, _sha = _ledger_get(key)
    except _LedgerDown:
        return False, ("Could not reach the activation service (internet "
                       "required for the first activation)."), None, ""
    if rec is None:
        return False, ("License not found in the registry - check the key "
                       "or contact %s." % _BRAND), None, ""
    if str(rec.get("status")) == "revoked":
        return False, ("This license was revoked by the seller. "
                       "Contact %s." % _BRAND), None, ""
    sig = str(rec.get("lic_sig") or "")
    lic_ok = None
    for include_machine in (False, True):
        cand = _lic_from_ledger_record(rec, include_machine)
        cand["sig"] = sig
        ok, _reason = validate_license(cand)
        if ok:
            lic_ok = cand
            break
    if lic_ok is None:
        return False, ("This license record failed verification (unknown "
                       "or tampered). Contact %s." % _BRAND), None, ""
    ok, msg, receipt = server_activate(lic_ok)
    if not ok:
        return False, msg, None, ""
    return True, msg, lic_ok, receipt


def server_verify(license_id, lic=None, timeout=5):
    """Silent heartbeat. Returns 'ok' | 'revoked' | 'moved' | 'unreachable'."""
    if not _ledger_ready():
        return "unreachable"
    try:
        rec, _sha = _ledger_get(str(license_id), timeout=timeout)
    except _LedgerDown:
        return "unreachable"
    if rec is None:
        # The ledger ANSWERED but the record is gone: the seller removed
        # or revoked the key. This must fail CLOSED (unlike a network
        # failure, which raises _LedgerDown and stays fail-open above).
        return "removed"
    if str(rec.get("status")) == "revoked":
        return "revoked"
    seats = _seat_map(rec)
    if not seats or machine_binding_id() not in seats:
        # Active record but this PC holds NO seat: the seller released
        # the seat (empty seat map) or it moved to another PC. Treating
        # this as 'ok' would let a released license run forever.
        return "moved"
    return "ok"


def server_trial_sync(timeout=4):
    """Best-effort: keep the earliest trial start per machine+product in
    the ledger so reinstalling the app does not reset the 7-day clock."""
    if not _ledger_ready():
        return
    mid = machine_binding_id()
    slug = hashlib.sha256(mid.encode("utf-8")).hexdigest()[:16]
    path = ("/repos/%s/contents/ledger/trials/%s_%s.json"
            % (LEDGER_REPO, slug, PRODUCT_ID))
    legacy_path = ("/repos/%s/contents/ledger/trials/%s-%s.json"
                   % (LEDGER_REPO, PRODUCT_ID, mid))
    now = int(time.time())
    for _attempt in range(3):
        try:
            st, body = _gh("GET", path, timeout=timeout)
        except _LedgerDown:
            return
        rec, sha = None, None
        if st == 200:
            try:
                rec = json.loads(base64.b64decode(body["content"])
                                 .decode("utf-8"))
                sha = body.get("sha")
            except Exception:
                return
        elif st != 404:
            return
        if st == 404:
            # Older shipped builds stored trials as
            # '<product>-<mid>.json'. Adopt their first_seen so a
            # reinstall/upgrade keeps the same trial clock.
            try:
                st_l, body_l = _gh("GET", legacy_path,
                                   timeout=timeout)
            except _LedgerDown:
                return
            if st_l == 200:
                try:
                    legacy = json.loads(base64.b64decode(
                        body_l["content"]).decode("utf-8"))
                    fs_l = int(legacy.get("first_seen") or 0)
                except Exception:
                    return
                if fs_l:
                    adopt_trial_start(fs_l)
                return  # legacy record exists - do not duplicate
        if rec:
            fs = int(rec.get("first_seen") or 0)
            if fs:
                adopt_trial_start(fs)
            return
        rec = {"machine": mid, "product": PRODUCT_ID, "first_seen": now}
        body_bytes = json.dumps(rec, sort_keys=True,
                                separators=(",", ":")).encode("utf-8")
        payload = {"message": "trial %s" % slug,
                   "content": base64.b64encode(body_bytes).decode("ascii")}
        if sha:
            payload["sha"] = sha
        try:
            st, _resp = _gh("PUT", path, payload, timeout=timeout)
        except _LedgerDown:
            return
        if st in (200, 201):
            return
        time.sleep(0.5)


# ===========================================================================
# Anti-debug / anti-VM (active in frozen exe builds)
# ===========================================================================
def debugger_present():
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
        if out.value:
            return True
    except Exception:
        pass
    return False


_VM_MARKERS = ("vmware", "virtualbox", "vbox", "qemu", "kvm", "xen",
               "virtual machine", "bochs")
_VM_MAC_PREFIXES = ("08:00:27", "00:0C:29", "00:05:69", "00:1C:14",
                    "52:54:00")


def vm_signal():
    """Return (detected, reason). Windows only."""
    if os.name != "nt":
        return False, ""
    try:
        import winreg
        for hive, path in ((winreg.HKEY_LOCAL_MACHINE,
                            r"HARDWARE\DESCRIPTION\System\BIOS"),):
            try:
                k = winreg.OpenKey(hive, path)
                for val in ("SystemManufacturer", "SystemProductName",
                            "BIOSVendor"):
                    try:
                        text, _ = winreg.QueryValueEx(k, val)
                        low = str(text).lower()
                        for m in _VM_MARKERS:
                            if m in low:
                                return True, "virtual machine (%s)" % val
                    except Exception:
                        continue
                winreg.CloseKey(k)
            except Exception:
                pass
    except Exception:
        pass
    try:
        hits = [p for p in (
            r"C:\Program Files\VMware\VMware Tools",
            r"C:\Program Files\Oracle\VirtualBox Guest Additions",
            r"C:\Program Files\qemu-ga",
        ) if os.path.exists(p)]
        if hits:
            return True, "virtual machine guest tools"
    except Exception:
        pass
    try:
        import uuid
        mac = uuid.getnode()
        mac_str = "%012x" % mac
        formatted = ":".join(mac_str[i:i + 2] for i in range(0, 12, 2)).upper()
        for pre in _VM_MAC_PREFIXES:
            if formatted.startswith(pre):
                return True, "virtual machine network adapter"
    except Exception:
        pass
    return False, ""


# ===========================================================================
# Activation window (standalone dark dialog, no ttk style interference)
# ===========================================================================
# 3sverse.com dark-hero brand tokens (match License Studio v1.1+)
_BG = "#07060b"          # site canvas hsl(250 28% 3%)
_PANEL = "#0d0c14"       # site card hsl(250 20% 6%)
_BORDER = "#232130"
_TEXT = "#f5f1e6"        # warm white hsl(49 39% 94%)
_DIM = "#9b97b2"
_CYAN = "#6ee7ef"        # brand cyan - primary actions
_ERR = "#ef6a52"         # error text (site destructive token)
_ACCENT = _CYAN          # primary button fill + accent subtitle
_OK = "#c7ef70"          # brand lime - success
_WARN = "#f5c518"


def _msgbox(kind, title, text):
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk()
        r.withdraw()
        if kind == "error":
            messagebox.showerror(title, text, parent=r)
        else:
            messagebox.showinfo(title, text, parent=r)
        r.destroy()
    except Exception:
        pass


def _show_activation_window(notice=None, trial_label=None):
    """Modal activation window. Returns True when activation succeeds or
    the user continues an active trial. trial_label is the human-readable
    remaining-time string shown on the Continue Trial button (never shown
    for an expired trial)."""
    import tkinter as tk
    result = {"ok": False}
    # Token-less (public trial) builds cannot reach the ledger, so the
    # key/.lic sections are replaced by a 'buy a license' panel.
    ledger_on = _ledger_ready()

    root = tk.Tk()
    root.title("%s - Activation" % APP_NAME)
    root.configure(bg=_BG)
    root.resizable(False, False)
    try:
        root.iconbitmap(default=_icon_path())
    except Exception:
        pass

    w, h = 500, 460
    try:
        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        root.geometry("%dx%d+%d+%d" % (w, h, (sw - w) // 2, (sh - h) // 2))
    except Exception:
        root.geometry("%dx%d" % (w, h))

    def _fit_window():
        """Grow the window to fit its content so nothing is cut off."""
        nonlocal w, h
        try:
            root.update_idletasks()
            sw = root.winfo_screenwidth()
            sh = root.winfo_screenheight()
            need_w = max(w, root.winfo_reqwidth() + 28)
            need_h = max(h, root.winfo_reqheight() + 28)
            w = max(w, min(need_w, sw - 40))
            h = max(h, min(need_h, sh - 60))
            root.geometry("%dx%d+%d+%d" % (w, h, max(0, (sw - w) // 2),
                                           max(0, (sh - h) // 2)))
        except Exception:
            pass

    def panel(**kw):
        f = tk.Frame(root, bg=_PANEL, highlightbackground=_BORDER,
                     highlightthickness=1, **kw)
        return f

    pad = {"padx": 18}

    head = tk.Frame(root, bg=_BG)
    head.pack(fill="x", padx=0, pady=(18, 6))
    tk.Label(head, text=APP_NAME, bg=_BG, fg=_TEXT,
             font=("Segoe UI", 15, "bold")).pack()
    tk.Label(head,
             text="Activation Required" if ledger_on else "Free Trial",
             bg=_BG, fg=_ACCENT if ledger_on else _WARN,
             font=("Segoe UI", 10, "bold")).pack(pady=(2, 0))

    if notice:
        nb = panel()
        nb.pack(fill="x", in_=root, padx=18, pady=(8, 0))
        tk.Label(nb, text=notice, bg=_PANEL, fg=_WARN, justify="left",
                 wraplength=440, font=("Segoe UI", 9)).pack(
            anchor="w", padx=10, pady=8)

    # machine id
    mid_box = panel()
    mid_box.pack(fill="x", **pad, pady=(12, 0))
    tk.Label(mid_box, text="Your Machine ID (this copy is locked to it):",
             bg=_PANEL, fg=_DIM, font=("Segoe UI", 9)).pack(
        anchor="w", padx=10, pady=(8, 2))
    row = tk.Frame(mid_box, bg=_PANEL)
    row.pack(fill="x", padx=10, pady=(0, 10))
    mid_var = tk.StringVar(value=machine_binding_id())
    mid_entry = tk.Entry(row, textvariable=mid_var, bg=_BG, fg=_TEXT,
                         relief="flat", insertbackground=_TEXT,
                         font=("Consolas", 10), state="readonly",
                         readonlybackground=_BG)
    mid_entry.pack(side="left", fill="x", expand=True, ipady=5)

    def copy_mid():
        root.clipboard_clear()
        root.clipboard_append(mid_var.get())
        mid_status.config(text="Machine ID copied.", fg=_OK)
        root.after(2500, lambda: mid_status.config(text="", fg=_DIM))

    tk.Button(row, text="Copy", command=copy_mid, bg=_PANEL, fg=_TEXT,
              activebackground=_BORDER, activeforeground=_TEXT,
              relief="flat", bd=0, padx=12, pady=4,
              font=("Segoe UI", 9, "bold"),
              highlightbackground=_BORDER).pack(side="left", padx=(8, 0))
    mid_status = tk.Label(mid_box, text="", bg=_PANEL, fg=_DIM,
                          font=("Segoe UI", 8))
    mid_status.pack(anchor="w", padx=10, pady=(0, 6))

    # license file (paid builds only - needs the ledger token)
    lic_box = panel()
    if ledger_on:
        lic_box.pack(fill="x", **pad, pady=(12, 0))
    else:
        buy_box = panel()
        buy_box.pack(fill="x", **pad, pady=(12, 0))
        tk.Label(buy_box, text="FREE TRIAL BUILD", bg=_PANEL, fg=_WARN,
                 font=("Segoe UI", 10, "bold")).pack(
            anchor="w", padx=10, pady=(8, 2))
        tk.Label(buy_box,
                 text="When the trial ends, buy a license at "
                      "www.3SVerse.com - you will receive a license key "
                      "and the full build, which unlocks right here.",
                 bg=_PANEL, fg=_TEXT, justify="left", wraplength=440,
                 font=("Segoe UI", 9)).pack(anchor="w", padx=10,
                                            pady=(0, 8))
    tk.Label(lic_box,
             text="Select the license file (.lic) you received "
                  "from the seller:",
             bg=_PANEL, fg=_DIM, font=("Segoe UI", 9)).pack(
        anchor="w", padx=10, pady=(8, 4))
    status = tk.Label(lic_box, text="", bg=_PANEL, fg=_DIM,
                      wraplength=440, justify="left",
                      font=("Segoe UI", 9))

    def set_status(text, color=_DIM):
        status.config(text=text, fg=color)
        lic_box.update_idletasks()
        _fit_window()

    def save_and_finish(lic, receipt):
        try:
            _ev = json.loads(base64.b64decode(receipt))
            _at = int(_ev.get("at") or time.time())
        except Exception:
            _at = int(time.time())
        rec = {"license": lic,
               "license_id": lic.get("license_id"),
               "machine_id": machine_binding_id(),
               "products": lic.get("products"),
               "model": lic.get("model"),
               "expires": lic.get("expires"),
               "activated_at": _at,
               "last_verify": int(time.time()),
               "receipt": receipt}
        saved = save_activation(rec)
        if saved:
            set_status("Activated successfully. Starting %s..." % APP_NAME,
                       _OK)
            result["ok"] = True
            root.after(1200, root.destroy)
        else:
            set_status("Activation verified, but saving the license on "
                       "this PC failed (antivirus?). Contact %s."
                       % _SUPPORT_EMAIL, _ERR)
            _lic_log("save_and_finish: ALL stores FAILED - user saw error")

    def _continue_trial():
        result["ok"] = True
        root.destroy()

    # license key entry - no .lic file required (paid builds only)
    key_box = panel()
    if ledger_on:
        key_box.pack(fill="x", **pad, pady=(12, 0))
    tk.Label(key_box, text="Or type your license key "
                           "(VP3S-XXXX-XXXX-XXXX-XXXX):",
             bg=_PANEL, fg=_DIM, font=("Segoe UI", 9)).pack(
        anchor="w", padx=10, pady=(8, 4))
    key_entry = tk.Entry(key_box, bg=_BG, fg=_TEXT, relief="flat",
                         insertbackground=_TEXT, font=("Consolas", 10),
                         highlightbackground=_BORDER, highlightthickness=1)
    key_entry.pack(fill="x", padx=10, pady=(0, 6), ipady=5)

    def do_activate_key():
        set_status("Verifying license key...")
        root.update_idletasks()
        ok, msg, lic, receipt = activate_with_license_key(key_entry.get())
        if not ok:
            set_status(msg, _ERR)
            return
        save_and_finish(lic, receipt)

    tk.Button(key_box, text="Activate with License Key",
              command=do_activate_key,
              bg=_ACCENT, fg="#07060b", activebackground="#9deff5",
              activeforeground="#07060b", relief="flat", bd=0, padx=14,
              pady=6, font=("Segoe UI", 9, "bold"),
              cursor="hand2").pack(anchor="w", padx=10, pady=(0, 10))

    def do_activate():
        from tkinter import filedialog
        path = filedialog.askopenfilename(
            title="Select license file",
            filetypes=[("License files", "*.lic"), ("All files", "*.*")])
        if not path:
            return
        set_status("Verifying license file...")
        root.update_idletasks()
        try:
            lic = load_license_file(path)
        except Exception:
            set_status("Could not read the license file.", _ERR)
            return
        ok, reason = validate_license(lic)
        if not ok:
            set_status(reason, _ERR)
            return
        set_status("License verified. Contacting activation server...")
        root.update_idletasks()
        ok, msg, receipt = server_activate(lic)
        if ok:
            save_and_finish(lic, receipt)
        else:
            set_status(msg, _ERR)

    btns = tk.Frame(lic_box, bg=_PANEL)
    btns.pack(fill="x", padx=10, pady=(2, 10))
    if trial_label is not None:
        tk.Button(btns, text="Continue Trial (%s)" % trial_label,
                  command=_continue_trial,
                  bg=_OK, fg="#07060b", activebackground="#d9f79f",
                  activeforeground="#07060b", relief="flat", bd=0,
                  padx=14, pady=6, font=("Segoe UI", 9, "bold"),
                  cursor="hand2").pack(side="left")
    tk.Button(btns, text="Select License File...", command=do_activate,
              bg=_ACCENT, fg="#07060b", activebackground="#9deff5",
              activeforeground="#07060b", relief="flat", bd=0, padx=14,
              pady=6, font=("Segoe UI", 9, "bold"),
              cursor="hand2").pack(side="left")
    status.pack(anchor="w", padx=10, pady=(0, 8))

    foot = tk.Frame(root, bg=_BG)
    foot.pack(fill="x", side="bottom", pady=(10, 12))

    def on_exit():
        root.destroy()

    tk.Button(foot, text="Exit", command=on_exit, bg=_BG, fg=_DIM,
              activebackground=_BORDER, activeforeground=_TEXT,
              relief="flat", bd=1, padx=14, pady=4,
              font=("Segoe UI", 9)).pack()
    tk.Label(foot, text=_BRAND, bg=_BG, fg=_DIM,
             font=("Segoe UI", 8)).pack(pady=(4, 0))
    tk.Label(foot, text="3S Verse is an independent vendor - not "
                        "affiliated with or endorsed by VidaPay.",
             bg=_BG, fg=_DIM, font=("Segoe UI", 7)).pack(pady=(1, 0))

    root.protocol("WM_DELETE_WINDOW", on_exit)
    _fit_window()
    root.attributes("-topmost", True)
    root.after(250, lambda: root.attributes("-topmost", False))
    root.mainloop()
    return result["ok"]


def _icon_path():
    try:
        base = os.path.dirname(os.path.abspath(__file__))
        cand = os.path.join(base, "vidapay_icon.ico")
        if os.path.exists(cand):
            return cand
        if getattr(sys, "frozen", False):
            cand = os.path.join(os.path.dirname(sys.executable),
                                "vidapay_icon.ico")
            if os.path.exists(cand):
                return cand
    except Exception:
        pass
    return None


# ===========================================================================
# Entry point: call BEFORE the app GUI is created
# ===========================================================================
RUNTIME_RECHECK_SECONDS = 3600   # seller release/revoke kicks in within 1 h


def _watchdog_start(license_id):
    """Hourly silent re-check while the app runs: a license the seller
    released or revoked closes the app instead of waiting for the next
    launch. Offline checks are fail-open (unreachable -> keep running)."""
    if getattr(_watchdog_start, "_running", False):
        return
    _watchdog_start._running = True
    if not _ledger_ready():
        return

    def _loop():
        while True:
            time.sleep(RUNTIME_RECHECK_SECONDS)
            rec = load_activation()
            if not rec:
                return
            state = server_verify(rec.get("license_id"), timeout=6)
            if state in ("revoked", "moved", "removed"):
                clear_activation()
                why = ("revoked by the seller" if state == "revoked"
                       else "removed or revoked by the seller"
                       if state == "removed"
                       else "released by the seller or activated on "
                            "another PC")
                _msgbox("warning", "%s - License stopped" % APP_NAME,
                        "This license was %s.\n\nThe application will "
                        "now close. Contact %s if you believe this is "
                        "a mistake." % (why, _BRAND))
                _exit_note("watchdog: license %s -> closing app" % state)
                os._exit(0)

    threading.Thread(target=_loop, daemon=True).start()


def ensure_licensed():
    """Return True when the app may run (activated or trial active)."""
    frozen = bool(getattr(sys, "frozen", False))

    if frozen and debugger_present():
        _exit_note("anti-tamper: debugger detected at launch")
        os._exit(0)  # silent anti-tamper

    if frozen:
        vm, _why = vm_signal()
        if vm:
            _msgbox("error", "%s - Not allowed" % APP_NAME,
                    "This application cannot run inside a virtual "
                    "machine.\n\n%s" % _BRAND)
            return False

    rec = load_activation()
    if rec:
        lid = rec.get("license_id")
        # Re-verify at EVERY launch (was: every HEARTBEAT_DAYS) - a
        # license the seller released or revoked must stop working at
        # the next start, not a week later. Offline stays fail-open.
        state = server_verify(lid, lic=rec.get("license"), timeout=4)
        if state in ("revoked", "moved", "removed"):
            clear_activation()
            why = ("revoked by the seller" if state == "revoked"
                   else "removed or revoked by the seller"
                   if state == "removed"
                   else "released by the seller or activated on "
                        "another PC")
            return _show_activation_window(
                "This license was %s. Contact %s "
                "if you believe this is a mistake." % (why, _BRAND))
        _watchdog_start(lid)
        rec["last_verify"] = int(time.time())
        save_activation(rec)
        return True

    # record existed but failed validation -> diagnose for the user
    notice = None
    raw = _load_activation_raw()
    if raw:
        lic = raw.get("license") or {}
        if lic.get("model") == "expiry":
            exp = str(lic.get("expires") or "")
            try:
                if _dt.date.today() > _dt.date.fromisoformat(exp):
                    notice = ("Your license EXPIRED on %s.\n"
                              "Renew it with the seller to keep using "
                              "%s." % (exp, APP_NAME))
            except Exception:
                pass
        if notice is None:
            notice = ("The saved license record is not valid anymore. "
                      "Please activate again.")

    if notice is None:
        # no valid activation -> trial (server-synced against reinstall reset)
        trial_server_sync()
        expired, days_left = trial_status()
        if not expired:
            label = _trial_label(days_left)
            # Always show the activation window while the trial is running
            # so every launch makes the trial status and activation explicit.
            if _ledger_ready():
                notice = ("You are on the %d-day free trial - %s.\n\n"
                          "Activate now with your license key or license "
                          "file (.lic) to unlock the full version, or "
                          "continue the trial."
                          % (TRIAL_DAYS, label))
            else:
                notice = ("You are on the %d-day free trial - %s.\n\n"
                          "This is the free trial build - continue the "
                          "trial, or buy a license at www.3SVerse.com to "
                          "unlock the full version."
                          % (TRIAL_DAYS, label))
            return _show_activation_window(notice, trial_label=label)
        if _ledger_ready():
            notice = ("Your %d-day free trial of %s has ended.\n"
                      "Activate with your license key or license file "
                      "(.lic) to keep using it." % (TRIAL_DAYS, APP_NAME))
        else:
            notice = ("Your %d-day free trial of %s has ended.\n"
                      "Buy a license at www.3SVerse.com to keep using it."
                      % (TRIAL_DAYS, APP_NAME))

    return _show_activation_window(notice)


def get_status_summary():
    """Human-readable one-liner, handy for logs/About boxes."""
    rec = load_activation()
    if rec:
        exp = rec.get("expires")
        if rec.get("model") == "expiry" and exp:
            return "Licensed to %s (valid till %s)" % (
                rec.get("license", {}).get("customer") or "-", exp)
        return "Licensed to %s (lifetime)" % (
            rec.get("license", {}).get("customer") or "-")
    expired, days_left = trial_status()
    if not expired:
        return "Free trial: %s" % _trial_label(days_left)
    return "Not activated"


if __name__ == "__main__":
    # standalone check window: python license_core.py
    print(get_status_summary())
    print("Machine ID:", machine_binding_id())
    ok = ensure_licensed()
    print("ensure_licensed ->", ok)
