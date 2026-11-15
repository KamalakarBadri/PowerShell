#!/usr/bin/env python3
"""
M365 Group Membership Report

Input:  CSV with user UPNs
Output: two CSVs
  1. <out>_by_group.csv  — one row per group: Group ID, Group Name, Type, User Count, Users
  2. <out>_by_user.csv   — one row per (user, group): User, Group ID, Group Name, Type

Notes:
  - Uses Microsoft Graph (client-credentials, app-only).
  - No owners. No member expansion beyond the given user list.
  - Sequential with per-user delay + 429 retry.
  - Ctrl+C stops cleanly; partial CSVs are saved.
  - If a user's memberOf is paginated, all pages are followed.

Usage:
  python m365_group_report.py --users users.csv
  python m365_group_report.py --users users.csv --out groups --user-delay 0.5
"""

import argparse
import base64
import csv
import json
import os
import signal
import sys
import time
import uuid
from datetime import datetime
from threading import Event, Lock
from typing import Dict, List, Optional, Tuple

import requests
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate

CONFIG = {
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "tenant_name": "geekbyteonline.onmicrosoft.com",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "client_secret": "CyG8Q~FYHuCMSyVmt4sNxt5IejrMc2c24Ziz4a.t",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    "scope": "https://graph.microsoft.com/.default",
}

STOP_EVENT = Event()
DELAY = {"user": 0.5, "retry": 10.0}


def log(msg: str) -> None:
    print(msg, flush=True)


# ---------------- AUTH ----------------
_token_cache: Dict[str, Tuple[str, float]] = {}
_token_lock = Lock()


def _tok_cert(scope: str) -> Optional[str]:
    try:
        if not os.path.exists(CONFIG["certificate_path"]) or not os.path.exists(CONFIG["private_key_path"]):
            return None
        with open(CONFIG["certificate_path"], "rb") as cf:
            cert = load_pem_x509_certificate(cf.read(), default_backend())
        with open(CONFIG["private_key_path"], "rb") as kf:
            key = load_pem_private_key(kf.read(), password=None, backend=default_backend())
        now = int(time.time())
        h = {"alg": "RS256", "typ": "JWT",
             "x5t": base64.urlsafe_b64encode(cert.fingerprint(hashes.SHA1())).decode().rstrip("=")}
        p = {"aud": f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
             "exp": now + 300, "iss": CONFIG["app_id"], "jti": str(uuid.uuid4()),
             "nbf": now, "sub": CONFIG["app_id"]}
        eh = base64.urlsafe_b64encode(json.dumps(h).encode()).decode().rstrip("=")
        ep = base64.urlsafe_b64encode(json.dumps(p).encode()).decode().rstrip("=")
        u = f"{eh}.{ep}"
        sig = key.sign(u.encode(), padding.PKCS1v15(), hashes.SHA256())
        es = base64.urlsafe_b64encode(sig).decode().rstrip("=")
        jwt = f"{u}.{es}"
        r = requests.post(
            f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            data={"client_id": CONFIG["app_id"], "client_assertion": jwt,
                  "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                  "scope": scope, "grant_type": "client_credentials"}, timeout=30)
        if r.status_code == 200:
            return r.json()["access_token"]
        log(f"  cert auth failed: {r.status_code} {r.text[:200]}")
    except Exception as e:
        log(f"  cert auth exception: {e}")
    return None


def _tok_secret(scope: str) -> Optional[str]:
    try:
        r = requests.post(
            f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            data={"client_id": CONFIG["app_id"], "client_secret": CONFIG["client_secret"],
                  "scope": scope, "grant_type": "client_credentials"}, timeout=30)
        if r.status_code == 200:
            return r.json()["access_token"]
        log(f"  secret auth failed: {r.status_code} {r.text[:200]}")
    except Exception as e:
        log(f"  secret auth exception: {e}")
    return None


def get_graph_token() -> Optional[str]:
    with _token_lock:
        c = _token_cache.get("graph")
        if c and c[1] > time.time() + 60:
            return c[0]
    tok = _tok_cert(CONFIG["scope"]) or _tok_secret(CONFIG["scope"])
    if tok:
        with _token_lock:
            _token_cache["graph"] = (tok, time.time() + 3300)
    return tok


# ---------------- HTTP with 429 retry ----------------
def http_get(url: str, token: str) -> requests.Response:
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    attempt = 0
    while True:
        if STOP_EVENT.is_set():
            raise KeyboardInterrupt("Stop requested")
        try:
            r = requests.get(url, headers=headers, timeout=30)
        except requests.RequestException:
            attempt += 1
            if attempt > 3:
                raise
            time.sleep(2 ** attempt)
            continue
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", DELAY["retry"]))
            log(f"    ⏸ 429 — waiting {wait}s")
            for _ in range(wait):
                if STOP_EVENT.is_set():
                    raise KeyboardInterrupt("Stop requested")
                time.sleep(1)
            attempt += 1
            if attempt > 5:
                return r
            continue
        return r


# ---------------- GRAPH: memberOf with paging ----------------
def get_member_of(upn: str, token: str) -> Tuple[bool, List[Dict], str]:
    """
    Returns (ok, list_of_group_dicts, error_detail).
    Follows @odata.nextLink until exhausted.
    Filters to groups only (#microsoft.graph.group).
    """
    groups: List[Dict] = []
    url = (f"https://graph.microsoft.com/v1.0/users/{upn}/memberOf"
           f"?$select=id,displayName,groupTypes,mailEnabled,securityEnabled&$top=999")
    try:
        while url:
            if STOP_EVENT.is_set():
                raise KeyboardInterrupt("Stop requested")
            r = http_get(url, token)
            if r.status_code == 404:
                return False, [], "user not found in tenant"
            if r.status_code != 200:
                return False, [], f"HTTP {r.status_code}: {r.text[:300]}"
            body = r.json()
            for g in body.get("value", []):
                if "#microsoft.graph.group" not in (g.get("@odata.type") or ""):
                    continue
                gtype = "M365" if "Unified" in (g.get("groupTypes") or []) else "Security"
                groups.append({
                    "id": g.get("id", ""),
                    "name": g.get("displayName", ""),
                    "type": gtype,
                })
            url = body.get("@odata.nextLink")
    except KeyboardInterrupt:
        raise
    except Exception as e:
        return False, groups, f"exception: {e}"
    return True, groups, ""


# ---------------- CSV ----------------
def read_users(path: str) -> List[str]:
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    out: List[str] = []
    seen = set()
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rd = csv.DictReader(f)
        if not rd.fieldnames:
            raise ValueError("users CSV has no header")
        cols = ["User UPN", "UPN", "upn", "User", "user",
                "Email", "email", "UserPrincipalName", "userPrincipalName"]
        col = next((c for c in cols if c in rd.fieldnames), rd.fieldnames[0])
        for row in rd:
            v = (row.get(col) or "").strip()
            if not v:
                continue
            if "@" not in v:
                log(f"⚠  skipping malformed UPN: {v!r}")
                continue
            key = v.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(v)
    return out


def write_by_group(path: str, groups: Dict[str, Dict]) -> None:
    fields = ["Group ID", "Group Name", "Type", "User Count", "Users"]
    rows = sorted(
        (
            {
                "Group ID": g["id"],
                "Group Name": g["name"],
                "Type": g["type"],
                "User Count": len(g["users"]),
                "Users": "; ".join(sorted(g["users"])),
            }
            for g in groups.values()
        ),
        key=lambda r: (-r["User Count"], r["Group Name"].lower()),
    )
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow(r)


def write_by_user(path: str, pairs: List[Dict]) -> None:
    fields = ["User", "Group ID", "Group Name", "Type"]
    rows = sorted(pairs, key=lambda r: (r["User"].lower(), r["Group Name"].lower()))
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow(r)


# ---------------- MAIN ----------------
def main():
    ap = argparse.ArgumentParser(description="M365 group membership report for a list of users.")
    ap.add_argument("--users", required=True, help="CSV with user UPNs")
    ap.add_argument("--out", default=None, help="Output prefix (default: m365_groups_<timestamp>)")
    ap.add_argument("--user-delay", type=float, default=0.5, help="Seconds between users (default 0.5)")
    ap.add_argument("--retry-wait", type=float, default=10.0, help="Fallback wait on 429")
    args = ap.parse_args()

    DELAY["user"] = args.user_delay
    DELAY["retry"] = args.retry_wait

    def _sigint(signum, frame):
        if not STOP_EVENT.is_set():
            log("\n\n⛔ Ctrl+C — stopping at next checkpoint (partial CSVs saved)…")
            STOP_EVENT.set()
        else:
            log("⛔ Second Ctrl+C — exiting.")
            sys.exit(1)
    signal.signal(signal.SIGINT, _sigint)

    print("=" * 78)
    print("  M365 Group Membership Report")
    print("=" * 78)

    users = read_users(args.users)
    if not users:
        print("❌ no users in input")
        sys.exit(1)
    print(f"\nUsers: {len(users)}")
    print(f"User delay: {DELAY['user']}s\n")

    token = get_graph_token()
    if not token:
        print("❌ no Graph token")
        sys.exit(1)
    print("🔑 Auth OK\n")

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    prefix = args.out or f"m365_groups_{ts}"
    out_by_group = f"{prefix}_by_group.csv"
    out_by_user = f"{prefix}_by_user.csv"

    # state
    by_group: Dict[str, Dict] = {}
    pairs: List[Dict] = []
    errors: List[Dict] = []
    interrupted = False
    start = time.time()

    try:
        for i, upn in enumerate(users, start=1):
            if STOP_EVENT.is_set():
                break
            if i > 1:
                time.sleep(DELAY["user"])

            ok, groups, err = get_member_of(upn, token)
            if not ok:
                log(f"  ✗ {upn}: {err}")
                errors.append({"User": upn, "Error": err})
            else:
                log(f"  • {upn}: {len(groups)} group(s)")
                for g in groups:
                    gid = g["id"]
                    if gid not in by_group:
                        by_group[gid] = {"id": gid, "name": g["name"], "type": g["type"], "users": []}
                    if upn not in by_group[gid]["users"]:
                        by_group[gid]["users"].append(upn)
                    pairs.append({"User": upn, "Group ID": gid,
                                  "Group Name": g["name"], "Type": g["type"]})

            # write partials every user so Ctrl+C never loses progress
            write_by_group(out_by_group, by_group)
            write_by_user(out_by_user, pairs)

            # progress line
            if ok:
                print(f"[{i}/{len(users)}] {upn}  →  {len(groups)} group(s), "
                      f"total groups so far: {len(by_group)}")

    except KeyboardInterrupt:
        interrupted = True

    elapsed = time.time() - start
    write_by_group(out_by_group, by_group)
    write_by_user(out_by_user, pairs)

    # summary
    total_memberships = len(pairs)
    users_with_groups = len({p["User"] for p in pairs})
    users_with_none = len(users) - users_with_groups - len(errors)

    print()
    print("=" * 78)
    print("  SUMMARY" + ("  (INTERRUPTED)" if interrupted else ""))
    print("=" * 78)
    print(f"  Users               : {len(users)}")
    print(f"  Users with groups   : {users_with_groups}")
    print(f"  Users with no group : {users_with_none}")
    print(f"  Users errored       : {len(errors)}")
    print(f"  Unique groups       : {len(by_group)}")
    print(f"  Total memberships   : {total_memberships}")
    print(f"  Time                : {elapsed:.1f}s\n")
    print(f"  📄 By group : {out_by_group}")
    print(f"  📄 By user  : {out_by_user}")
    if errors:
        err_csv = f"{prefix}_errors.csv"
        with open(err_csv, "w", encoding="utf-8", newline="") as f:
            w = csv.DictWriter(f, fieldnames=["User", "Error"])
            w.writeheader()
            for e in errors:
                w.writerow(e)
        print(f"  ⚠  Errors   : {err_csv}")
    print("=" * 78)


if __name__ == "__main__":
    main()