#!/usr/bin/env python3
"""
Find / remove users in SharePoint siteusers.

Modes:
  --check  (default) : READ-ONLY. Report which sites have each user in siteusers.
  --delete           : Remove the user from siteusers on sites where found.

No group logic. Just:
  1. GET siteusers
  2. Look for the UPN in Email or LoginName
  3. If --delete and found -> POST siteusers/removebyid({id})

Output CSV per (site, user):
  Site URL, Site Name, User, Found, Site User ID, Email, LoginName, Status, Detail

Status values:
  FOUND         (check mode)  — user exists on this site
  NOT_FOUND     (both modes)  — user is not in siteusers (nothing changed)
  REMOVED       (delete mode) — user removed from siteusers
  FAILED        (delete mode) — removebyid blocked (detail has reason)

Usage:
  python find_or_remove_users.py --sites sites.csv --users users.csv --check
  python find_or_remove_users.py --sites sites.csv --users users.csv --delete
  python find_or_remove_users.py --sites sites.csv --users users.csv --delete --workers 8
"""

import argparse
import base64
import csv
import json
import os
import sys
import time
import uuid
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from threading import Lock
from typing import Dict, List, Optional, Tuple

import requests
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate

# ------------------------------------------------------------------ #
# CONFIG
# ------------------------------------------------------------------ #
CONFIG = {
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "tenant_name": "geekbyteonline.onmicrosoft.com",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "client_secret": "CyG8Q~FYHuCMSyVmt4sNxt5IejrMc2c24Ziz4a.t",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    "scopes": {"sharepoint": "https://geekbyteonline.sharepoint.com/.default"},
}

_token_cache: Dict[str, Tuple[str, float]] = {}
_token_lock = Lock()


def log(msg: str) -> None:
    print(msg, flush=True)


# ------------------------------------------------------------------ #
# AUTH
# ------------------------------------------------------------------ #
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


def get_sp_token() -> Optional[str]:
    scope = CONFIG["scopes"]["sharepoint"]
    with _token_lock:
        c = _token_cache.get(scope)
        if c and c[1] > time.time() + 60:
            return c[0]
    tok = _tok_cert(scope) or _tok_secret(scope)
    if tok:
        with _token_lock:
            _token_cache[scope] = (tok, time.time() + 3300)
    return tok


# ------------------------------------------------------------------ #
# HELPERS
# ------------------------------------------------------------------ #
def normalize_site_url(s: str) -> str:
    s = s.strip()
    if not s:
        return s
    if s.startswith("http"):
        return s.rstrip("/")
    if not s.startswith("sites/"):
        s = f"sites/{s}"
    tenant = CONFIG["tenant_name"].split(".")[0]
    return f"https://{tenant}.sharepoint.com/{s}".rstrip("/")


def get_digest(site_url: str, token: str) -> str:
    r = requests.post(
        f"{site_url.rstrip('/')}/_api/contextinfo",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json;odata=verbose"},
        timeout=30)
    if r.status_code != 200:
        raise Exception(f"digest {r.status_code}: {r.text[:200]}")
    return r.json()["d"]["GetContextWebInformation"]["FormDigestValue"]


def list_site_users(site_url: str, token: str) -> List[Dict]:
    """GET /_api/web/siteusers — read-only."""
    r = requests.get(
        f"{site_url.rstrip('/')}/_api/web/siteusers",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/xml"},
        timeout=60)
    if r.status_code != 200:
        raise Exception(f"siteusers GET {r.status_code}: {r.text[:200]}")
    ns = {"d": "http://schemas.microsoft.com/ado/2007/08/dataservices",
          "m": "http://schemas.microsoft.com/ado/2007/08/dataservices/metadata",
          "atom": "http://www.w3.org/2005/Atom"}
    out = []
    for entry in ET.fromstring(r.text).findall(".//atom:entry", ns):
        content = entry.find("atom:content", ns)
        if content is None:
            continue
        props = content.find("m:properties", ns)
        if props is None:
            continue

        def g(tag):
            e = props.find(f"d:{tag}", ns)
            return e.text if e is not None else None

        out.append({
            "Id": g("Id"),
            "Email": g("Email"),
            "LoginName": g("LoginName"),
            "Title": g("Title"),
        })
    return out


def find_user(site_users: List[Dict], upn: str) -> Optional[Dict]:
    upn_l = upn.lower()
    for u in site_users:
        email = (u.get("Email") or "").lower()
        login = (u.get("LoginName") or "").lower()
        if email == upn_l or login == upn_l or login.endswith("|" + upn_l):
            return u
    return None


def remove_from_site(site_url: str, token: str, digest: str, uid: str) -> Tuple[bool, str]:
    """POST /_api/web/siteusers/removebyid({uid})."""
    r = requests.post(
        f"{site_url.rstrip('/')}/_api/web/siteusers/removebyid({uid})",
        headers={"Authorization": f"Bearer {token}",
                 "Accept": "application/json;odata=verbose",
                 "Content-Type": "application/json;odata=verbose",
                 "X-RequestDigest": digest}, timeout=30)
    if r.status_code in (200, 204):
        return True, ""
    try:
        e = r.json().get("error", {})
        msg = e.get("message", {}).get("value") or str(e)
        code = e.get("code", "")
        return False, f"{code}: {msg}" if code else msg
    except Exception:
        return False, f"HTTP {r.status_code}: {r.text[:300]}"


# ------------------------------------------------------------------ #
# CORE per (site, user)
# ------------------------------------------------------------------ #
def process_one(site_url: str, upn: str, token: str, do_delete: bool) -> Dict[str, str]:
    """
    Returns: status, detail, user_id, email, login_name
    status: FOUND | NOT_FOUND | REMOVED | FAILED
    """
    # 1. list site users
    try:
        users = list_site_users(site_url, token)
    except Exception as e:
        return {"status": "FAILED", "detail": f"list siteusers: {e}",
                "user_id": "", "email": "", "login_name": ""}

    # 2. find user
    found = find_user(users, upn)
    if not found:
        return {"status": "NOT_FOUND",
                "detail": "User not present in siteusers (no action taken)",
                "user_id": "", "email": "", "login_name": ""}

    uid = str(found["Id"])
    email = found.get("Email") or ""
    login = found.get("LoginName") or ""

    # 3a. check mode — report only
    if not do_delete:
        return {"status": "FOUND",
                "detail": f"Present in siteusers (id={uid})",
                "user_id": uid, "email": email, "login_name": login}

    # 3b. delete mode — remove
    try:
        digest = get_digest(site_url, token)
    except Exception as e:
        return {"status": "FAILED", "detail": f"digest: {e}",
                "user_id": uid, "email": email, "login_name": login}

    ok, err = remove_from_site(site_url, token, digest, uid)
    if ok:
        return {"status": "REMOVED",
                "detail": "Removed from siteusers",
                "user_id": uid, "email": email, "login_name": login}

    return {"status": "FAILED",
            "detail": f"{err} | (user still in siteusers)",
            "user_id": uid, "email": email, "login_name": login}


# ------------------------------------------------------------------ #
# CSV INPUT
# ------------------------------------------------------------------ #
def read_sites(path: str) -> List[Dict[str, str]]:
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    out = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        rd = csv.DictReader(f)
        if not rd.fieldnames:
            raise ValueError("sites CSV has no header")
        url_col = next((c for c in ["Site URL", "Site", "site_url", "site", "URL", "url", "Url"]
                        if c in rd.fieldnames), None)
        if not url_col:
            raise ValueError(f"sites CSV missing URL column: {rd.fieldnames}")
        name_col = next((c for c in ["Site Name", "Name", "site_name", "name"]
                         if c in rd.fieldnames), None)
        for row in rd:
            url = (row.get(url_col) or "").strip()
            if not url:
                continue
            name = (row.get(name_col) or "").strip() if name_col else url
            out.append({"input": url, "name": name})
    return out


def read_users(path: str) -> List[str]:
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    out = []
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
            out.append(v)
    return out


# ------------------------------------------------------------------ #
# MAIN
# ------------------------------------------------------------------ #
def main():
    ap = argparse.ArgumentParser(
        description="Find (or remove) users in SharePoint siteusers. No group logic.")
    ap.add_argument("--sites", required=True, help="CSV with site URLs")
    ap.add_argument("--users", required=True, help="CSV with user UPNs")
    ap.add_argument("--out", default=None, help="Output CSV path")
    ap.add_argument("--workers", type=int, default=5, help="Parallel threads (default 5)")
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument("--check", action="store_true",
                      help="READ-ONLY: report which sites have the users (default)")
    mode.add_argument("--delete", action="store_true",
                      help="Actually remove found users from siteusers")
    args = ap.parse_args()

    do_delete = bool(args.delete)
    mode_label = "DELETE (removes users from siteusers)" if do_delete else "CHECK (read-only)"

    print("=" * 78)
    print(f"  Mode: {mode_label}")
    print("=" * 78)

    sites = read_sites(args.sites)
    users = read_users(args.users)
    if not sites or not users:
        print("❌ empty input")
        sys.exit(1)

    print(f"\nSites: {len(sites)} | Users: {len(users)} | Pairs: {len(sites)*len(users)}")
    print(f"Workers: {args.workers}\n")

    token = get_sp_token()
    if not token:
        print("❌ no SharePoint token")
        sys.exit(1)
    print("🔑 Auth OK\n")

    tasks = [{"site_url": normalize_site_url(s["input"]), "site_name": s["name"], "upn": u}
             for s in sites for u in users]

    results: List[Dict] = []
    start = time.time()
    done = 0

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
        fut_map = {ex.submit(process_one, t["site_url"], t["upn"], token, do_delete): t
                   for t in tasks}
        for fut in as_completed(fut_map):
            t = fut_map[fut]
            try:
                r = fut.result()
            except Exception as e:
                r = {"status": "FAILED", "detail": f"unhandled: {e}",
                     "user_id": "", "email": "", "login_name": ""}
            done += 1
            results.append({
                "Site URL": t["site_url"],
                "Site Name": t["site_name"],
                "User": t["upn"],
                "Status": r["status"],
                "Detail": r["detail"],
                "Site User ID": r["user_id"],
                "Email": r["email"],
                "LoginName": r["login_name"],
            })
            icon = {
                "FOUND": "✅",
                "NOT_FOUND": "⊘",
                "REMOVED": "🗑",
                "FAILED": "❌",
            }.get(r["status"], "•")
            print(f"[{done}/{len(tasks)}] {icon} {t['upn']:<40} @ {t['site_url']:<55} → {r['status']}")

    elapsed = time.time() - start
    results.sort(key=lambda x: (x["Site URL"].lower(), x["User"].lower()))

    out = args.out or f"user_{'remove' if do_delete else 'check'}_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    with open(out, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["Site URL", "Site Name", "User", "Status",
                                          "Detail", "Site User ID", "Email", "LoginName"])
        w.writeheader()
        for r in results:
            w.writerow(r)

    counts: Dict[str, int] = {}
    for r in results:
        counts[r["Status"]] = counts.get(r["Status"], 0) + 1

    print()
    print("=" * 78)
    print("  SUMMARY")
    print("=" * 78)
    print(f"  Total pairs : {len(results)}")
    print(f"  Time        : {elapsed:.1f}s\n")
    for k in ("FOUND", "REMOVED", "NOT_FOUND", "FAILED"):
        if k in counts:
            print(f"  {k:<10}: {counts[k]}")
    print(f"\n  📄 Report: {out}")
    print("=" * 78)


if __name__ == "__main__":
    main()