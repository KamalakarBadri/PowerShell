#!/usr/bin/env python3
"""
Remove Users from SharePoint Sites (CLI) — with detailed debug

Lookup-only: never calls ensureuser. If the user is not on the site, nothing changes.

Debug output shows, per (site, user):
  ▶ STEP 1 — GET siteusers
  ▶ STEP 2 — find user by UPN
  ▶ STEP 3 — determine membership (DIRECT / GROUP / BOTH)
  ▶ STEP 4 — remove from each group
  ▶ STEP 5 — remove from site

Run:
  python remove_users_from_sites.py --sites sites.csv --users users.csv --debug
"""

import argparse
import base64
import csv
import json
import logging
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

# Debug flags (set by CLI)
DEBUG = False           # print detailed step trace
DEBUG_HTTP = False      # include request/response bodies


def log(msg: str, level: str = "info") -> None:
    """Central logger with level-aware formatting."""
    prefix = {
        "info":    "ℹ",
        "debug":   "·",
        "step":    "▶",
        "ok":      "✓",
        "fail":    "✗",
        "skip":    "⊘",
        "warn":    "⚠",
        "arrow":   "→",
    }.get(level, "•")
    print(f"{prefix} {msg}", flush=True)


_token_cache: Dict[str, Tuple[str, float]] = {}
_token_lock = Lock()


# ------------------------------------------------------------------ #
# AUTH
# ------------------------------------------------------------------ #
def _tok_cert(scope: str) -> Optional[str]:
    log(f"AUTH: trying certificate flow (cert={CONFIG['certificate_path']}, key={CONFIG['private_key_path']})", "debug")
    try:
        if not os.path.exists(CONFIG["certificate_path"]) or not os.path.exists(CONFIG["private_key_path"]):
            log("AUTH: certificate or key file missing → skipping cert flow", "warn")
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
        log(f"AUTH(cert): HTTP {r.status_code}", "debug")
        if r.status_code == 200:
            log("AUTH(cert): token acquired", "ok")
            return r.json()["access_token"]
        log(f"AUTH(cert): failed — {r.text[:200]}", "fail")
    except Exception as e:
        log(f"AUTH(cert): exception — {e}", "fail")
    return None


def _tok_secret(scope: str) -> Optional[str]:
    log("AUTH: trying client-secret flow", "debug")
    try:
        r = requests.post(
            f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            data={"client_id": CONFIG["app_id"], "client_secret": CONFIG["client_secret"],
                  "scope": scope, "grant_type": "client_credentials"}, timeout=30)
        log(f"AUTH(secret): HTTP {r.status_code}", "debug")
        if r.status_code == 200:
            log("AUTH(secret): token acquired", "ok")
            return r.json()["access_token"]
        log(f"AUTH(secret): failed — {r.text[:200]}", "fail")
    except Exception as e:
        log(f"AUTH(secret): exception — {e}", "fail")
    return None


def get_sp_token() -> Optional[str]:
    scope = CONFIG["scopes"]["sharepoint"]
    with _token_lock:
        c = _token_cache.get(scope)
        if c and c[1] > time.time() + 60:
            log(f"AUTH: using cached token (expires in {int(c[1]-time.time())}s)", "debug")
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
        out = s.rstrip("/")
    elif s.startswith("sites/"):
        tenant = CONFIG["tenant_name"].split(".")[0]
        out = f"https://{tenant}.sharepoint.com/{s}".rstrip("/")
    else:
        tenant = CONFIG["tenant_name"].split(".")[0]
        out = f"https://{tenant}.sharepoint.com/sites/{s}".rstrip("/")
    log(f"URL: '{s}' → '{out}'", "debug")
    return out


def get_digest(site_url: str, token: str) -> str:
    url = f"{site_url.rstrip('/')}/_api/contextinfo"
    log(f"HTTP POST {url}  (get digest)", "debug")
    r = requests.post(url,
                      headers={"Authorization": f"Bearer {token}",
                               "Accept": "application/json;odata=verbose"},
                      timeout=30)
    log(f"HTTP POST contextinfo → {r.status_code}", "debug")
    if DEBUG_HTTP and r.status_code != 200:
        log(f"  body: {r.text[:400]}", "debug")
    if r.status_code != 200:
        raise Exception(f"digest {r.status_code}: {r.text[:200]}")
    dg = r.json()["d"]["GetContextWebInformation"]["FormDigestValue"]
    log(f"  digest acquired: {dg[:20]}… (len={len(dg)})", "debug")
    return dg


def list_site_users(site_url: str, token: str) -> List[Dict]:
    url = f"{site_url.rstrip('/')}/_api/web/siteusers"
    log(f"HTTP GET  {url}", "debug")
    r = requests.get(url,
                     headers={"Authorization": f"Bearer {token}", "Accept": "application/xml"},
                     timeout=60)
    log(f"HTTP GET  siteusers → {r.status_code} (bytes={len(r.content)})", "debug")
    if DEBUG_HTTP and r.status_code != 200:
        log(f"  body: {r.text[:400]}", "debug")
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
            "PrincipalType": int(g("PrincipalType")) if g("PrincipalType") else None,
        })
    log(f"  parsed {len(out)} site users", "debug")
    if DEBUG and out:
        for u in out[:5]:
            log(f"    id={u['Id']} email={u['Email']} login={u['LoginName']} type={u['PrincipalType']}", "debug")
        if len(out) > 5:
            log(f"    … and {len(out)-5} more", "debug")
    return out


def find_user(site_users: List[Dict], upn: str) -> Optional[Dict]:
    upn_l = upn.lower()
    log(f"  matching UPN '{upn_l}' against {len(site_users)} site users", "debug")
    for u in site_users:
        email = (u.get("Email") or "").lower()
        login = (u.get("LoginName") or "").lower()
        if email == upn_l:
            log(f"    MATCH email==upn → id={u['Id']}", "debug")
            return u
        if login == upn_l:
            log(f"    MATCH login==upn → id={u['Id']}", "debug")
            return u
        if login.endswith("|" + upn_l):
            log(f"    MATCH login.endswith(|upn) → id={u['Id']}", "debug")
            return u
    return None


def list_site_groups(site_url: str, token: str) -> List[Dict]:
    url = f"{site_url.rstrip('/')}/_api/web/sitegroups?$select=Id,Title"
    log(f"HTTP GET  {url}", "debug")
    r = requests.get(url,
                     headers={"Authorization": f"Bearer {token}",
                              "Accept": "application/json;odata=verbose"},
                     timeout=30)
    log(f"HTTP GET  sitegroups → {r.status_code}", "debug")
    if r.status_code != 200:
        log(f"  groups fetch failed: {r.text[:200]}", "warn")
        return []
    groups = r.json().get("d", {}).get("results", [])
    log(f"  {len(groups)} group(s) on site: {[g.get('Title') for g in groups][:8]}", "debug")
    return groups


def user_groups(site_url: str, token: str, user_id: str) -> List[Tuple[int, str]]:
    hits = []
    for g in list_site_groups(site_url, token):
        gid, gtitle = g.get("Id"), g.get("Title", "")
        if not gid:
            continue
        url = f"{site_url.rstrip('/')}/_api/web/sitegroups({gid})/users?$filter=Id eq {user_id}"
        log(f"HTTP GET  {url}", "debug")
        r = requests.get(url,
                         headers={"Authorization": f"Bearer {token}",
                                  "Accept": "application/json;odata=verbose"},
                         timeout=20)
        n = 0
        if r.status_code == 200:
            n = len(r.json().get("d", {}).get("results", []))
        log(f"    group {gid} '{gtitle}' → HTTP {r.status_code}, members matching user={n}", "debug")
        if n:
            hits.append((gid, gtitle))
    return hits


def remove_from_group(site_url: str, token: str, digest: str, gid: int, uid: str) -> Tuple[bool, str]:
    url = f"{site_url.rstrip('/')}/_api/web/sitegroups({gid})/users/removebyid({uid})"
    log(f"HTTP POST {url}", "debug")
    r = requests.post(url,
                      headers={"Authorization": f"Bearer {token}",
                               "Accept": "application/json;odata=verbose",
                               "Content-Type": "application/json;odata=verbose",
                               "X-RequestDigest": digest}, timeout=30)
    log(f"HTTP POST removebyid(group={gid}) → {r.status_code}", "debug")
    if r.status_code in (200, 204):
        return True, ""
    if DEBUG_HTTP:
        log(f"  body: {r.text[:400]}", "debug")
    return False, f"HTTP {r.status_code}: {r.text[:200]}"


def remove_from_site(site_url: str, token: str, digest: str, uid: str) -> Tuple[bool, str]:
    url = f"{site_url.rstrip('/')}/_api/web/siteusers/removebyid({uid})"
    log(f"HTTP POST {url}", "debug")
    r = requests.post(url,
                      headers={"Authorization": f"Bearer {token}",
                               "Accept": "application/json;odata=verbose",
                               "Content-Type": "application/json;odata=verbose",
                               "X-RequestDigest": digest}, timeout=30)
    log(f"HTTP POST removebyid(site users) → {r.status_code}", "debug")
    if r.status_code in (200, 204):
        return True, ""
    try:
        e = r.json().get("error", {})
        msg = e.get("message", {}).get("value") or str(e)
        code = e.get("code", "")
        if DEBUG_HTTP:
            log(f"  error body: {json.dumps(e)[:400]}", "debug")
        return False, f"{code}: {msg}" if code else msg
    except Exception:
        if DEBUG_HTTP:
            log(f"  raw body: {r.text[:400]}", "debug")
        return False, f"HTTP {r.status_code}: {r.text[:300]}"


# ------------------------------------------------------------------ #
# CORE — remove_one with per-step debug
# ------------------------------------------------------------------ #
def remove_one(site_url: str, upn: str, token: str, dry_run: bool) -> Dict[str, str]:
    ctx = f"[{upn} @ {site_url}]"
    log(f"═══════════════════════════════════════════════════════════", "info")
    log(f"START {ctx}", "info")
    log(f"═══════════════════════════════════════════════════════════", "info")

    # ----- STEP 1: list site users -----
    log(f"▶ STEP 1 — GET /_api/web/siteusers  (read-only)", "step")
    try:
        users = list_site_users(site_url, token)
    except Exception as e:
        log(f"✗ STEP 1 failed: {e}", "fail")
        log(f"→ STATUS=FAILED  detail='list siteusers: {e}'", "arrow")
        return {"status": "FAILED", "membership": "", "detail": f"list siteusers: {e}",
                "groups_removed": "", "user_id": ""}
    log(f"✓ STEP 1 OK — {len(users)} site users", "ok")

    # ----- STEP 2: find user -----
    log(f"▶ STEP 2 — find '{upn}' in site users  (read-only)", "step")
    found = find_user(users, upn)
    if not found:
        log(f"⊘ STEP 2 — user NOT on this site", "skip")
        log(f"→ STATUS=NOT_FOUND  (nothing was created, nothing changed)", "arrow")
        return {"status": "NOT_FOUND", "membership": "",
                "detail": "User not present on site (no action taken)",
                "groups_removed": "", "user_id": ""}
    uid = str(found["Id"])
    log(f"✓ STEP 2 OK — id={uid} email={found.get('Email')} login={found.get('LoginName')}", "ok")

    # ----- STEP 3: determine membership -----
    log(f"▶ STEP 3 — determine membership (DIRECT / GROUP / BOTH)", "step")
    try:
        groups = user_groups(site_url, token, uid)
    except Exception as e:
        groups = []
        log(f"⚠ STEP 3 — group enumeration failed: {e}", "warn")

    ptype = found.get("PrincipalType")
    login = (found.get("LoginName") or "").lower()
    is_group_principal = (ptype == 4) or login.startswith("c:0o.c")
    has_group_memberships = len(groups) > 0

    if has_group_memberships and not is_group_principal:
        membership = "BOTH"
    elif has_group_memberships:
        membership = "GROUP"
    else:
        membership = "DIRECT"

    log(f"  principal_type={ptype}, is_group_principal={is_group_principal}", "debug")
    log(f"  group_memberships={len(groups)}: {[t for _, t in groups]}", "debug")
    log(f"✓ STEP 3 OK — membership={membership}", "ok")

    # ----- STEP 4: dry-run stop here -----
    if dry_run:
        log(f"🟡 DRY-RUN — would remove (groups={len(groups)}, direct={not is_group_principal})", "warn")
        log(f"→ STATUS=DRY_RUN  membership={membership}", "arrow")
        return {"status": "DRY_RUN", "membership": membership,
                "detail": f"Would remove (groups={len(groups)}, direct={not is_group_principal})",
                "groups_removed": str(len(groups)), "user_id": uid}

    # ----- STEP 5: remove from each group -----
    if groups:
        log(f"▶ STEP 4 — remove user from {len(groups)} group(s)", "step")
        try:
            digest = get_digest(site_url, token)
        except Exception as e:
            log(f"✗ STEP 4 failed at digest: {e}", "fail")
            log(f"→ STATUS=FAILED  detail='digest: {e}'", "arrow")
            return {"status": "FAILED", "membership": membership,
                    "detail": f"digest: {e}", "groups_removed": "0", "user_id": uid}

        groups_removed = 0
        for gid, gtitle in groups:
            ok, err = remove_from_group(site_url, token, digest, gid, uid)
            if ok:
                groups_removed += 1
                log(f"  ✓ removed from group {gid} '{gtitle}'", "ok")
            else:
                log(f"  ✗ failed to remove from group {gid} '{gtitle}': {err}", "fail")
        log(f"✓ STEP 4 done — {groups_removed}/{len(groups)} groups cleared", "ok")
    else:
        log(f"⊘ STEP 4 — user is not in any group, skipping group removal", "skip")
        groups_removed = 0
        try:
            digest = get_digest(site_url, token)
        except Exception as e:
            log(f"✗ could not get digest for final step: {e}", "fail")
            log(f"→ STATUS=FAILED  detail='digest: {e}'", "arrow")
            return {"status": "FAILED", "membership": membership,
                    "detail": f"digest: {e}", "groups_removed": "0", "user_id": uid}

    # ----- STEP 5: remove from site -----
    log(f"▶ STEP 5 — POST /_api/web/siteusers/removebyid({uid})", "step")
    ok, err = remove_from_site(site_url, token, digest, uid)
    if ok:
        log(f"✓ STEP 5 OK — user removed from site", "ok")
        log(f"→ STATUS=REMOVED  membership={membership}  groups_cleared={groups_removed}", "arrow")
        return {"status": "REMOVED", "membership": membership,
                "detail": f"Removed from site (groups cleared: {groups_removed})",
                "groups_removed": str(groups_removed), "user_id": uid}

    log(f"✗ STEP 5 failed: {err}", "fail")
    log(f"  HINT: SharePoint blocks siteusers/removebyid when the user has:", "warn")
    log(f"    · authored content (Created By / Modified By on any item)", "warn")
    log(f"    · unique item/folder permissions (broken inheritance)", "warn")
    log(f"    · Site Collection Admin flag set (clear IsSiteAdmin first)", "warn")
    log(f"→ STATUS=FAILED  membership={membership}", "arrow")
    return {"status": "FAILED", "membership": membership,
            "detail": f"{err} | groups cleared before failure: {groups_removed}",
            "groups_removed": str(groups_removed), "user_id": uid}


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
    log(f"CSV: loaded {len(out)} site(s) from {path}", "ok")
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
                log(f"CSV: skipping malformed UPN: {v!r}", "warn")
                continue
            out.append(v)
    log(f"CSV: loaded {len(out)} user(s) from {path}", "ok")
    return out


# ------------------------------------------------------------------ #
# MAIN
# ------------------------------------------------------------------ #
def main():
    global DEBUG, DEBUG_HTTP

    ap = argparse.ArgumentParser(description="Remove users from SharePoint sites (lookup-only, verbose debug).")
    ap.add_argument("--sites", required=True)
    ap.add_argument("--users", required=True)
    ap.add_argument("--out", default=None)
    ap.add_argument("--workers", type=int, default=1,
                    help="Parallel threads. Use 1 for clean per-step debug output (default: 1).")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--debug", action="store_true", help="Verbose step-by-step trace")
    ap.add_argument("--debug-http", action="store_true", help="Also print HTTP bodies on errors")
    args = ap.parse_args()

    DEBUG = args.debug
    DEBUG_HTTP = args.debug_http

    print("=" * 78)
    print("  Remove Users from SharePoint Sites — DEBUG BUILD")
    print("=" * 78)

    if not DEBUG:
        log("Tip: add --debug for step-by-step trace, --debug-http for response bodies", "info")

    sites = read_sites(args.sites)
    users = read_users(args.users)
    if not sites or not users:
        log("Empty input", "fail")
        sys.exit(1)

    log(f"Sites: {len(sites)} | Users: {len(users)} | Pairs: {len(sites)*len(users)}", "info")
    log(f"Workers: {args.workers}" + ("  |  DRY-RUN" if args.dry_run else ""), "info")

    token = get_sp_token()
    if not token:
        log("No SharePoint token — aborting", "fail")
        sys.exit(1)
    log("Auth OK — starting processing", "ok")
    print()

    tasks = [{"site_url": normalize_site_url(s["input"]), "site_name": s["name"], "upn": u}
             for s in sites for u in users]

    results: List[Dict] = []
    start = time.time()
    done = 0

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
        fut_map = {ex.submit(remove_one, t["site_url"], t["upn"], token, args.dry_run): t for t in tasks}
        for fut in as_completed(fut_map):
            t = fut_map[fut]
            try:
                r = fut.result()
            except Exception as e:
                log(f"✗ unhandled exception in {t['upn']} @ {t['site_url']}: {e}", "fail")
                r = {"status": "FAILED", "membership": "", "detail": f"unhandled: {e}",
                     "groups_removed": "", "user_id": ""}
            done += 1
            results.append({
                "Site URL": t["site_url"], "Site Name": t["site_name"], "User": t["upn"],
                "Status": r["status"], "Membership": r["membership"], "Detail": r["detail"],
                "Groups Removed": r["groups_removed"], "Site User ID": r["user_id"],
            })
            icon = {"REMOVED": "✅", "NOT_FOUND": "🔎", "FAILED": "❌", "DRY_RUN": "🟡"}.get(r["status"], "•")
            tag = f" [{r['membership']}]" if r["membership"] else ""
            print(f"[{done}/{len(tasks)}] {icon} {t['upn']:<40} @ {t['site_url']:<50}{tag} → {r['status']}")

    elapsed = time.time() - start
    results.sort(key=lambda x: (x["Site URL"].lower(), x["User"].lower()))

    out = args.out or f"remove_users_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    with open(out, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["Site URL", "Site Name", "User", "Status",
                                          "Membership", "Detail", "Groups Removed", "Site User ID"])
        w.writeheader()
        for r in results:
            w.writerow(r)

    counts = {}
    for r in results:
        counts[r["Status"]] = counts.get(r["Status"], 0) + 1

    print()
    print("=" * 78)
    print("  SUMMARY")
    print("=" * 78)
    print(f"  Total pairs : {len(results)}")
    print(f"  Time        : {elapsed:.1f}s")
    print()
    for k in ("REMOVED", "NOT_FOUND", "FAILED", "DRY_RUN"):
        if k in counts:
            print(f"  {k:<12}: {counts[k]}")
    print(f"\n  📄 Report: {out}")
    print("=" * 78)


if __name__ == "__main__":
    main()