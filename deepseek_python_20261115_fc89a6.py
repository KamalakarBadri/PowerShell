#!/usr/bin/env python3
"""
Remove users from SharePoint sites — sequential, admin-clear first.

Per site (one at a time):
  1. GET siteusers  (once)
  2. For each user:
       a. find in siteusers  → NOT_FOUND if missing
       b. if IsSiteAdmin == true → POST MERGE IsSiteAdmin=false  (clear admin flag)
       c. POST siteusers/removebyid(id)                          (remove from site)
       d. status REMOVED / FAILED (with reason)

Notes:
  - No ensureuser. Users are never added.
  - No group logic. No Graph.
  - Ctrl+C stops cleanly at the next checkpoint; partial CSV saved after each site.
  - 429 retry with Retry-After.

Usage:
  python remove_users_from_sites.py --sites sites.csv --users users.csv
  python remove_users_from_sites.py --sites sites.csv --users users.csv --dry-run
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
import xml.etree.ElementTree as ET
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
    "scopes": {"sharepoint": "https://geekbyteonline.sharepoint.com/.default"},
}

STOP_EVENT = Event()
DELAY = {"site": 2.0, "user": 0.5, "retry": 10.0}


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
    except Exception:
        pass
    return None


def _tok_secret(scope: str) -> Optional[str]:
    try:
        r = requests.post(
            f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            data={"client_id": CONFIG["app_id"], "client_secret": CONFIG["client_secret"],
                  "scope": scope, "grant_type": "client_credentials"}, timeout=30)
        if r.status_code == 200:
            return r.json()["access_token"]
    except Exception:
        pass
    return None


def get_token() -> Optional[str]:
    scope = CONFIG["scopes"]["sharepoint"]
    with _token_lock:
        c = _token_cache.get("sp")
        if c and c[1] > time.time() + 60:
            return c[0]
    tok = _tok_cert(scope) or _tok_secret(scope)
    if tok:
        with _token_lock:
            _token_cache["sp"] = (tok, time.time() + 3300)
    return tok


# ---------------- HTTP with 429 retry ----------------
def http(method: str, url: str, *, token: str, **kw) -> requests.Response:
    headers = kw.pop("headers", {}) or {}
    headers.setdefault("Authorization", f"Bearer {token}")
    timeout = kw.pop("timeout", 60)
    attempt = 0
    while True:
        if STOP_EVENT.is_set():
            raise KeyboardInterrupt("Stop requested")
        try:
            r = requests.request(method, url, headers=headers, timeout=timeout, **kw)
        except requests.RequestException:
            attempt += 1
            if attempt > 3:
                raise
            time.sleep(2 ** attempt)
            continue
        if r.status_code == 429:
            wait = int(r.headers.get("Retry-After", DELAY["retry"]))
            log(f"    ⏸ 429 on {method} {url[:80]} — waiting {wait}s")
            for _ in range(wait):
                if STOP_EVENT.is_set():
                    raise KeyboardInterrupt("Stop requested")
                time.sleep(1)
            attempt += 1
            if attempt > 5:
                return r
            continue
        return r


# ---------------- HELPERS ----------------
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
    r = http("POST", f"{site_url.rstrip('/')}/_api/contextinfo", token=token,
             headers={"Accept": "application/json;odata=verbose"}, timeout=30)
    if r.status_code != 200:
        raise Exception(f"digest {r.status_code}: {r.text[:200]}")
    return r.json()["d"]["GetContextWebInformation"]["FormDigestValue"]


def list_site_users(site_url: str, token: str) -> List[Dict]:
    r = http("GET", f"{site_url.rstrip('/')}/_api/web/siteusers", token=token,
             headers={"Accept": "application/xml"}, timeout=60)
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
            "IsSiteAdmin": (g("IsSiteAdmin") or "").lower() == "true",
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


def clear_site_admin(site_url: str, token: str, digest: str, uid: str) -> Tuple[bool, str]:
    """MERGE IsSiteAdmin=false on the site user."""
    r = http("POST", f"{site_url.rstrip('/')}/_api/web/siteusers/getbyid({uid})",
             token=token,
             headers={"Accept": "application/json;odata=verbose",
                      "Content-Type": "application/json;odata=verbose",
                      "X-RequestDigest": digest,
                      "IF-MATCH": "*",
                      "X-HTTP-Method": "MERGE"},
             data=json.dumps({"__metadata": {"type": "SP.User"}, "IsSiteAdmin": False}),
             timeout=30)
    if r.status_code in (200, 204):
        return True, ""
    try:
        e = r.json().get("error", {})
        msg = e.get("message", {}).get("value") or str(e)
        code = e.get("code", "")
        return False, f"{code}: {msg}" if code else msg
    except Exception:
        return False, f"HTTP {r.status_code}: {r.text[:300]}"


def remove_from_site(site_url: str, token: str, digest: str, uid: str) -> Tuple[bool, str]:
    r = http("POST", f"{site_url.rstrip('/')}/_api/web/siteusers/removebyid({uid})",
             token=token,
             headers={"Accept": "application/json;odata=verbose",
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


# ---------------- CORE ----------------
def process_site(site_url: str, site_name: str, users: List[str],
                 token: str, dry_run: bool) -> List[Dict[str, str]]:
    results: List[Dict[str, str]] = []
    if STOP_EVENT.is_set():
        raise KeyboardInterrupt("Stop requested")

    log(f"\n▶ Site: {site_url}")
    try:
        site_users = list_site_users(site_url, token)
        log(f"  siteusers: {len(site_users)} entries")
    except Exception as e:
        log(f"  ✗ list siteusers failed: {e}")
        for upn in users:
            results.append({
                "Site URL": site_url, "Site Name": site_name, "User": upn,
                "Status": "FAILED", "Detail": f"list siteusers: {e}",
                "Site User ID": "", "Email": "", "LoginName": "",
                "Was Site Admin": "", "Admin Cleared": "", "Removed From Site": "",
            })
        return results

    digest: Optional[str] = None

    for idx, upn in enumerate(users):
        if STOP_EVENT.is_set():
            raise KeyboardInterrupt("Stop requested")
        if idx > 0:
            time.sleep(DELAY["user"])

        found = find_user(site_users, upn)
        if not found:
            log(f"  ⊘ {upn}: NOT_FOUND on site")
            results.append({
                "Site URL": site_url, "Site Name": site_name, "User": upn,
                "Status": "NOT_FOUND", "Detail": "User not present in siteusers",
                "Site User ID": "", "Email": "", "LoginName": "",
                "Was Site Admin": "", "Admin Cleared": "", "Removed From Site": "",
            })
            continue

        uid = str(found["Id"])
        email = found.get("Email") or ""
        login = found.get("LoginName") or ""
        was_admin = bool(found.get("IsSiteAdmin"))
        log(f"  • {upn}: found (id={uid}, siteAdmin={was_admin})")

        admin_cleared = ""
        removed = ""

        if dry_run:
            admin_cleared = "DRY_RUN" if was_admin else "N/A"
            removed = "DRY_RUN"
            log(f"    🟡 {upn}: would clear admin={was_admin}, then remove")
            results.append({
                "Site URL": site_url, "Site Name": site_name, "User": upn,
                "Status": "DRY_RUN", "Detail": "Would clear admin (if needed) and remove",
                "Site User ID": uid, "Email": email, "LoginName": login,
                "Was Site Admin": "TRUE" if was_admin else "FALSE",
                "Admin Cleared": admin_cleared, "Removed From Site": removed,
            })
            continue

        # get digest lazily
        try:
            if digest is None:
                digest = get_digest(site_url, token)
        except Exception as e:
            log(f"    ❌ {upn}: digest failed — {e}")
            results.append({
                "Site URL": site_url, "Site Name": site_name, "User": upn,
                "Status": "FAILED", "Detail": f"digest: {e}",
                "Site User ID": uid, "Email": email, "LoginName": login,
                "Was Site Admin": "TRUE" if was_admin else "FALSE",
                "Admin Cleared": "", "Removed From Site": "",
            })
            continue

        # step 1: clear admin if needed
        if was_admin:
            ok_adm, err_adm = clear_site_admin(site_url, token, digest, uid)
            if ok_adm:
                admin_cleared = "YES"
                log(f"    🔓 {upn}: site admin flag cleared")
            else:
                admin_cleared = f"FAILED: {err_adm}"
                log(f"    ⚠ {upn}: could not clear admin — {err_adm}")
                results.append({
                    "Site URL": site_url, "Site Name": site_name, "User": upn,
                    "Status": "FAILED", "Detail": f"clear admin failed: {err_adm} | user still on site",
                    "Site User ID": uid, "Email": email, "LoginName": login,
                    "Was Site Admin": "TRUE", "Admin Cleared": admin_cleared,
                    "Removed From Site": "NO",
                })
                continue
        else:
            admin_cleared = "N/A"

        # step 2: remove from siteusers
        ok_rem, err_rem = remove_from_site(site_url, token, digest, uid)
        if ok_rem:
            removed = "YES"
            log(f"    🗑 {upn}: REMOVED")
            results.append({
                "Site URL": site_url, "Site Name": site_name, "User": upn,
                "Status": "REMOVED", "Detail": "Removed from siteusers",
                "Site User ID": uid, "Email": email, "LoginName": login,
                "Was Site Admin": "TRUE" if was_admin else "FALSE",
                "Admin Cleared": admin_cleared, "Removed From Site": removed,
            })
        else:
            removed = "NO"
            log(f"    ❌ {upn}: removebyid failed — {err_rem}")
            results.append({
                "Site URL": site_url, "Site Name": site_name, "User": upn,
                "Status": "FAILED", "Detail": f"removebyid: {err_rem} | user still on site",
                "Site User ID": uid, "Email": email, "LoginName": login,
                "Was Site Admin": "TRUE" if was_admin else "FALSE",
                "Admin Cleared": admin_cleared, "Removed From Site": removed,
            })

    return results


# ---------------- CSV ----------------
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


def _write_csv(path: str, fields: List[str], rows: List[Dict]) -> None:
    rows_sorted = sorted(rows, key=lambda x: (x["Site URL"].lower(), x["User"].lower()))
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows_sorted:
            w.writerow(r)


# ---------------- MAIN ----------------
def main():
    ap = argparse.ArgumentParser(
        description="Remove users from SharePoint sites — clears admin flag first, then removes from siteusers.")
    ap.add_argument("--sites", required=True)
    ap.add_argument("--users", required=True)
    ap.add_argument("--out", default=None)
    ap.add_argument("--site-delay", type=float, default=2.0, help="Seconds between sites")
    ap.add_argument("--user-delay", type=float, default=0.5, help="Seconds between users on same site")
    ap.add_argument("--retry-wait", type=float, default=10.0, help="Fallback wait on 429")
    ap.add_argument("--dry-run", action="store_true", help="Report only, no changes")
    args = ap.parse_args()

    DELAY["site"] = args.site_delay
    DELAY["user"] = args.user_delay
    DELAY["retry"] = args.retry_wait

    def _sigint(signum, frame):
        if not STOP_EVENT.is_set():
            log("\n\n⛔ Ctrl+C — stopping at next checkpoint (partial CSV saved)…")
            STOP_EVENT.set()
        else:
            log("⛔ Second Ctrl+C — exiting.")
            sys.exit(1)
    signal.signal(signal.SIGINT, _sigint)

    print("=" * 78)
    print(f"  Mode: {'DRY-RUN' if args.dry_run else 'DELETE'}")
    print(f"  Per user: clear IsSiteAdmin (if set) → removebyid(siteusers)")
    print(f"  Delays: site={DELAY['site']}s user={DELAY['user']}s")
    print("=" * 78)

    sites = read_sites(args.sites)
    users = read_users(args.users)
    if not sites or not users:
        print("❌ empty input")
        sys.exit(1)
    print(f"\nSites: {len(sites)} | Users: {len(users)} | Pairs: {len(sites)*len(users)}\n")

    token = get_token()
    if not token:
        print("❌ no SharePoint token")
        sys.exit(1)
    print("🔑 Auth OK\n")

    out = args.out or f"remove_users_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    fields = ["Site URL", "Site Name", "User", "Status", "Detail",
              "Site User ID", "Email", "LoginName",
              "Was Site Admin", "Admin Cleared", "Removed From Site"]

    all_results: List[Dict] = []
    interrupted = False
    start = time.time()

    try:
        for i, s in enumerate(sites, start=1):
            if STOP_EVENT.is_set():
                break
            site_url = normalize_site_url(s["input"])
            try:
                site_results = process_site(site_url, s["name"], users, token, args.dry_run)
            except KeyboardInterrupt:
                interrupted = True
                break

            all_results.extend(site_results)
            _write_csv(out, fields, all_results)

            counts = {}
            for r in site_results:
                counts[r["Status"]] = counts.get(r["Status"], 0) + 1
            bits = ", ".join(f"{k}={v}" for k, v in sorted(counts.items()))
            print(f"[site {i}/{len(sites)}] {site_url}  →  {bits}")

            if i < len(sites):
                for _ in range(int(DELAY["site"] * 10)):
                    if STOP_EVENT.is_set():
                        break
                    time.sleep(0.1)
    except KeyboardInterrupt:
        interrupted = True

    elapsed = time.time() - start
    if all_results:
        _write_csv(out, fields, all_results)

    totals: Dict[str, int] = {}
    for r in all_results:
        totals[r["Status"]] = totals.get(r["Status"], 0) + 1

    print()
    print("=" * 78)
    print("  SUMMARY" + ("  (INTERRUPTED)" if interrupted else ""))
    print("=" * 78)
    print(f"  Sites reported  : {len({r['Site URL'] for r in all_results})}/{len(sites)}")
    print(f"  Users           : {len(users)}")
    print(f"  Pairs reported  : {len(all_results)}")
    print(f"  Time            : {elapsed:.1f}s\n")
    for k in ("REMOVED", "NOT_FOUND", "FAILED", "DRY_RUN"):
        if k in totals:
            print(f"  {k:<10}: {totals[k]}")
    if all_results:
        print(f"\n  📄 Report: {out}")
    print("=" * 78)


if __name__ == "__main__":
    main()