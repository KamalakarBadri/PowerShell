#!/usr/bin/env python3
"""
Remove users from SharePoint sites — one-row-per-site report.

Per site (sequential):
  1. GET siteusers (once)
  2. For each user:
       a. find in siteusers  → skip if not found
       b. if IsSiteAdmin → MERGE IsSiteAdmin=false
       c. removebyid (remove from siteusers)
  3. Sleep --site-delay → next site

Output CSV — one row per site:
  Site URL, Site Name, Users Removed, Users Not Found

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


# ---------------- HTTP ----------------
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


def clear_admin(site_url: str, token: str, digest: str, uid: str) -> bool:
    r = http("POST", f"{site_url.rstrip('/')}/_api/web/siteusers/getbyid({uid})",
             token=token,
             headers={"Accept": "application/json;odata=verbose",
                      "Content-Type": "application/json;odata=verbose",
                      "X-RequestDigest": digest,
                      "IF-MATCH": "*",
                      "X-HTTP-Method": "MERGE"},
             data=json.dumps({"__metadata": {"type": "SP.User"}, "IsSiteAdmin": False}),
             timeout=30)
    return r.status_code in (200, 204)


def remove_from_site(site_url: str, token: str, digest: str, uid: str) -> bool:
    r = http("POST", f"{site_url.rstrip('/')}/_api/web/siteusers/removebyid({uid})",
             token=token,
             headers={"Accept": "application/json;odata=verbose",
                      "Content-Type": "application/json;odata=verbose",
                      "X-RequestDigest": digest}, timeout=30)
    return r.status_code in (200, 204)


# ---------------- CORE ----------------
def process_site(site_url: str, site_name: str, users: List[str],
                 token: str, dry_run: bool) -> Dict[str, str]:
    """Returns ONE row for the report."""
    removed: List[str] = []
    not_found: List[str] = []
    failed: List[str] = []
    note = ""

    if STOP_EVENT.is_set():
        raise KeyboardInterrupt("Stop requested")

    log(f"\n▶ Site: {site_url}")
    try:
        site_users = list_site_users(site_url, token)
        log(f"  siteusers: {len(site_users)} entries")
    except Exception as e:
        log(f"  ✗ list siteusers failed: {e}")
        return {
            "Site URL": site_url,
            "Site Name": site_name,
            "Users Removed": "",
            "Users Not Found": ", ".join(users),
            "Users Failed": "",
            "Note": f"list siteusers failed: {e}",
        }

    digest: Optional[str] = None

    for idx, upn in enumerate(users):
        if STOP_EVENT.is_set():
            raise KeyboardInterrupt("Stop requested")
        if idx > 0:
            time.sleep(DELAY["user"])

        found = find_user(site_users, upn)
        if not found:
            not_found.append(upn)
            log(f"  ⊘ {upn}: NOT_FOUND")
            continue

        uid = str(found["Id"])
        was_admin = bool(found.get("IsSiteAdmin"))

        if dry_run:
            removed.append(upn)
            log(f"  🟡 {upn}: would clear admin={was_admin} and remove")
            continue

        try:
            if digest is None:
                digest = get_digest(site_url, token)
        except Exception as e:
            failed.append(f"{upn} (digest: {e})")
            log(f"  ❌ {upn}: digest failed")
            continue

        if was_admin and not clear_admin(site_url, token, digest, uid):
            failed.append(f"{upn} (clear admin failed)")
            log(f"  ❌ {upn}: clear admin failed")
            continue

        if remove_from_site(site_url, token, digest, uid):
            removed.append(upn)
            log(f"  🗑 {upn}: REMOVED" + (" (was admin)" if was_admin else ""))
        else:
            failed.append(f"{upn} (removebyid failed)")
            log(f"  ❌ {upn}: removebyid failed")

    return {
        "Site URL": site_url,
        "Site Name": site_name,
        "Users Removed": ", ".join(removed),
        "Users Not Found": ", ".join(not_found),
        "Users Failed": ", ".join(failed),
        "Note": note,
    }


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
            if not v or "@" not in v:
                continue
            if v.lower() in seen:
                continue
            seen.add(v.lower())
            out.append(v)
    return out


def write_report(path: str, rows: List[Dict[str, str]]) -> None:
    fields = ["Site URL", "Site Name", "Users Removed", "Users Not Found", "Users Failed", "Note"]
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow(r)


# ---------------- MAIN ----------------
def main():
    ap = argparse.ArgumentParser(description="Remove users from sites — one-row-per-site report.")
    ap.add_argument("--sites", required=True)
    ap.add_argument("--users", required=True)
    ap.add_argument("--out", default=None)
    ap.add_argument("--site-delay", type=float, default=2.0)
    ap.add_argument("--user-delay", type=float, default=0.5)
    ap.add_argument("--retry-wait", type=float, default=10.0)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    DELAY["site"] = args.site_delay
    DELAY["user"] = args.user_delay
    DELAY["retry"] = args.retry_wait

    def _sigint(signum, frame):
        if not STOP_EVENT.is_set():
            log("\n\n⛔ Ctrl+C — stopping (partial CSV saved)…")
            STOP_EVENT.set()
        else:
            log("⛔ Second Ctrl+C — exiting.")
            sys.exit(1)
    signal.signal(signal.SIGINT, _sigint)

    print("=" * 78)
    print(f"  Mode: {'DRY-RUN' if args.dry_run else 'DELETE'}")
    print(f"  Report: one row per site (comma-separated user lists)")
    print("=" * 78)

    sites = read_sites(args.sites)
    users = read_users(args.users)
    if not sites or not users:
        print("❌ empty input")
        sys.exit(1)
    print(f"\nSites: {len(sites)} | Users: {len(users)}\n")

    token = get_token()
    if not token:
        print("❌ no SharePoint token")
        sys.exit(1)
    print("🔑 Auth OK\n")

    out = args.out or f"remove_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    rows: List[Dict] = []
    interrupted = False
    start = time.time()

    try:
        for i, s in enumerate(sites, start=1):
            if STOP_EVENT.is_set():
                break
            site_url = normalize_site_url(s["input"])
            try:
                row = process_site(site_url, s["name"], users, token, args.dry_run)
            except KeyboardInterrupt:
                interrupted = True
                break
            rows.append(row)
            write_report(out, rows)

            # site summary line
            n_ok = len([x for x in row["Users Removed"].split(", ") if x])
            n_nf = len([x for x in row["Users Not Found"].split(", ") if x])
            n_f = len([x for x in row["Users Failed"].split(", ") if x])
            print(f"[site {i}/{len(sites)}] {site_url}  →  removed={n_ok} not_found={n_nf} failed={n_f}")

            if i < len(sites):
                for _ in range(int(DELAY["site"] * 10)):
                    if STOP_EVENT.is_set():
                        break
                    time.sleep(0.1)
    except KeyboardInterrupt:
        interrupted = True

    elapsed = time.time() - start
    write_report(out, rows)

    print()
    print("=" * 78)
    print("  DONE" + ("  (INTERRUPTED)" if interrupted else ""))
    print("=" * 78)
    print(f"  Sites processed : {len(rows)}/{len(sites)}")
    print(f"  Time            : {elapsed:.1f}s")
    print(f"  📄 Report       : {out}")
    print("=" * 78)


if __name__ == "__main__":
    main()