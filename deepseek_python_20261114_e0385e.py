#!/usr/bin/env python3
"""
Remove Users from SharePoint Sites (CLI)

For every (site, user) pair:
  1. Resolve the user on the site (ensureuser -> ID)
  2. If not resolvable        -> NOT_FOUND
  3. Remove user from all SharePoint groups on that site
  4. Remove the user from the site's siteusers list (removebyid)

Output CSV (one row per site+user):
  Site URL, Site Name, User, Status, Detail, Groups Removed, Site User ID

Status values:
  REMOVED     — user fully removed from site
  NOT_FOUND   — user not resolvable / not on site
  FAILED      — removebyid or group removal blocked (detail has the SharePoint error)

Usage:
  python remove_users_from_sites.py --sites sites.csv --users users.csv
  python remove_users_from_sites.py --sites sites.csv --users users.csv --out report.csv --workers 8
  python remove_users_from_sites.py --sites sites.csv --users users.csv --dry-run
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
# CONFIG  — update these for your tenant
# ------------------------------------------------------------------ #
CONFIG = {
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "tenant_name": "geekbyteonline.onmicrosoft.com",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "client_secret": "CyG8Q~FYHuCMSyVmt4sNxt5IejrMc2c24Ziz4a.t",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    "scopes": {
        "sharepoint": "https://geekbyteonline.sharepoint.com/.default",
    },
}

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
log = logging.getLogger(__name__)

_token_cache: Dict[str, Tuple[str, float]] = {}
_token_lock = Lock()


# ------------------------------------------------------------------ #
# AUTH
# ------------------------------------------------------------------ #
def _get_token_with_certificate(scope: str) -> Optional[str]:
    try:
        if not os.path.exists(CONFIG["certificate_path"]) or not os.path.exists(CONFIG["private_key_path"]):
            return None
        with open(CONFIG["certificate_path"], "rb") as cf:
            cert = load_pem_x509_certificate(cf.read(), default_backend())
        with open(CONFIG["private_key_path"], "rb") as kf:
            key = load_pem_private_key(kf.read(), password=None, backend=default_backend())

        now = int(time.time())
        header = {
            "alg": "RS256",
            "typ": "JWT",
            "x5t": base64.urlsafe_b64encode(cert.fingerprint(hashes.SHA1())).decode().rstrip("="),
        }
        payload = {
            "aud": f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            "exp": now + 300,
            "iss": CONFIG["app_id"],
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": CONFIG["app_id"],
        }
        eh = base64.urlsafe_b64encode(json.dumps(header).encode()).decode().rstrip("=")
        ep = base64.urlsafe_b64encode(json.dumps(payload).encode()).decode().rstrip("=")
        unsigned = f"{eh}.{ep}"
        sig = key.sign(unsigned.encode(), padding.PKCS1v15(), hashes.SHA256())
        es = base64.urlsafe_b64encode(sig).decode().rstrip("=")
        jwt = f"{unsigned}.{es}"

        r = requests.post(
            f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            data={
                "client_id": CONFIG["app_id"],
                "client_assertion": jwt,
                "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                "scope": scope,
                "grant_type": "client_credentials",
            },
            timeout=30,
        )
        if r.status_code == 200:
            return r.json()["access_token"]
        log.error("Cert auth failed: %s", r.text[:200])
        return None
    except Exception:
        log.exception("Cert auth exception")
        return None


def _get_token_with_secret(scope: str) -> Optional[str]:
    try:
        r = requests.post(
            f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            data={
                "client_id": CONFIG["app_id"],
                "client_secret": CONFIG["client_secret"],
                "scope": scope,
                "grant_type": "client_credentials",
            },
            timeout=30,
        )
        if r.status_code == 200:
            return r.json()["access_token"]
        log.error("Secret auth failed: %s", r.text[:200])
        return None
    except Exception:
        log.exception("Secret auth exception")
        return None


def get_sharepoint_token() -> Optional[str]:
    scope = CONFIG["scopes"]["sharepoint"]
    with _token_lock:
        cached = _token_cache.get(scope)
        if cached and cached[1] > time.time() + 60:
            return cached[0]
    tok = _get_token_with_certificate(scope) or _get_token_with_secret(scope)
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


def get_request_digest(site_url: str, token: str) -> str:
    r = requests.post(
        f"{site_url.rstrip('/')}/_api/contextinfo",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json;odata=verbose"},
        timeout=30,
    )
    if r.status_code != 200:
        raise Exception(f"Digest failed {r.status_code}: {r.text[:200]}")
    return r.json()["d"]["GetContextWebInformation"]["FormDigestValue"]


def ensure_user(site_url: str, token: str, digest: str, upn: str) -> Optional[str]:
    """POST ensureuser — returns site-user ID or None."""
    url = f"{site_url.rstrip('/')}/_api/web/ensureuser('{upn}')"
    try:
        r = requests.post(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": digest,
            },
            timeout=30,
        )
        if r.status_code == 200:
            return str(r.json()["d"]["Id"])
        log.debug("ensureuser %s -> %s: %s", upn, r.status_code, r.text[:200])
    except Exception:
        log.exception("ensureuser exception for %s", upn)
    return None


def list_site_groups(site_url: str, token: str) -> List[Dict]:
    """Return all SharePoint groups on the site with their Id and Title."""
    r = requests.get(
        f"{site_url.rstrip('/')}/_api/web/sitegroups?$select=Id,Title",
        headers={"Authorization": f"Bearer {token}", "Accept": "application/json;odata=verbose"},
        timeout=30,
    )
    if r.status_code != 200:
        return []
    return r.json().get("d", {}).get("results", [])


def remove_from_group(site_url: str, token: str, digest: str, group_id: int, user_id: str) -> bool:
    """Remove a single user from a single SharePoint group."""
    url = f"{site_url.rstrip('/')}/_api/web/sitegroups({group_id})/users/removebyid({user_id})"
    try:
        r = requests.post(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": digest,
            },
            timeout=30,
        )
        return r.status_code in (200, 204)
    except Exception:
        log.exception("remove_from_group exception (group=%s user=%s)", group_id, user_id)
        return False


def remove_from_site(site_url: str, token: str, digest: str, user_id: str) -> Tuple[bool, str]:
    """Remove the user from the site's siteusers list. Returns (ok, error_detail)."""
    url = f"{site_url.rstrip('/')}/_api/web/siteusers/removebyid({user_id})"
    try:
        r = requests.post(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": digest,
            },
            timeout=30,
        )
        if r.status_code in (200, 204):
            return True, ""
        # SharePoint often returns 400 with a verbose JSON error
        try:
            err = r.json().get("error", {})
            msg = err.get("message", {}).get("value") or str(err)
            code = err.get("code", "")
            return False, f"{code}: {msg}" if code else msg
        except Exception:
            return False, f"HTTP {r.status_code}: {r.text[:300]}"
    except Exception as e:
        log.exception("remove_from_site exception")
        return False, f"Exception: {e}"


# ------------------------------------------------------------------ #
# CORE: process one (site, user) pair
# ------------------------------------------------------------------ #
def remove_one(
    site_url: str,
    user_upn: str,
    token: str,
    dry_run: bool,
) -> Dict[str, str]:
    """
    Returns dict: status, detail, groups_removed, user_id
    status: REMOVED | NOT_FOUND | FAILED | DRY_RUN
    """
    try:
        digest = get_request_digest(site_url, token)
    except Exception as e:
        return {"status": "FAILED", "detail": f"Digest failed: {e}", "groups_removed": "", "user_id": ""}

    # Resolve user on the site
    user_id = ensure_user(site_url, token, digest, user_upn)
    if not user_id:
        return {
            "status": "NOT_FOUND",
            "detail": "User not resolvable on site (ensureuser returned no ID)",
            "groups_removed": "",
            "user_id": "",
        }

    # Enumerate groups the user belongs to
    groups_removed = 0
    groups_to_remove: List[Tuple[int, str]] = []
    try:
        site_groups = list_site_groups(site_url, token)
        for g in site_groups:
            gid = g.get("Id")
            gtitle = g.get("Title", "")
            if not gid:
                continue
            # Check membership quickly
            check = requests.get(
                f"{site_url.rstrip('/')}/_api/web/sitegroups({gid})/users?$filter=Id eq {user_id}",
                headers={"Authorization": f"Bearer {token}", "Accept": "application/json;odata=verbose"},
                timeout=20,
            )
            if check.status_code == 200 and check.json().get("d", {}).get("results"):
                groups_to_remove.append((gid, gtitle))
    except Exception:
        log.exception("Failed to enumerate groups for %s on %s", user_upn, site_url)

    if dry_run:
        return {
            "status": "DRY_RUN",
            "detail": (
                f"Would remove from site (user_id={user_id}); "
                f"{len(groups_to_remove)} group(s) would be cleared first"
            ),
            "groups_removed": str(len(groups_to_remove)),
            "user_id": user_id,
        }

    # Step 1: remove from all groups
    for gid, gtitle in groups_to_remove:
        if remove_from_group(site_url, token, digest, gid, user_id):
            groups_removed += 1
        else:
            log.warning("Could not remove user %s from group %s (%s)", user_upn, gid, gtitle)

    # Step 2: remove from site
    ok, err = remove_from_site(site_url, token, digest, user_id)
    if ok:
        return {
            "status": "REMOVED",
            "detail": f"Removed from site (groups cleared: {groups_removed})",
            "groups_removed": str(groups_removed),
            "user_id": user_id,
        }

    return {
        "status": "FAILED",
        "detail": (
            f"{err} | Groups removed before failure: {groups_removed}. "
            "Note: SharePoint blocks removal if the user authored content, "
            "has unique permissions, or is a site collection admin — clear those first."
        ),
        "groups_removed": str(groups_removed),
        "user_id": user_id,
    }


# ------------------------------------------------------------------ #
# CSV INPUT
# ------------------------------------------------------------------ #
def read_sites_csv(path: str) -> List[Dict[str, str]]:
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    out = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError("Sites CSV has no header row")
        url_cols = ["Site URL", "Site", "site_url", "site", "URL", "url", "Url"]
        name_cols = ["Site Name", "Name", "site_name", "name"]
        url_col = next((c for c in url_cols if c in reader.fieldnames), None)
        if not url_col:
            raise ValueError(f"Sites CSV missing URL column. Found: {reader.fieldnames}")
        name_col = next((c for c in name_cols if c in reader.fieldnames), None)
        for row in reader:
            raw_url = (row.get(url_col) or "").strip()
            if not raw_url:
                continue
            raw_name = (row.get(name_col) or "").strip() if name_col else raw_url
            out.append({"input": raw_url, "name": raw_name})
    return out


def read_users_csv(path: str) -> List[str]:
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    out: List[str] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            raise ValueError("Users CSV has no header row")
        upn_cols = ["User UPN", "UPN", "upn", "User", "user", "Email", "email",
                    "UserPrincipalName", "userPrincipalName"]
        upn_col = next((c for c in upn_cols if c in reader.fieldnames), None)
        if not upn_col and reader.fieldnames:
            upn_col = reader.fieldnames[0]
            log.warning("Users CSV: no UPN column recognised, using first column '%s'", upn_col)
        if not upn_col:
            raise ValueError(f"Users CSV missing UPN column. Found: {reader.fieldnames}")
        for row in reader:
            upn = (row.get(upn_col) or "").strip()
            if not upn:
                continue
            if "@" not in upn:
                log.warning("Skipping malformed UPN: %r", upn)
                continue
            out.append(upn)
    return out


# ------------------------------------------------------------------ #
# MAIN
# ------------------------------------------------------------------ #
def main():
    ap = argparse.ArgumentParser(description="Remove users from SharePoint sites (CSV in / CSV out).")
    ap.add_argument("--sites", required=True, help="CSV file with site URLs")
    ap.add_argument("--users", required=True, help="CSV file with user UPNs")
    ap.add_argument("--out", default=None, help="Output CSV path")
    ap.add_argument("--workers", type=int, default=5, help="Parallel threads (default: 5)")
    ap.add_argument("--dry-run", action="store_true", help="Preview only, no changes")
    args = ap.parse_args()

    print("=" * 78)
    print("  Remove Users from SharePoint Sites")
    print("=" * 78)

    try:
        sites = read_sites_csv(args.sites)
    except Exception as e:
        print(f"❌ Failed to read sites CSV: {e}")
        sys.exit(1)
    try:
        users = read_users_csv(args.users)
    except Exception as e:
        print(f"❌ Failed to read users CSV: {e}")
        sys.exit(1)

    if not sites:
        print("❌ No sites in input")
        sys.exit(1)
    if not users:
        print("❌ No users in input")
        sys.exit(1)

    print(f"\n📄 Sites:   {len(sites)}")
    print(f"📄 Users:   {len(users)}")
    print(f"🔢 Pairs:   {len(sites) * len(users)}")
    print(f"🧵 Workers: {args.workers}")
    if args.dry_run:
        print("🟡 DRY-RUN — no changes will be made")
    print()

    token = get_sharepoint_token()
    if not token:
        print("❌ Could not acquire SharePoint token")
        sys.exit(1)
    print("🔑 Auth OK\n")

    tasks = []
    for s in sites:
        site_url = normalize_site_url(s["input"])
        for upn in users:
            tasks.append({"site_url": site_url, "site_name": s["name"], "upn": upn})

    results: List[Dict] = []
    start = time.time()
    done = 0

    with ThreadPoolExecutor(max_workers=max(1, args.workers)) as ex:
        future_map = {
            ex.submit(remove_one, t["site_url"], t["upn"], token, args.dry_run): t
            for t in tasks
        }
        for fut in as_completed(future_map):
            t = future_map[fut]
            try:
                r = fut.result()
            except Exception as e:
                r = {"status": "FAILED", "detail": f"Unhandled: {e}", "groups_removed": "", "user_id": ""}
            done += 1
            results.append({
                "Site URL": t["site_url"],
                "Site Name": t["site_name"],
                "User": t["upn"],
                "Status": r["status"],
                "Detail": r["detail"],
                "Groups Removed": r["groups_removed"],
                "Site User ID": r["user_id"],
            })
            icon = {"REMOVED": "✅", "NOT_FOUND": "🔎", "FAILED": "❌", "DRY_RUN": "🟡"}.get(r["status"], "•")
            print(f"[{done}/{len(tasks)}] {icon} {t['upn']:<40} @ {t['site_url']:<55} → {r['status']}")

    elapsed = time.time() - start
    results.sort(key=lambda x: (x["Site URL"].lower(), x["User"].lower()))

    out_path = args.out or f"remove_users_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    with open(out_path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=[
            "Site URL", "Site Name", "User", "Status", "Detail", "Groups Removed", "Site User ID"
        ])
        w.writeheader()
        for r in results:
            w.writerow(r)

    counts: Dict[str, int] = {}
    for r in results:
        counts[r["Status"]] = counts.get(r["Status"], 0) + 1

    print("\n" + "=" * 78)
    print("  SUMMARY")
    print("=" * 78)
    print(f"  Total pairs : {len(results)}")
    print(f"  Time        : {elapsed:.1f}s\n")
    for k in ("REMOVED", "NOT_FOUND", "FAILED", "DRY_RUN"):
        if k in counts:
            print(f"  {k:<12}: {counts[k]}")
    print(f"\n  📄 Report: {out_path}")
    print("=" * 78)


if __name__ == "__main__":
    main()