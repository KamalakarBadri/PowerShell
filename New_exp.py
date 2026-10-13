#!/usr/bin/env python3
"""
PUID Mismatch Repair for "Not Found -> Enabled" OneDrive owners.

Flow:
  1. Read master report.
  2. Find rows where Owner Status in (Not Found, Disabled, Deleted, ...).
  3. Re-check owner in Graph -> if NOW Enabled, add to File1.csv.
  4. For each File1 row (repair_status != success):
       - Compare site owner NameId vs reference-site NameId
       - Mismatch  -> File2.csv
       - Match     -> File3.csv
  5. If run_mode = apply, repair each File2 row -> File4 (success) / File5 (failed).
  6. Write PUID status back to master report.

Outputs (under reports/):
  File1.csv  -- "Not Found -> Enabled" users (the input list)
  File2.csv  -- mismatch found
  File3.csv  -- mismatch not found
  File4.csv  -- repair success (appended)
  File5.csv  -- repair failed  (appended)
  run.log
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
from datetime import datetime, timezone
from pathlib import Path
from threading import Lock
from typing import Any, Dict, List, Optional, Tuple

import requests
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate


# ============================================================================
# CONFIG
# ============================================================================

DEFAULT_CONFIG: Dict[str, Any] = {
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "client_secret": "REPLACE_ME",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    "scopes": {
        "graph": "https://graph.microsoft.com/.default",
        "sharepoint": "https://geekbyteonline.sharepoint.com/.default",
    },

    # Master report (produced by your existing master script)
    "master_report_path": "geekbyteonline_onedrive_master_report.csv",

    # PUID repair
    "repair_account": "edit@geekbyte.online",
    "new_id_site_url": "https://geekbyteonline.sharepoint.com/sites/2DayRetention",
    "run_mode": "report_only",              # report_only | apply
    "report_root": "reports",
    "max_workers": 5,
    "sleep_after_remove_seconds": 2,
    "api_throttle_seconds": 0.25,
    "readded_user_site_admin": True,
    "cleanup_reference_site_user": True,
    "request_timeout_seconds": 60,
}

FILE1 = "File1.csv"   # Not Found -> Enabled users (PUID workflow input)
FILE2 = "File2.csv"   # Mismatch found
FILE3 = "File3.csv"   # Mismatch not found
FILE4 = "File4.csv"   # Repair success (append)
FILE5 = "File5.csv"   # Repair failed  (append)

PROBLEM_STATUSES = {"not found", "disabled", "deleted", "not_found"}


# ============================================================================
# CONFIG / LOGGING
# ============================================================================

def load_config(path: Optional[str]) -> Dict[str, Any]:
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))
    if path and os.path.exists(path):
        with open(path, "r", encoding="utf-8") as fh:
            _merge(cfg, json.load(fh))
    return cfg


def _merge(target: Dict[str, Any], source: Dict[str, Any]) -> None:
    for k, v in source.items():
        if isinstance(v, dict) and isinstance(target.get(k), dict):
            _merge(target[k], v)
        else:
            target[k] = v


def setup_logger(log_dir: Path) -> logging.Logger:
    log_dir.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("puid_notfound_enabled")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    fh = logging.FileHandler(log_dir / "run.log", encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)
    return logger


# ============================================================================
# UTILS
# ============================================================================

def now_iso() -> str:
    return datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")


def safe_str(v: Any, d: str = "") -> str:
    return d if v is None else str(v).strip()


def read_csv(path: Path) -> List[Dict[str, str]]:
    if not path.exists():
        return []
    with open(path, "r", encoding="utf-8", newline="") as fh:
        return list(csv.DictReader(fh))


def write_csv(path: Path, rows: List[Dict[str, Any]], fields: List[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def append_csv(path: Path, rows: List[Dict[str, Any]], fields: List[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    exists = path.exists()
    with open(path, "a", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, extrasaction="ignore")
        if not exists:
            w.writeheader()
        w.writerows(rows)


# ============================================================================
# MICROSOFT 365 CLIENT
# ============================================================================

class MS365Client:
    def __init__(self, cfg: Dict[str, Any], log: logging.Logger):
        self.cfg = cfg
        self.log = log
        self.timeout = cfg.get("request_timeout_seconds", 60)
        self.throttle = cfg.get("api_throttle_seconds", 0.25)
        self._tokens: Dict[str, str] = {}
        self._lock = Lock()

    def _wait(self) -> None:
        if self.throttle and self.throttle > 0:
            time.sleep(self.throttle)

    # ---------- tokens ----------
    def token(self, key: str) -> str:
        scope = self.cfg["scopes"][key]
        if scope in self._tokens:
            return self._tokens[scope]
        t = self._cert_token(scope) or self._secret_token(scope)
        if not t:
            raise RuntimeError(f"Failed to obtain token for {scope}")
        self._tokens[scope] = t
        return t

    def _cert_token(self, scope: str) -> Optional[str]:
        try:
            cp, kp = self.cfg["certificate_path"], self.cfg["private_key_path"]
            if not (os.path.exists(cp) and os.path.exists(kp)):
                return None
            with open(cp, "rb") as f:
                cert = load_pem_x509_certificate(f.read(), default_backend())
            with open(kp, "rb") as f:
                key = load_pem_private_key(f.read(), password=None, backend=default_backend())
            now = int(time.time())
            header = {
                "alg": "RS256", "typ": "JWT",
                "x5t": base64.urlsafe_b64encode(
                    cert.fingerprint(hashes.SHA1())
                ).decode().rstrip("="),
            }
            payload = {
                "aud": f"https://login.microsoftonline.com/{self.cfg['tenant_id']}/oauth2/v2.0/token",
                "exp": now + 300, "iss": self.cfg["app_id"],
                "jti": str(uuid.uuid4()), "nbf": now, "sub": self.cfg["app_id"],
            }
            eh = base64.urlsafe_b64encode(json.dumps(header).encode()).decode().rstrip("=")
            ep = base64.urlsafe_b64encode(json.dumps(payload).encode()).decode().rstrip("=")
            unsigned = f"{eh}.{ep}"
            sig = key.sign(unsigned.encode(), padding.PKCS1v15(), hashes.SHA256())
            assertion = f"{unsigned}.{base64.urlsafe_b64encode(sig).decode().rstrip('=')}"
            url = f"https://login.microsoftonline.com/{self.cfg['tenant_id']}/oauth2/v2.0/token"
            self._wait()
            r = requests.post(url, data={
                "client_id": self.cfg["app_id"],
                "client_assertion": assertion,
                "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                "scope": scope, "grant_type": "client_credentials",
            }, timeout=self.timeout)
            if r.status_code == 200:
                return r.json()["access_token"]
            self.log.warning("Cert auth failed: %s", r.text[:300])
            return None
        except Exception:
            self.log.exception("Cert auth error")
            return None

    def _secret_token(self, scope: str) -> Optional[str]:
        try:
            url = f"https://login.microsoftonline.com/{self.cfg['tenant_id']}/oauth2/v2.0/token"
            self._wait()
            r = requests.post(url, data={
                "client_id": self.cfg["app_id"],
                "client_secret": self.cfg["client_secret"],
                "scope": scope, "grant_type": "client_credentials",
            }, timeout=self.timeout)
            if r.status_code == 200:
                return r.json()["access_token"]
            self.log.warning("Secret auth failed: %s", r.text[:300])
            return None
        except Exception:
            self.log.exception("Secret auth error")
            return None

    # ---------- Graph: is this user enabled now? ----------
    def graph_user_enabled(self, email: str) -> Optional[Dict[str, Any]]:
        """
        Returns:
          {"enabled": True/False, "upn": ..., "mail": ..., "display_name": ..., "id": ...}
          or None if user not found.
        """
        if not email:
            return None
        url = (f"https://graph.microsoft.com/v1.0/users?$filter=mail eq "
               f"'{requests.utils.quote(email)}'"
               f"&$select=id,userPrincipalName,mail,displayName,accountEnabled")
        self._wait()
        r = requests.get(url, headers={
            "Authorization": f"Bearer {self.token('graph')}",
            "Accept": "application/json",
        }, timeout=self.timeout)
        if r.status_code == 404:
            return None
        r.raise_for_status()
        vals = r.json().get("value", [])
        if not vals:
            return None
        u = vals[0]
        return {
            "id": u.get("id", ""),
            "upn": u.get("userPrincipalName", ""),
            "mail": u.get("mail", email),
            "display_name": u.get("displayName", ""),
            "enabled": bool(u.get("accountEnabled", False)),
        }

    # ---------- SharePoint site ops ----------
    def _sp_headers(self, token: str) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
        }

    def get_digest(self, site_url: str, token: str) -> str:
        self._wait()
        r = requests.post(f"{site_url.rstrip('/')}/_api/contextinfo",
                          headers=self._sp_headers(token), timeout=self.timeout)
        r.raise_for_status()
        return r.json()["d"]["GetContextWebInformation"]["FormDigestValue"]

    def get_site_owner(self, site_url: str) -> Dict[str, Optional[str]]:
        token = self.token("sharepoint")
        self._wait()
        r = requests.get(
            f"{site_url.rstrip('/')}/_api/site/owner",
            headers={"Authorization": f"Bearer {token}",
                     "Accept": "application/json;odata=verbose"},
            timeout=self.timeout,
        )
        r.raise_for_status()
        d = r.json().get("d", {})
        upn = d.get("UserPrincipalName") or d.get("Email")
        return {
            "user_id": str(d.get("Id")) if d.get("Id") is not None else None,
            "owner_upn": upn.lower() if upn else None,
            "current_nameid": (d.get("UserId") or {}).get("NameId"),
            "email": d.get("Email"),
            "title": d.get("Title"),
        }

    def ensure_user(self, site_url: str, token: str, digest: str, upn: str) -> Dict[str, Any]:
        self._wait()
        r = requests.post(f"{site_url.rstrip('/')}/_api/web/ensureuser",
                          headers={**self._sp_headers(token), "X-RequestDigest": digest},
                          json={"logonName": upn}, timeout=self.timeout)
        r.raise_for_status()
        return r.json()["d"]

    def set_admin(self, site_url: str, token: str, digest: str, uid: str, admin: bool) -> None:
        self._wait()
        r = requests.post(f"{site_url.rstrip('/')}/_api/web/getuserbyid({uid})",
                          headers={**self._sp_headers(token), "X-RequestDigest": digest,
                                   "X-HTTP-Method": "MERGE", "IF-MATCH": "*"},
                          json={"__metadata": {"type": "SP.User"}, "IsSiteAdmin": admin},
                          timeout=self.timeout)
        if r.status_code not in (200, 204):
            raise RuntimeError(f"set_admin failed {uid}: {r.text[:300]}")

    def remove_user(self, site_url: str, token: str, digest: str, uid: str) -> None:
        self._wait()
        r = requests.post(f"{site_url.rstrip('/')}/_api/web/siteusers/removebyid({uid})",
                          headers={**self._sp_headers(token), "X-RequestDigest": digest},
                          timeout=self.timeout)
        if r.status_code not in (200, 204):
            raise RuntimeError(f"remove_user failed {uid}: {r.text[:300]}")

    def get_reference_nameid(self, upn: str) -> Optional[str]:
        token = self.token("sharepoint")
        ref = self.cfg["new_id_site_url"]
        digest = self.get_digest(ref, token)
        ensured = self.ensure_user(ref, token, digest, upn)
        ref_uid = str(ensured.get("Id")) if ensured.get("Id") is not None else None
        nameid = (ensured.get("UserId") or {}).get("NameId")
        if self.cfg.get("cleanup_reference_site_user", True) and ref_uid:
            try:
                self.remove_user(ref, token, digest, ref_uid)
            except Exception as exc:
                self.log.warning("Reference cleanup failed: %s", exc)
        return nameid


# ============================================================================
# FILE SCHEMAS
# ============================================================================

FILE1_FIELDS = [
    "site_url", "site_id", "site_title", "site_created",
    "owner_email", "owner_upn", "owner_display_name",
    "owner_status", "account_enabled", "graph_user_id",
    "mismatch_status", "mismatch_checked_at",
    "repair_status", "repair_attempted_at", "repair_message",
    "current_nameid", "reference_nameid", "verified_nameid",
    "resolved", "last_updated", "error",
]

MISMATCH_FIELDS = [
    "site_url", "site_id", "site_title", "site_created",
    "owner_upn", "owner_email", "owner_display_name",
    "current_nameid", "reference_nameid", "nameid_match",
    "checked_at", "error",
]

REPAIR_FIELDS = [
    "site_url", "site_id", "owner_upn",
    "old_nameid", "new_nameid", "reference_nameid", "verified_nameid",
    "verified_match", "repair_account_user_id", "readded_user_id",
    "resolved", "status", "message", "repaired_at", "error",
]


# ============================================================================
# STEP 1 — BUILD FILE1 FROM MASTER (Not Found -> Enabled only)
# ============================================================================

def build_file1_notfound_enabled(
    client: MS365Client,
    master_rows: List[Dict[str, str]],
    existing_file1: Dict[str, Dict[str, str]],
    max_workers: int,
    log: logging.Logger,
) -> List[Dict[str, Any]]:
    """
    For each master row whose owner was Not Found/Disabled/Deleted,
    re-check Graph. If the user is NOW Enabled -> include in File1.
    Otherwise keep existing File1 state (for retry) but don't add new rows.
    """
    ts = now_iso()
    workflow = {
        "mismatch_status", "mismatch_checked_at",
        "repair_status", "repair_attempted_at", "repair_message",
        "current_nameid", "reference_nameid", "verified_nameid", "resolved",
    }

    # 1) Filter master rows to problem candidates
    candidates: List[Dict[str, str]] = []
    for r in master_rows:
        owner_status = safe_str(r.get("Owner Status", "")).lower()
        owner_exists = safe_str(r.get("Owner Exists", "")).lower()
        user_status = safe_str(r.get("User Account Status", "")).lower()
        if (owner_exists == "no"
                or owner_status in PROBLEM_STATUSES
                or user_status in PROBLEM_STATUSES):
            candidates.append(r)

    log.info("Problem-owner candidates from master: %d", len(candidates))

    # 2) Re-check in Graph in parallel
    def check(row: Dict[str, str]) -> Optional[Dict[str, Any]]:
        email = safe_str(row.get("Owner Email", "") or row.get("User Email", ""))
        if not email:
            return None
        try:
            u = client.graph_user_enabled(email)
        except Exception as e:
            log.warning("Graph lookup failed %s: %s", email, e)
            return None
        if not u or not u.get("enabled"):
            return None
        old_status = safe_str(row.get("Owner Status", "")) or \
                     safe_str(row.get("User Account Status", "")) or "Not Found"
        return {
            "site_url": safe_str(row.get("Site URL", "")),
            "site_id": safe_str(row.get("Site ID", "")),
            "site_title": safe_str(row.get("Title", "")),
            "site_created": safe_str(row.get("Created On", "")),
            "owner_email": email,
            "owner_upn": u.get("upn", "") or safe_str(row.get("User UPN", "")),
            "owner_display_name": u.get("display_name", "") or
                                  safe_str(row.get("User Display Name", "")),
            "owner_status": "not_found",              # was problematic
            "new_status": "enabled",                  # now enabled
            "old_master_status": old_status,
            "account_enabled": "Yes",
            "graph_user_id": u.get("id", ""),
            "last_updated": ts,
        }

    fresh_problems: List[Dict[str, Any]] = []
    if candidates:
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futs = {ex.submit(check, r): r for r in candidates}
            for fut in as_completed(futs):
                try:
                    rec = fut.result()
                    if rec:
                        fresh_problems.append(rec)
                except Exception:
                    log.exception("Owner check failed")

    log.info("Not Found -> Enabled users found: %d", len(fresh_problems))

    # 3) Merge with existing File1 (state-machine)
    merged: Dict[str, Dict[str, Any]] = {}
    seen_urls = set()
    for rec in fresh_problems:
        url = rec["site_url"]
        seen_urls.add(url)
        old = existing_file1.get(url)
        if old:
            upd = dict(old)
            for k, v in rec.items():
                if k in workflow or k == "new_status":
                    continue
                upd[k] = v
            upd["last_updated"] = ts
            # If owner_status changed from what it was, reset workflow (unless success)
            if (old.get("owner_status") != rec["owner_status"]
                    and (old.get("repair_status") or "").lower() != "success"):
                upd["mismatch_status"] = "pending"
                upd["mismatch_checked_at"] = ""
                upd["repair_status"] = "pending"
                upd["repair_attempted_at"] = ""
                upd["repair_message"] = ""
                upd["resolved"] = "no"
            merged[url] = upd
        else:
            rec.update({
                "mismatch_status": "pending", "mismatch_checked_at": "",
                "repair_status": "pending", "repair_attempted_at": "",
                "repair_message": "", "current_nameid": "",
                "reference_nameid": "", "verified_nameid": "",
                "resolved": "no", "error": "",
            })
            merged[url] = rec

    # Preserve existing File1 rows that are still not success
    # (so retries happen even if Graph check was skipped/failed this run)
    for url, old in existing_file1.items():
        if url in merged:
            continue
        if (old.get("repair_status") or "").lower() == "success":
            merged[url] = dict(old)
            continue
        # Keep rows that were "not found/disabled" but not re-checked
        merged[url] = dict(old)

    return list(merged.values())


# ============================================================================
# STEP 2 — MISMATCH CHECK
# ============================================================================

def check_mismatch(client: MS365Client, row: Dict[str, str],
                   log: logging.Logger) -> Dict[str, Any]:
    rec = {
        "site_url": row.get("site_url", ""),
        "site_id": row.get("site_id", ""),
        "site_title": row.get("site_title", ""),
        "site_created": row.get("site_created", ""),
        "owner_upn": row.get("owner_upn", ""),
        "owner_email": row.get("owner_email", ""),
        "owner_display_name": row.get("owner_display_name", ""),
        "current_nameid": "", "reference_nameid": "",
        "nameid_match": False, "checked_at": now_iso(), "error": "",
    }
    try:
        upn = row.get("owner_upn", "").strip()
        if not upn:
            rec["error"] = "No owner UPN"
            return rec
        current = client.get_site_owner(rec["site_url"]).get("current_nameid") or ""
        reference = client.get_reference_nameid(upn) or ""
        rec["current_nameid"] = current
        rec["reference_nameid"] = reference
        rec["nameid_match"] = bool(current and reference and current == reference)
    except Exception as exc:
        rec["error"] = str(exc)
        log.exception("Mismatch check failed: %s", rec["site_url"])
    return rec


def run_mismatch(client: MS365Client, file1_rows: List[Dict[str, Any]],
                 max_workers: int, log: logging.Logger
                 ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    candidates = [
        r for r in file1_rows
        if (r.get("repair_status") or "").lower() != "success"
        and (r.get("owner_status") or "").lower() in PROBLEM_STATUSES
    ]
    log.info("Mismatch candidates: %d", len(candidates))

    found, not_found = [], []
    if not candidates:
        return found, not_found

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = {ex.submit(check_mismatch, client, r, log): r for r in candidates}
        for fut in as_completed(futs):
            row = futs[fut]
            try:
                rec = fut.result()
                row["mismatch_checked_at"] = rec["checked_at"]
                row["current_nameid"] = rec["current_nameid"]
                row["reference_nameid"] = rec["reference_nameid"]
                row["last_updated"] = now_iso()
                if rec["error"]:
                    row["mismatch_status"] = "error"
                    row["error"] = rec["error"]
                    log.warning("Mismatch error: %s -> %s", rec["site_url"], rec["error"])
                elif rec["nameid_match"]:
                    row["mismatch_status"] = "not_found"
                    not_found.append(rec)
                    log.info("MATCH   : %s", rec["site_url"])
                else:
                    row["mismatch_status"] = "found"
                    found.append(rec)
                    log.info("MISMATCH: %s", rec["site_url"])
            except Exception:
                log.exception("Future failed: %s", row.get("site_url"))
    return found, not_found


# ============================================================================
# STEP 3 — REPAIR
# ============================================================================

def repair_one(client: MS365Client, row: Dict[str, str],
               log: logging.Logger) -> Dict[str, Any]:
    rec = {
        "site_url": row.get("site_url", ""),
        "site_id": row.get("site_id", ""),
        "owner_upn": row.get("owner_upn", ""),
        "old_nameid": row.get("current_nameid", ""),
        "reference_nameid": row.get("reference_nameid", ""),
        "new_nameid": "", "verified_nameid": "",
        "verified_match": False,
        "repair_account_user_id": "", "readded_user_id": "",
        "resolved": "no", "status": "", "message": "",
        "repaired_at": now_iso(), "error": "",
    }
    site_url = rec["site_url"]
    upn = rec["owner_upn"]
    try:
        token = client.token("sharepoint")
        digest = client.get_digest(site_url, token)

        owner = client.get_site_owner(site_url)
        current_uid = owner.get("user_id")
        if not current_uid:
            raise RuntimeError("Owner user ID missing on site")

        repair_user = client.ensure_user(site_url, token, digest, client.cfg["repair_account"])
        repair_uid = str(repair_user.get("Id"))
        rec["repair_account_user_id"] = repair_uid
        client.set_admin(site_url, token, digest, repair_uid, True)

        client.set_admin(site_url, token, digest, current_uid, False)
        client.remove_user(site_url, token, digest, current_uid)
        time.sleep(client.cfg.get("sleep_after_remove_seconds", 2))

        readded = client.ensure_user(site_url, token, digest, upn)
        readded_uid = str(readded.get("Id"))
        rec["readded_user_id"] = readded_uid

        if client.cfg.get("readded_user_site_admin", True):
            client.set_admin(site_url, token, digest, readded_uid, True)

        try:
            client.set_admin(site_url, token, digest, repair_uid, False)
            client.remove_user(site_url, token, digest, repair_uid)
        except Exception as exc:
            log.warning("Repair account cleanup failed: %s", exc)

        verified = client.get_site_owner(site_url).get("current_nameid") or ""
        rec["verified_nameid"] = verified
        rec["new_nameid"] = verified
        rec["verified_match"] = bool(rec["reference_nameid"]) and verified == rec["reference_nameid"]

        if rec["verified_match"]:
            rec["status"] = "success"
            rec["resolved"] = "yes"
            rec["message"] = "NameId verified after re-add"
        else:
            rec["status"] = "failed"
            rec["resolved"] = "no"
            rec["message"] = "NameId not verified after re-add"
    except Exception as exc:
        rec["status"] = "error"
        rec["resolved"] = "error"
        rec["error"] = str(exc)
        rec["message"] = str(exc)
        log.exception("Repair failed: %s", site_url)
    return rec


def run_repairs(client: MS365Client, mismatch_rows: List[Dict[str, str]],
                max_workers: int, log: logging.Logger
                ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    success, failed = [], []
    if not mismatch_rows:
        return success, failed

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = {ex.submit(repair_one, client, r, log): r for r in mismatch_rows}
        for fut in as_completed(futs):
            row = futs[fut]
            try:
                rec = fut.result()
                if rec["status"] == "success":
                    success.append(rec)
                    log.info("REPAIR OK  : %s", rec["site_url"])
                else:
                    failed.append(rec)
                    log.warning("REPAIR FAIL: %s - %s", rec["site_url"], rec["message"])
            except Exception:
                log.exception("Repair future failed: %s", row.get("site_url"))
    return success, failed


# ============================================================================
# STEP 4 — WRITE PUID STATUS BACK TO MASTER
# ============================================================================

PUID_COLS = [
    "PUID Repair Status", "PUID Repair Attempted At", "PUID Repair Message",
    "PUID Current NameId", "PUID Reference NameId", "PUID Verified NameId",
]


def update_master_with_puid(master_path: Path, file1_rows: List[Dict[str, Any]],
                            log: logging.Logger) -> None:
    master_rows = read_csv(master_path)
    if not master_rows:
        log.warning("Master report missing/empty: %s", master_path)
        return

    by_url = {r["site_url"]: r for r in file1_rows if r.get("site_url")}
    by_id = {r.get("site_id"): r for r in file1_rows if r.get("site_id")}

    updated = 0
    for m in master_rows:
        sid = safe_str(m.get("Site ID", ""))
        surl = safe_str(m.get("Site URL", ""))
        f = by_id.get(sid) or by_url.get(surl)
        if f:
            m["PUID Repair Status"] = f.get("repair_status") or f.get("mismatch_status") or ""
            m["PUID Repair Attempted At"] = f.get("repair_attempted_at", "")
            m["PUID Repair Message"] = f.get("repair_message", "")
            m["PUID Current NameId"] = f.get("current_nameid", "")
            m["PUID Reference NameId"] = f.get("reference_nameid", "")
            m["PUID Verified NameId"] = f.get("verified_nameid", "")
            updated += 1
        for c in PUID_COLS:
            m.setdefault(c, "")

    base_fields = list(master_rows[0].keys())
    for c in PUID_COLS:
        if c not in base_fields:
            base_fields.append(c)

    write_csv(master_path, master_rows, base_fields)
    log.info("Master report updated with PUID results: %d row(s)", updated)


# ============================================================================
# MAIN
# ============================================================================

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="PUID repair for 'Not Found -> Enabled' OneDrive owners")
    p.add_argument("--config", default=None)
    p.add_argument("--master-report", default=None, help="Override master CSV path")
    p.add_argument("--skip-mismatch", action="store_true",
                   help="Use existing File2.csv (skip mismatch recheck)")
    p.add_argument("--max-workers", type=int, default=None)
    return p.parse_args()


def main() -> int:
    args = parse_args()
    cfg = load_config(args.config)
    apply_changes = (cfg.get("run_mode") or "report_only").lower() == "apply"

    report_dir = Path(cfg["report_root"])
    log = setup_logger(report_dir)

    master_path = Path(args.master_report or cfg["master_report_path"])
    if not master_path.exists():
        log.error("Master report not found: %s", master_path)
        return 1

    file1_path = report_dir / FILE1
    file2_path = report_dir / FILE2
    file3_path = report_dir / FILE3
    file4_path = report_dir / FILE4
    file5_path = report_dir / FILE5

    log.info("=" * 70)
    log.info("PUID Repair — Not Found -> Enabled workflow")
    log.info("Mode: %s", "APPLY" if apply_changes else "REPORT ONLY")
    log.info("Master: %s", master_path)
    log.info("=" * 70)

    client = MS365Client(cfg, log)
    max_workers = args.max_workers or cfg.get("max_workers", 5)

    try:
        # ---- STEP 1: Build File1 from master ----
        master_rows = read_csv(master_path)
        existing_file1 = {r.get("site_url", ""): r for r in read_csv(file1_path)
                          if r.get("site_url")}

        file1_rows = build_file1_notfound_enabled(
            client, master_rows, existing_file1, max_workers, log
        )
        write_csv(file1_path, file1_rows, FILE1_FIELDS)
        log.info("File1.csv written: %d row(s)", len(file1_rows))

        # ---- STEP 2: Mismatch check ----
        if args.skip_mismatch and file2_path.exists():
            found = read_csv(file2_path)
            log.info("Skipping mismatch check, using File2.csv: %d row(s)", len(found))
        else:
            found, not_found = run_mismatch(client, file1_rows, max_workers, log)
            write_csv(file2_path, found, MISMATCH_FIELDS)
            write_csv(file3_path, not_found, MISMATCH_FIELDS)
            log.info("File2 (mismatch found): %d row(s)", len(found))
            log.info("File3 (mismatch not found): %d row(s)", len(not_found))
            write_csv(file1_path, file1_rows, FILE1_FIELDS)

        # ---- STEP 3: Repair ----
        if not apply_changes:
            log.info("Run mode 'report_only' -> repairs skipped")
        elif not found:
            log.info("No mismatch records -> no repairs")
        else:
            success, failed = run_repairs(client, found, max_workers, log)
            append_csv(file4_path, success, REPAIR_FIELDS)
            append_csv(file5_path, failed, REPAIR_FIELDS)
            log.info("File4 (success): appended %d", len(success))
            log.info("File5 (failed): appended %d", len(failed))

            idx = {r["site_url"]: r for r in file1_rows if r.get("site_url")}
            for rec in success:
                r = idx.get(rec["site_url"])
                if r:
                    r["repair_status"] = "success"
                    r["repair_attempted_at"] = rec["repaired_at"]
                    r["repair_message"] = rec["message"]
                    r["verified_nameid"] = rec["verified_nameid"]
                    r["resolved"] = "yes"
                    r["last_updated"] = now_iso()
            for rec in failed:
                r = idx.get(rec["site_url"])
                if r:
                    r["repair_status"] = "failed"
                    r["repair_attempted_at"] = rec["repaired_at"]
                    r["repair_message"] = rec["message"]
                    r["verified_nameid"] = rec["verified_nameid"]
                    r["resolved"] = rec.get("resolved", "no")
                    r["last_updated"] = now_iso()
            write_csv(file1_path, file1_rows, FILE1_FIELDS)

        # ---- STEP 4: Master report PUID status ----
        update_master_with_puid(master_path, file1_rows, log)

        log.info("=" * 70)
        log.info("DONE")
        log.info("  File1 (Not Found -> Enabled) : %s", file1_path)
        log.info("  File2 (mismatch found)       : %s", file2_path)
        log.info("  File3 (mismatch not found)   : %s", file3_path)
        log.info("  File4 (repair success)       : %s", file4_path)
        log.info("  File5 (repair failed)        : %s", file5_path)
        log.info("=" * 70)
        return 0

    except Exception:
        log.exception("Workflow failed")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
