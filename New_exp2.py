#!/usr/bin/env python3
"""
OneDrive PUID Mismatch Repair — two workflows + date filter + repair_master.

IMPORTANT: This script NEVER modifies geekbyteonline_onedrive_master_report.csv.
           The master is READ-ONLY. Your master_report.py remains the only writer.
           All PUID results live in repair_master.csv and File1..File10.

Workflow A: User status change (Not Found -> Enabled)
  File1 -> File2/File3 -> File4/File5

Workflow B: Manager change
  File6 -> File7/File8 -> File9/File10

All candidates must have a transition timestamp >= --from-date.

Successfully repaired sites are written to repair_master.csv.
"""

import argparse
import base64
import csv
import json
import logging
import os
import re
import sys
import time
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone
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
        "sharepoint": "https://geekbyteonline.sharepoint.com/.default",
    },
    "master_report_path": "geekbyteonline_onedrive_master_report.csv",
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
    "default_days_back": 7,
}

# Workflow A: Not Found -> Enabled
FILE1 = "File1.csv"
FILE2 = "File2.csv"
FILE3 = "File3.csv"
FILE4 = "File4.csv"
FILE5 = "File5.csv"

# Workflow B: Manager Change
FILE6 = "File6.csv"
FILE7 = "File7.csv"
FILE8 = "File8.csv"
FILE9 = "File9.csv"
FILE10 = "File10.csv"

# Repair master
REPAIR_MASTER = "repair_master.csv"


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
    logger = logging.getLogger("puid_repair")
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


def extract_timestamps(history_text: str) -> List[datetime]:
    """
    Extract all timestamps from a change-history string.
    Supports:
      [2026-10-13 09:00:00]
      [2026-10-13T09:00:00Z]
    """
    if not history_text:
        return []
    out: List[datetime] = []
    for m in re.finditer(r"\[([0-9T:\- ]+Z?)\]", history_text):
        raw = m.group(1).strip()
        dt: Optional[datetime] = None
        try:
            if raw.endswith("Z"):
                dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
            else:
                for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"):
                    try:
                        dt = datetime.strptime(raw, fmt)
                        break
                    except ValueError:
                        continue
        except Exception:
            continue
        if dt is None:
            continue
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        out.append(dt)
    return out


def latest_timestamp(history_text: str) -> Optional[datetime]:
    ts = extract_timestamps(history_text)
    return max(ts) if ts else None


def is_within_date_range(history_text: str, from_dt: datetime) -> bool:
    latest = latest_timestamp(history_text)
    if latest is None:
        return False
    return latest >= from_dt


# ============================================================================
# SHAREPOINT CLIENT
# ============================================================================

class SharePointClient:
    def __init__(self, cfg: Dict[str, Any], log: logging.Logger):
        self.cfg = cfg
        self.log = log
        self.timeout = cfg.get("request_timeout_seconds", 60)
        self.throttle = cfg.get("api_throttle_seconds", 0.25)
        self._token: Optional[str] = None
        self._lock = Lock()

    def _wait(self) -> None:
        if self.throttle and self.throttle > 0:
            time.sleep(self.throttle)

    def token(self) -> str:
        with self._lock:
            if not self._token:
                self._token = self._acquire()
            return self._token

    def _acquire(self) -> str:
        scope = self.cfg["scopes"]["sharepoint"]
        assertion = self._assertion()
        url = f"https://login.microsoftonline.com/{self.cfg['tenant_id']}/oauth2/v2.0/token"
        data = {
            "client_id": self.cfg["app_id"],
            "client_assertion": assertion,
            "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            "scope": scope, "grant_type": "client_credentials",
        }
        r = requests.post(url, data=data, timeout=self.timeout)
        if r.status_code == 200:
            return r.json()["access_token"]
        r2 = requests.post(url, data={
            "client_id": self.cfg["app_id"],
            "client_secret": self.cfg["client_secret"],
            "scope": scope, "grant_type": "client_credentials",
        }, timeout=self.timeout)
        r2.raise_for_status()
        return r2.json()["access_token"]

    def _assertion(self) -> str:
        with open(self.cfg["certificate_path"], "rb") as f:
            cert = load_pem_x509_certificate(f.read(), default_backend())
        with open(self.cfg["private_key_path"], "rb") as f:
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
        return f"{unsigned}.{base64.urlsafe_b64encode(sig).decode().rstrip('=')}"

    def _sp_headers(self) -> Dict[str, str]:
        return {
            "Authorization": f"Bearer {self.token()}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
        }

    def get_digest(self, site_url: str) -> str:
        self._wait()
        r = requests.post(f"{site_url.rstrip('/')}/_api/contextinfo",
                          headers=self._sp_headers(), timeout=self.timeout)
        r.raise_for_status()
        return r.json()["d"]["GetContextWebInformation"]["FormDigestValue"]

    def get_site_owner(self, site_url: str) -> Dict[str, Optional[str]]:
        self._wait()
        r = requests.get(
            f"{site_url.rstrip('/')}/_api/site/owner",
            headers={"Authorization": f"Bearer {self.token()}",
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

    def ensure_user(self, site_url: str, digest: str, upn: str) -> Dict[str, Any]:
        self._wait()
        r = requests.post(f"{site_url.rstrip('/')}/_api/web/ensureuser",
                          headers={**self._sp_headers(), "X-RequestDigest": digest},
                          json={"logonName": upn}, timeout=self.timeout)
        r.raise_for_status()
        return r.json()["d"]

    def set_admin(self, site_url: str, digest: str, uid: str, admin: bool) -> None:
        self._wait()
        r = requests.post(f"{site_url.rstrip('/')}/_api/web/getuserbyid({uid})",
                          headers={**self._sp_headers(), "X-RequestDigest": digest,
                                   "X-HTTP-Method": "MERGE", "IF-MATCH": "*"},
                          json={"__metadata": {"type": "SP.User"}, "IsSiteAdmin": admin},
                          timeout=self.timeout)
        if r.status_code not in (200, 204):
            raise RuntimeError(f"set_admin failed {uid}: {r.text[:300]}")

    def remove_user(self, site_url: str, digest: str, uid: str) -> None:
        self._wait()
        r = requests.post(f"{site_url.rstrip('/')}/_api/web/siteusers/removebyid({uid})",
                          headers={**self._sp_headers(), "X-RequestDigest": digest},
                          timeout=self.timeout)
        if r.status_code not in (200, 204):
            raise RuntimeError(f"remove_user failed {uid}: {r.text[:300]}")

    def get_reference_nameid(self, upn: str) -> Optional[str]:
        ref = self.cfg["new_id_site_url"]
        digest = self.get_digest(ref)
        ensured = self.ensure_user(ref, digest, upn)
        ref_uid = str(ensured.get("Id")) if ensured.get("Id") is not None else None
        nameid = (ensured.get("UserId") or {}).get("NameId")
        if self.cfg.get("cleanup_reference_site_user", True) and ref_uid:
            try:
                self.remove_user(ref, digest, ref_uid)
            except Exception as exc:
                self.log.warning("Reference cleanup failed: %s", exc)
        return nameid


# ============================================================================
# FILE SCHEMAS
# ============================================================================

CANDIDATE_FIELDS = [
    "site_url", "site_id", "site_title", "site_created",
    "owner_email", "owner_upn", "owner_display_name",
    "trigger", "trigger_history", "trigger_timestamp",
    "mismatch_status", "mismatch_checked_at",
    "repair_status", "repair_attempted_at", "repair_message",
    "current_nameid", "reference_nameid", "verified_nameid",
    "resolved", "last_updated", "error",
]

MISMATCH_FIELDS = [
    "site_url", "site_id", "site_title", "site_created",
    "owner_upn", "owner_email", "owner_display_name",
    "trigger", "trigger_timestamp",
    "current_nameid", "reference_nameid", "nameid_match",
    "checked_at", "error",
]

REPAIR_FIELDS = [
    "site_url", "site_id", "owner_upn",
    "trigger", "trigger_timestamp",
    "old_nameid", "new_nameid", "reference_nameid", "verified_nameid",
    "verified_match", "repair_account_user_id", "readded_user_id",
    "resolved", "status", "message", "repaired_at", "error",
]

REPAIR_MASTER_FIELDS = [
    "site_url", "site_id", "site_title",
    "owner_email", "owner_upn", "owner_display_name",
    "trigger", "trigger_details", "trigger_timestamp",
    "old_nameid", "new_nameid", "reference_nameid", "verified_nameid",
    "verified_match",
    "repair_status", "repair_message", "repair_attempted_at",
    "first_repaired_at", "last_updated",
]


# ============================================================================
# DETECTORS (with date filter)
# ============================================================================

def is_not_found_to_enabled(row: Dict[str, str], from_dt: datetime) -> Optional[Dict[str, str]]:
    """Workflow A: Not Found -> Enabled, timestamp >= from_dt."""
    history = safe_str(row.get("User Account Status Change History", ""))
    current = safe_str(row.get("User Account Status", "")).lower()
    if current != "enabled" or not history:
        return None

    h = history.lower()
    h = re.sub(r"\s*>>\s*", " >> ", h)
    h = re.sub(r"\s*→\s*", " >> ", h)
    h = re.sub(r"\s+", " ", h)

    if "not found >> enabled" not in h:
        return None

    matches = re.findall(r">>\s*([a-z ]+)", h)
    if matches and matches[-1].strip() != "enabled":
        return None

    if not is_within_date_range(history, from_dt):
        return None

    latest = latest_timestamp(history)
    return {
        "history": history,
        "timestamp": latest.strftime("%Y-%m-%d %H:%M:%S") if latest else "",
    }


def is_manager_changed(row: Dict[str, str], from_dt: datetime) -> Optional[Dict[str, str]]:
    """Workflow B: Manager change, timestamp >= from_dt."""
    history = safe_str(row.get("Manager Change History", ""))
    if not history:
        return None

    h = history.lower()
    h = re.sub(r"\s*>>\s*", " >> ", h)
    h = re.sub(r"\s*→\s*", " >> ", h)

    if " >> " not in h:
        return None

    if not is_within_date_range(history, from_dt):
        return None

    latest = latest_timestamp(history)
    return {
        "history": history,
        "timestamp": latest.strftime("%Y-%m-%d %H:%M:%S") if latest else "",
    }


# ============================================================================
# BUILD CANDIDATE FILES
# ============================================================================

def build_candidates(
    master_rows: List[Dict[str, str]],
    detector,
    trigger_label: str,
    from_dt: datetime,
    existing_candidates: Dict[str, Dict[str, str]],
    log: logging.Logger,
) -> List[Dict[str, Any]]:
    ts = now_iso()
    workflow = {
        "mismatch_status", "mismatch_checked_at",
        "repair_status", "repair_attempted_at", "repair_message",
        "current_nameid", "reference_nameid", "verified_nameid", "resolved",
    }

    candidates: List[Dict[str, Any]] = []
    for r in master_rows:
        hit = detector(r, from_dt)
        if not hit:
            continue
        surl = safe_str(r.get("Site URL", ""))
        prior = existing_candidates.get(surl, {})
        if (prior.get("repair_status") or "").lower() == "success":
            continue
        candidates.append({
            "site_url": surl,
            "site_id": safe_str(r.get("Site ID", "")),
            "site_title": safe_str(r.get("Title", "")),
            "site_created": safe_str(r.get("Created On", "")),
            "owner_email": safe_str(r.get("Owner Email", "") or r.get("User Email", "")),
            "owner_upn": safe_str(r.get("User UPN", "")),
            "owner_display_name": safe_str(r.get("User Display Name", "")),
            "trigger": trigger_label,
            "trigger_history": hit["history"],
            "trigger_timestamp": hit["timestamp"],
            "last_updated": ts,
        })

    log.info("[%s] candidates from master (from %s): %d",
             trigger_label, from_dt.strftime("%Y-%m-%d"), len(candidates))

    merged: Dict[str, Dict[str, Any]] = {}
    seen = set()
    for rec in candidates:
        url = rec["site_url"]
        seen.add(url)
        old = existing_candidates.get(url)
        if old:
            upd = dict(old)
            for k, v in rec.items():
                if k in workflow:
                    continue
                upd[k] = v
            upd["last_updated"] = ts
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

    for url, old in existing_candidates.items():
        if url in seen:
            continue
        if (old.get("repair_status") or "").lower() != "success":
            merged[url] = dict(old)

    return list(merged.values())


# ============================================================================
# MISMATCH CHECK
# ============================================================================

def check_mismatch(client: SharePointClient, row: Dict[str, str],
                   log: logging.Logger) -> Dict[str, Any]:
    rec = {
        "site_url": row.get("site_url", ""),
        "site_id": row.get("site_id", ""),
        "site_title": row.get("site_title", ""),
        "site_created": row.get("site_created", ""),
        "owner_upn": row.get("owner_upn", ""),
        "owner_email": row.get("owner_email", ""),
        "owner_display_name": row.get("owner_display_name", ""),
        "trigger": row.get("trigger", ""),
        "trigger_timestamp": row.get("trigger_timestamp", ""),
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


def run_mismatch(client: SharePointClient, rows: List[Dict[str, Any]],
                 max_workers: int, log: logging.Logger
                 ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    candidates = [
        r for r in rows
        if (r.get("repair_status") or "").lower() != "success"
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
# REPAIR
# ============================================================================

def repair_one(client: SharePointClient, row: Dict[str, str],
               log: logging.Logger) -> Dict[str, Any]:
    rec = {
        "site_url": row.get("site_url", ""),
        "site_id": row.get("site_id", ""),
        "owner_upn": row.get("owner_upn", ""),
        "trigger": row.get("trigger", ""),
        "trigger_timestamp": row.get("trigger_timestamp", ""),
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
        digest = client.get_digest(site_url)

        owner = client.get_site_owner(site_url)
        current_uid = owner.get("user_id")
        if not current_uid:
            raise RuntimeError("Owner user ID missing on site")

        repair_user = client.ensure_user(site_url, digest, client.cfg["repair_account"])
        repair_uid = str(repair_user.get("Id"))
        rec["repair_account_user_id"] = repair_uid
        client.set_admin(site_url, digest, repair_uid, True)

        client.set_admin(site_url, digest, current_uid, False)
        client.remove_user(site_url, digest, current_uid)
        time.sleep(client.cfg.get("sleep_after_remove_seconds", 2))

        readded = client.ensure_user(site_url, digest, upn)
        readded_uid = str(readded.get("Id"))
        rec["readded_user_id"] = readded_uid

        if client.cfg.get("readded_user_site_admin", True):
            client.set_admin(site_url, digest, readded_uid, True)

        try:
            client.set_admin(site_url, digest, repair_uid, False)
            client.remove_user(site_url, digest, repair_uid)
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


def run_repairs(client: SharePointClient, rows: List[Dict[str, str]],
                max_workers: int, log: logging.Logger
                ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    success, failed = [], []
    if not rows:
        return success, failed

    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = {ex.submit(repair_one, client, r, log): r for r in rows}
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
# REPAIR MASTER (File11)
# ============================================================================

def update_repair_master(
    repair_master_path: Path,
    workflow_a_success: List[Dict[str, Any]],
    workflow_b_success: List[Dict[str, Any]],
    master_rows_by_url: Dict[str, Dict[str, str]],
    log: logging.Logger,
) -> None:
    """
    Write/refresh repair_master.csv with successfully repaired sites.
    This file is INDEPENDENT of the master report.
    """
    existing = read_csv(repair_master_path)
    index: Dict[str, Dict[str, Any]] = {
        r.get("site_url", ""): dict(r) for r in existing if r.get("site_url")
    }

    ts = now_iso()

    def add_success(rec: Dict[str, Any], trigger_label: str) -> None:
        url = rec.get("site_url", "")
        if not url:
            return

        master_row = master_rows_by_url.get(url, {})
        prior = index.get(url, {})

        prev_trigger = prior.get("trigger", "")
        if not prev_trigger:
            merged_trigger = trigger_label
        elif trigger_label in prev_trigger.split("|"):
            merged_trigger = prev_trigger
        else:
            merged_trigger = f"{prev_trigger}|{trigger_label}"

        detail_parts = []
        if prior.get("trigger_details"):
            detail_parts.append(prior["trigger_details"])
        detail_parts.append(f"{trigger_label} at {rec.get('repaired_at', '')}")

        index[url] = {
            "site_url": url,
            "site_id": rec.get("site_id", ""),
            "site_title": master_row.get("Title", ""),
            "owner_email": master_row.get("Owner Email", "") or master_row.get("User Email", ""),
            "owner_upn": rec.get("owner_upn", "") or master_row.get("User UPN", ""),
            "owner_display_name": master_row.get("User Display Name", ""),
            "trigger": merged_trigger,
            "trigger_details": " | ".join(p for p in detail_parts if p),
            "trigger_timestamp": rec.get("trigger_timestamp", "") or prior.get("trigger_timestamp", ""),
            "old_nameid": rec.get("old_nameid", ""),
            "new_nameid": rec.get("new_nameid", ""),
            "reference_nameid": rec.get("reference_nameid", ""),
            "verified_nameid": rec.get("verified_nameid", ""),
            "verified_match": str(rec.get("verified_match", False)),
            "repair_status": rec.get("status", "success"),
            "repair_message": rec.get("message", ""),
            "repair_attempted_at": rec.get("repaired_at", ""),
            "first_repaired_at": prior.get("first_repaired_at", "") or rec.get("repaired_at", ""),
            "last_updated": ts,
        }

    for rec in workflow_a_success:
        add_success(rec, "user_status_change")
    for rec in workflow_b_success:
        add_success(rec, "manager_change")

    rows = list(index.values())
    write_csv(repair_master_path, rows, REPAIR_MASTER_FIELDS)
    log.info("repair_master.csv written: %d row(s)", len(rows))


# ============================================================================
# WORKFLOW RUNNER
# ============================================================================

def run_workflow(
    client: SharePointClient,
    report_dir: Path,
    master_rows: List[Dict[str, str]],
    detector,
    trigger_label: str,
    from_dt: datetime,
    candidate_file: str,
    mismatch_found_file: str,
    mismatch_notfound_file: str,
    success_file: str,
    failed_file: str,
    apply_changes: bool,
    max_workers: int,
    skip_mismatch: bool,
    log: logging.Logger,
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    log.info("=" * 70)
    log.info("Workflow: %s (from %s)", trigger_label, from_dt.strftime("%Y-%m-%d"))
    log.info("=" * 70)

    cand_path = report_dir / candidate_file
    f_found_path = report_dir / mismatch_found_file
    f_notfound_path = report_dir / mismatch_notfound_file
    f_success_path = report_dir / success_file
    f_failed_path = report_dir / failed_file

    existing = {r.get("site_url", ""): r for r in read_csv(cand_path) if r.get("site_url")}
    candidates = build_candidates(
        master_rows, detector, trigger_label, from_dt, existing, log
    )
    write_csv(cand_path, candidates, CANDIDATE_FIELDS)
    log.info("%s written: %d row(s)", candidate_file, len(candidates))

    if skip_mismatch and f_found_path.exists():
        found = read_csv(f_found_path)
        log.info("Skipping mismatch, using %s: %d row(s)", mismatch_found_file, len(found))
    else:
        found, not_found = run_mismatch(client, candidates, max_workers, log)
        write_csv(f_found_path, found, MISMATCH_FIELDS)
        write_csv(f_notfound_path, not_found, MISMATCH_FIELDS)
        log.info("%s (mismatch found): %d", mismatch_found_file, len(found))
        log.info("%s (mismatch not found): %d", mismatch_notfound_file, len(not_found))
        write_csv(cand_path, candidates, CANDIDATE_FIELDS)

    success: List[Dict[str, Any]] = []
    if not apply_changes:
        log.info("Run mode 'report_only' -> repairs skipped")
    elif not found:
        log.info("No mismatch records -> no repairs")
    else:
        success, failed = run_repairs(client, found, max_workers, log)
        append_csv(f_success_path, success, REPAIR_FIELDS)
        append_csv(f_failed_path, failed, REPAIR_FIELDS)
        log.info("%s (success): appended %d", success_file, len(success))
        log.info("%s (failed): appended %d", failed_file, len(failed))

        idx = {r["site_url"]: r for r in candidates if r.get("site_url")}
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
        write_csv(cand_path, candidates, CANDIDATE_FIELDS)

    return candidates, success


# ============================================================================
# MAIN
# ============================================================================

def resolve_from_date(args, cfg) -> datetime:
    if args.from_date:
        return datetime.fromisoformat(f"{args.from_date}T00:00:00+00:00")
    days = int(cfg.get("default_days_back", 7))
    return (datetime.now(timezone.utc) - timedelta(days=days)).replace(
        hour=0, minute=0, second=0, microsecond=0
    )


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="PUID repair — date-filtered, two workflows, repair_master output")
    p.add_argument("--config", default=None)
    p.add_argument("--master-report", default=None)
    p.add_argument("--from-date", default=None,
                   help="YYYY-MM-DD. Transitions on/after this date only.")
    p.add_argument("--skip-mismatch", action="store_true")
    p.add_argument("--only", choices=["notfound", "manager", "both"], default="both")
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

    from_dt = resolve_from_date(args, cfg)
    log.info("=" * 70)
    log.info("PUID Repair — starting")
    log.info("Mode: %s", "APPLY" if apply_changes else "REPORT ONLY")
    log.info("Master (read-only): %s", master_path)
    log.info("Date filter (>=): %s", from_dt.strftime("%Y-%m-%d %H:%M:%S"))
    log.info("=" * 70)

    client = SharePointClient(cfg, log)
    max_workers = args.max_workers or cfg.get("max_workers", 5)

    # ---- Read master ONCE. Never write to it. ----
    master_rows = read_csv(master_path)
    master_by_url = {safe_str(r.get("Site URL", "")): r for r in master_rows}

    success_a: List[Dict[str, Any]] = []
    success_b: List[Dict[str, Any]] = []

    try:
        # ---- Workflow A ----
        if args.only in ("notfound", "both"):
            _, success_a = run_workflow(
                client, report_dir, master_rows,
                detector=is_not_found_to_enabled,
                trigger_label="user_status_change",
                from_dt=from_dt,
                candidate_file=FILE1,
                mismatch_found_file=FILE2,
                mismatch_notfound_file=FILE3,
                success_file=FILE4,
                failed_file=FILE5,
                apply_changes=apply_changes,
                max_workers=max_workers,
                skip_mismatch=args.skip_mismatch,
                log=log,
            )

        # ---- Workflow B ----
        if args.only in ("manager", "both"):
            _, success_b = run_workflow(
                client, report_dir, master_rows,
                detector=is_manager_changed,
                trigger_label="manager_change",
                from_dt=from_dt,
                candidate_file=FILE6,
                mismatch_found_file=FILE7,
                mismatch_notfound_file=FILE8,
                success_file=FILE9,
                failed_file=FILE10,
                apply_changes=apply_changes,
                max_workers=max_workers,
                skip_mismatch=args.skip_mismatch,
                log=log,
            )

        # ---- Repair master (independent file, NOT the master report) ----
        update_repair_master(
            report_dir / REPAIR_MASTER,
            workflow_a_success=success_a,
            workflow_b_success=success_b,
            master_rows_by_url=master_by_url,
            log=log,
        )

        # ---- Master report is NOT touched ----
        log.info("Master report left unmodified: %s", master_path)

        log.info("=" * 70)
        log.info("DONE")
        log.info("  Workflow A (User status change):")
        log.info("    File1  : %s", report_dir / FILE1)
        log.info("    File2  : %s", report_dir / FILE2)
        log.info("    File3  : %s", report_dir / FILE3)
        log.info("    File4  : %s", report_dir / FILE4)
        log.info("    File5  : %s", report_dir / FILE5)
        log.info("  Workflow B (Manager change):")
        log.info("    File6  : %s", report_dir / FILE6)
        log.info("    File7  : %s", report_dir / FILE7)
        log.info("    File8  : %s", report_dir / FILE8)
        log.info("    File9  : %s", report_dir / FILE9)
        log.info("    File10 : %s", report_dir / FILE10)
        log.info("  Repair master (independent):")
        log.info("    repair_master.csv : %s", report_dir / REPAIR_MASTER)
        log.info("=" * 70)
        return 0

    except Exception:
        log.exception("Workflow failed")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
