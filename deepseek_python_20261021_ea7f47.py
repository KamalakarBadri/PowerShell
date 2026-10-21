#!/usr/bin/env python3
"""
OneDrive PUID Mismatch Repair — two workflows, date-filtered, cumulative files.

- File1.csv : User status change (Not Found -> Enabled), one row per site_url
- File2.csv : Manager change (any change), one row per site_url

- Date filter: only transitions whose latest timestamp >= cutoff date.
- No enabled-state check. Manager workflow processes any change in the window.
- Master report is READ-ONLY. Nothing is ever written to it.

Usage:
  python puid_repair.py --days 7
  python puid_repair.py --days 7 --apply
  python puid_repair.py --from-date 2026-10-01 --apply
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
    "default_days": 7,
}

FILE1 = "File1.csv"   # User status change
FILE2 = "File2.csv"   # Manager change


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


def extract_timestamps(history_text: str) -> List[datetime]:
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
# FILE SCHEMA (both files)
# ============================================================================

FILE_FIELDS = [
    "site_url", "site_id", "site_title",
    "owner_email", "owner_upn", "owner_display_name",
    "trigger", "trigger_timestamp", "change_events",
    "mismatch_status", "mismatch_checked_at",
    "repair_status", "repair_attempted_at", "repair_message",
    "current_nameid", "reference_nameid", "verified_nameid",
    "resolved", "last_updated", "error",
]


# ============================================================================
# DETECTORS  (no enabled check)
# ============================================================================

def detect_user_status_change(row: Dict[str, str]) -> Optional[Dict[str, Any]]:
    """
    User status change: Not Found >> Enabled.
    Current status must be Enabled (this IS the pattern — not an enabled gate).
    """
    history = safe_str(row.get("User Account Status Change History", ""))
    current = safe_str(row.get("User Account Status", "")).lower()

    # This is required by the pattern itself, not an extra gate.
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

    latest = latest_timestamp(history)
    if not latest:
        return None

    return {
        "history": history,
        "timestamp": latest.strftime("%Y-%m-%d %H:%M:%S"),
        "latest_dt": latest,
    }


def detect_manager_change(row: Dict[str, str]) -> Optional[Dict[str, Any]]:
    """
    Manager change: any >> transition in Manager Change History.
    NO enabled check — process even if user is disabled/deleted/not found.
    """
    history = safe_str(row.get("Manager Change History", ""))
    if not history:
        return None

    h = history.lower()
    h = re.sub(r"\s*>>\s*", " >> ", h)
    h = re.sub(r"\s*→\s*", " >> ", h)
    if " >> " not in h:
        return None

    latest = latest_timestamp(history)
    if not latest:
        return None

    return {
        "history": history,
        "timestamp": latest.strftime("%Y-%m-%d %H:%M:%S"),
        "latest_dt": latest,
    }


# ============================================================================
# BUILD FRESH ROWS (date filter + skip logic)
# ============================================================================

def build_fresh_rows(
    master_rows: List[Dict[str, str]],
    detector,
    cutoff_dt: datetime,
    existing_index: Dict[str, Dict[str, str]],
    log: logging.Logger,
) -> List[Dict[str, Any]]:
    ts_now = now_iso()
    fresh: List[Dict[str, Any]] = []
    stats = {"scanned": 0, "pass_filter": 0, "skip_final": 0}

    for r in master_rows:
        stats["scanned"] += 1
        hit = detector(r)
        if not hit:
            continue

        latest_dt = hit["latest_dt"]
        if latest_dt < cutoff_dt:
            continue
        stats["pass_filter"] += 1

        site_url = safe_str(r.get("Site URL", ""))
        if not site_url:
            continue

        prior = existing_index.get(site_url)
        prior_status = (prior or {}).get("repair_status", "").lower()
        prior_ts_str = (prior or {}).get("trigger_timestamp", "")
        prior_dt = None
        if prior_ts_str:
            try:
                prior_dt = datetime.strptime(prior_ts_str, "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)
            except ValueError:
                prior_dt = None

        is_newer_change = (prior_dt is None) or (latest_dt > prior_dt)
        is_retry = prior_status in ("failed", "error", "pending")

        # Skip if already final and no newer change
        if prior and prior_status in ("success", "repaired", "no_mismatch") and not is_newer_change:
            stats["skip_final"] += 1
            continue

        row: Dict[str, Any] = {
            "site_url": site_url,
            "site_id": safe_str(r.get("Site ID", "")),
            "site_title": safe_str(r.get("Title", "")),
            "owner_email": safe_str(r.get("Owner Email", "") or r.get("User Email", "")),
            "owner_upn": safe_str(r.get("User UPN", "")),
            "owner_display_name": safe_str(r.get("User Display Name", "")),
            "trigger": "",
            "trigger_timestamp": hit["timestamp"],
            "change_events": hit["history"],
            "mismatch_status": "pending",
            "mismatch_checked_at": "",
            "repair_status": "pending",
            "repair_attempted_at": "",
            "repair_message": "",
            "current_nameid": "",
            "reference_nameid": "",
            "verified_nameid": "",
            "resolved": "no",
            "last_updated": ts_now,
            "error": "",
        }

        if is_retry and prior:
            row["current_nameid"] = prior.get("current_nameid", "")
            row["reference_nameid"] = prior.get("reference_nameid", "")

        fresh.append(row)

    log.info("  Date filter >= %s | scanned=%d | passed=%d | skip(final)=%d",
             cutoff_dt.strftime("%Y-%m-%d"), stats["scanned"],
             stats["pass_filter"], stats["skip_final"])
    return fresh


# ============================================================================
# UPSERT (no duplicates)
# ============================================================================

def upsert_into_file(path: Path, fresh_rows: List[Dict[str, Any]],
                     log: logging.Logger) -> None:
    existing = read_csv(path)
    index: Dict[str, Dict[str, Any]] = {
        r.get("site_url", ""): dict(r) for r in existing if r.get("site_url")
    }

    created = updated = 0
    for rec in fresh_rows:
        url = rec.get("site_url", "")
        if not url:
            continue
        if url in index:
            index[url] = rec
            updated += 1
        else:
            index[url] = rec
            created += 1

    write_csv(path, list(index.values()), FILE_FIELDS)
    log.info("%s: %d total | new=%d | updated=%d",
             path.name, len(index), created, updated)


# ============================================================================
# MISMATCH CHECK
# ============================================================================

def check_mismatch(client: SharePointClient, row: Dict[str, str],
                   log: logging.Logger) -> Dict[str, Any]:
    rec = dict(row)
    try:
        upn = row.get("owner_upn", "").strip()
        if not upn:
            rec.update(mismatch_status="error", repair_status="error",
                       error="No owner UPN",
                       mismatch_checked_at=now_iso(), last_updated=now_iso())
            return rec

        current = client.get_site_owner(row["site_url"]).get("current_nameid") or ""
        reference = client.get_reference_nameid(upn) or ""
        rec["current_nameid"] = current
        rec["reference_nameid"] = reference
        rec["mismatch_checked_at"] = now_iso()
        rec["last_updated"] = now_iso()

        if not current or not reference:
            rec.update(mismatch_status="error", repair_status="error",
                       error="Could not read current or reference NameId")
        elif current == reference:
            rec.update(mismatch_status="not_found", repair_status="no_mismatch",
                       repair_message="NameId already matches reference",
                       resolved="yes")
        else:
            rec.update(mismatch_status="found", repair_status="pending")
    except Exception as exc:
        rec.update(mismatch_status="error", repair_status="error",
                   error=str(exc),
                   mismatch_checked_at=now_iso(), last_updated=now_iso())
        log.exception("Mismatch check failed: %s", rec.get("site_url"))
    return rec


def run_mismatch(client: SharePointClient, rows: List[Dict[str, Any]],
                 max_workers: int, log: logging.Logger) -> List[Dict[str, Any]]:
    if not rows:
        return []
    log.info("Mismatch check for %d row(s)", len(rows))
    out: List[Dict[str, Any]] = []
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = [ex.submit(check_mismatch, client, r, log) for r in rows]
        for fut in as_completed(futs):
            try:
                rec = fut.result()
                out.append(rec)
                log.info("%s : %s", rec.get("mismatch_status", "").upper(),
                         rec.get("site_url", ""))
            except Exception:
                log.exception("Mismatch future failed")
    return out


# ============================================================================
# REPAIR
# ============================================================================

def repair_one(client: SharePointClient, row: Dict[str, str],
               log: logging.Logger) -> Dict[str, Any]:
    rec = dict(row)
    rec["repair_attempted_at"] = now_iso()
    site_url = row["site_url"]
    upn = row["owner_upn"]
    try:
        digest = client.get_digest(site_url)

        owner = client.get_site_owner(site_url)
        current_uid = owner.get("user_id")
        if not current_uid:
            raise RuntimeError("Owner user ID missing on site")

        repair_user = client.ensure_user(site_url, digest, client.cfg["repair_account"])
        repair_uid = str(repair_user.get("Id"))
        client.set_admin(site_url, digest, repair_uid, True)

        client.set_admin(site_url, digest, current_uid, False)
        client.remove_user(site_url, digest, current_uid)
        time.sleep(client.cfg.get("sleep_after_remove_seconds", 2))

        readded = client.ensure_user(site_url, digest, upn)
        readded_uid = str(readded.get("Id"))

        if client.cfg.get("readded_user_site_admin", True):
            client.set_admin(site_url, digest, readded_uid, True)

        try:
            client.set_admin(site_url, digest, repair_uid, False)
            client.remove_user(site_url, digest, repair_uid)
        except Exception as exc:
            log.warning("Repair account cleanup failed: %s", exc)

        verified = client.get_site_owner(site_url).get("current_nameid") or ""
        rec["verified_nameid"] = verified

        if rec.get("reference_nameid") and verified == rec["reference_nameid"]:
            rec.update(repair_status="success", resolved="yes",
                       repair_message="NameId verified after re-add")
        else:
            rec.update(repair_status="failed", resolved="no",
                       repair_message="NameId not verified after re-add")
    except Exception as exc:
        rec.update(repair_status="error", resolved="error",
                   error=str(exc), repair_message=str(exc))
        log.exception("Repair failed: %s", site_url)

    rec["last_updated"] = now_iso()
    return rec


def run_repairs(client: SharePointClient, rows: List[Dict[str, Any]],
                max_workers: int, log: logging.Logger) -> List[Dict[str, Any]]:
    if not rows:
        return []
    log.info("Repairing %d site(s)", len(rows))
    out: List[Dict[str, Any]] = []
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = [ex.submit(repair_one, client, r, log) for r in rows]
        for fut in as_completed(futs):
            try:
                rec = fut.result()
                out.append(rec)
                log.info("REPAIR %s : %s",
                         rec.get("repair_status", "").upper(),
                         rec.get("site_url", ""))
            except Exception:
                log.exception("Repair future failed")
    return out


# ============================================================================
# WORKFLOW RUNNER
# ============================================================================

def run_workflow(
    client: SharePointClient,
    path: Path,
    master_rows: List[Dict[str, str]],
    detector,
    cutoff_dt: datetime,
    apply_changes: bool,
    max_workers: int,
    log: logging.Logger,
    label: str,
) -> None:
    log.info("=" * 70)
    log.info("Workflow: %s (>= %s)", label, cutoff_dt.strftime("%Y-%m-%d"))
    log.info("=" * 70)

    existing = {r.get("site_url", ""): r for r in read_csv(path) if r.get("site_url")}

    fresh = build_fresh_rows(master_rows, detector, cutoff_dt, existing, log)
    log.info("  Fresh rows to process: %d", len(fresh))
    if not fresh:
        log.info("  Nothing to do.")
        return

    checked = run_mismatch(client, fresh, max_workers, log)

    if apply_changes:
        to_repair = [r for r in checked
                     if r.get("mismatch_status") == "found"
                     and r.get("repair_status") == "pending"]
        if to_repair:
            repaired = run_repairs(client, to_repair, max_workers, log)
            by_url = {r["site_url"]: r for r in repaired}
            for i, r in enumerate(checked):
                if r["site_url"] in by_url:
                    checked[i] = by_url[r["site_url"]]
        else:
            log.info("  No mismatches to repair.")
    else:
        log.info("  Report-only mode -> repairs skipped.")

    upsert_into_file(path, checked, log)


# ============================================================================
# MAIN
# ============================================================================

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="PUID repair — 2 cumulative files, date-filtered, no duplicates, master read-only")
    p.add_argument("--config", default=None)
    p.add_argument("--master-report", default=None)
    p.add_argument("--days", type=int, default=None,
                   help="Process transitions within the last N days (default from config).")
    p.add_argument("--from-date", default=None,
                   help="YYYY-MM-DD. Process transitions on/after this date.")
    p.add_argument("--apply", action="store_true",
                   help="Actually repair (overrides run_mode).")
    p.add_argument("--max-workers", type=int, default=None)
    return p.parse_args()


def resolve_cutoff(args, cfg) -> datetime:
    if args.from_date:
        return datetime.fromisoformat(f"{args.from_date}T00:00:00+00:00")
    days = args.days if args.days is not None else int(cfg.get("default_days", 7))
    return (datetime.now(timezone.utc) - timedelta(days=days)).replace(
        hour=0, minute=0, second=0, microsecond=0
    )


def main() -> int:
    args = parse_args()
    cfg = load_config(args.config)
    apply_changes = args.apply or (cfg.get("run_mode") or "report_only").lower() == "apply"

    report_dir = Path(cfg["report_root"])
    log = setup_logger(report_dir)

    master_path = Path(args.master_report or cfg["master_report_path"])
    if not master_path.exists():
        log.error("Master report not found: %s", master_path)
        return 1

    cutoff_dt = resolve_cutoff(args, cfg)
    log.info("=" * 70)
    log.info("PUID Repair — starting")
    log.info("Mode: %s", "APPLY" if apply_changes else "REPORT ONLY")
    log.info("Master (read-only): %s", master_path)
    log.info("Cutoff date (>=): %s", cutoff_dt.strftime("%Y-%m-%d %H:%M:%S"))
    log.info("  File1 (user status)  : %s", report_dir / FILE1)
    log.info("  File2 (manager)      : %s", report_dir / FILE2)
    log.info("=" * 70)

    client = SharePointClient(cfg, log)
    max_workers = args.max_workers or cfg.get("max_workers", 5)

    # Read master ONCE. Never written.
    master_rows = read_csv(master_path)
    log.info("Master rows loaded: %d", len(master_rows))

    try:
        run_workflow(
            client, report_dir / FILE1, master_rows,
            detector=detect_user_status_change,
            cutoff_dt=cutoff_dt,
            apply_changes=apply_changes,
            max_workers=max_workers,
            log=log,
            label="User status change",
        )
        run_workflow(
            client, report_dir / FILE2, master_rows,
            detector=detect_manager_change,
            cutoff_dt=cutoff_dt,
            apply_changes=apply_changes,
            max_workers=max_workers,
            log=log,
            label="Manager change",
        )

        log.info("Master report left unmodified: %s", master_path)
        log.info("=" * 70)
        log.info("DONE")
        log.info("  File1 : %s", report_dir / FILE1)
        log.info("  File2 : %s", report_dir / FILE2)
        log.info("=" * 70)
        return 0

    except Exception:
        log.exception("Workflow failed")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())