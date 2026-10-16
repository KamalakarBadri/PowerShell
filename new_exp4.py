#!/usr/bin/env python3
"""
OneDrive PUID Mismatch Repair — master-driven, single output file.

Workflow:
  1. Read master CSV (read-only).
  2. Filter rows: history has "Not Found >> Enabled", current status Enabled,
     and not already PUID-repaired in our own results file.
  3. For each: compare site NameId vs reference-site NameId.
  4. report_only -> stop.  apply -> repair mismatches, verify.
  5. Write ONE file: reports/puid_repair.csv (upserted, never appended).

Master CSV is NEVER modified.
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
}

RESULT_FILE = "puid_repair.csv"

# Final statuses — these are never retried.
TERMINAL_OK = {"ok", "repair_success"}
# Retryable statuses — re-examined on the next run.
RETRYABLE = {"error_check", "mismatch", "repair_failed", "pending"}

RESULT_FIELDS = [
    "site_url", "site_id",
    "owner_upn", "owner_email",
    "current_nameid", "reference_nameid",
    "status", "message", "updated_at",
]


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
    """Full rewrite — upsert semantics, no duplicates, no append."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


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
# STEP 1 — FILTER MASTER (read-only)
# ============================================================================

def was_not_found_now_enabled(row: Dict[str, str]) -> bool:
    """True if history shows a 'Not Found >> Enabled' transition and
    the latest state is still Enabled."""
    history = safe_str(row.get("User Account Status Change History", ""))
    current = safe_str(row.get("User Account Status", "")).lower()

    if current != "enabled" or not history:
        return False

    h = history.lower()
    h = re.sub(r"\s*>>\s*", " >> ", h)
    h = re.sub(r"\s*→\s*", " >> ", h)
    h = re.sub(r"\s+", " ", h)

    if "not found >> enabled" not in h:
        return False

    matches = re.findall(r">>\s*([a-z ]+)", h)
    if matches and matches[-1].strip() != "enabled":
        return False
    return True


def build_worklist(master_rows: List[Dict[str, str]],
                   existing: Dict[str, Dict[str, str]],
                   log: logging.Logger) -> List[Dict[str, Any]]:
    """
    Build the worklist.  Upserts existing rows, preserving their status
    when it's terminal OK.  Never leaves anything at 'pending' forever:
    rows that are still pending after this function will be processed
    by the mismatch step.
    """
    ts = now_iso()
    work: Dict[str, Dict[str, Any]] = {}

    for r in master_rows:
        if not was_not_found_now_enabled(r):
            continue
        url = safe_str(r.get("Site URL", ""))
        if not url:
            continue

        prev = existing.get(url, {})
        prev_status = safe_str(prev.get("status", "")).lower()

        # Terminal success -> keep as-is, do not re-queue.
        if prev_status in TERMINAL_OK:
            work[url] = dict(prev)
            continue

        work[url] = {
            "site_url": url,
            "site_id": safe_str(r.get("Site ID", "")),
            "owner_upn": safe_str(r.get("User UPN", "")),
            "owner_email": safe_str(r.get("Owner Email", "") or r.get("User Email", "")),
            "current_nameid": safe_str(prev.get("current_nameid", "")),
            "reference_nameid": safe_str(prev.get("reference_nameid", "")),
            "status": "pending",
            "message": "",
            "updated_at": ts,
        }

    # Keep prior retryable rows that are no longer in master (site may
    # have dropped out of the report).  Terminal rows are dropped.
    for url, prev in existing.items():
        if url in work:
            continue
        if safe_str(prev.get("status", "")).lower() in RETRYABLE:
            work[url] = dict(prev)

    log.info("Worklist: %d site(s)", len(work))
    return list(work.values())


# ============================================================================
# STEP 2 — MISMATCH CHECK (sets terminal 'ok' or 'error_check')
# ============================================================================

def check_mismatch(client: SharePointClient, row: Dict[str, Any],
                   log: logging.Logger) -> Dict[str, Any]:
    rec = dict(row)
    rec["updated_at"] = now_iso()
    try:
        upn = safe_str(row.get("owner_upn", ""))
        if not upn:
            rec["status"] = "error_check"
            rec["message"] = "No owner UPN"
            return rec

        current = client.get_site_owner(rec["site_url"]).get("current_nameid") or ""
        reference = client.get_reference_nameid(upn) or ""
        rec["current_nameid"] = current
        rec["reference_nameid"] = reference

        if not current or not reference:
            rec["status"] = "error_check"
            rec["message"] = "Missing NameId (current or reference)"
        elif current == reference:
            rec["status"] = "ok"
            rec["message"] = "NameId already matches — nothing to do"
        else:
            rec["status"] = "mismatch"
            rec["message"] = "NameId mismatch — repair needed"
    except Exception as exc:
        rec["status"] = "error_check"
        rec["message"] = str(exc)
        log.exception("Mismatch check failed: %s", rec.get("site_url"))
    return rec


def run_mismatch(client: SharePointClient, worklist: List[Dict[str, Any]],
                 max_workers: int, log: logging.Logger) -> List[Dict[str, Any]]:
    todo = [r for r in worklist
            if safe_str(r.get("status", "")).lower() not in TERMINAL_OK]
    log.info("Mismatch checks: %d", len(todo))
    if not todo:
        return worklist

    by_url = {r["site_url"]: r for r in worklist}
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = {ex.submit(check_mismatch, client, r, log): r for r in todo}
        for fut in as_completed(futs):
            src = futs[fut]
            try:
                rec = fut.result()
                by_url[rec["site_url"]] = rec
                log.info("%-14s %s | %s",
                         rec["status"], rec["site_url"], rec["message"])
            except Exception:
                log.exception("Future failed: %s", src.get("site_url"))
    return list(by_url.values())


# ============================================================================
# STEP 3 — REPAIR (only when run_mode == apply)
# ============================================================================

def repair_one(client: SharePointClient, row: Dict[str, Any],
               log: logging.Logger) -> Dict[str, Any]:
    rec = dict(row)
    rec["updated_at"] = now_iso()
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
        rec["current_nameid"] = verified
        if rec["reference_nameid"] and verified == rec["reference_nameid"]:
            rec["status"] = "repair_success"
            rec["message"] = "NameId verified after re-add"
        else:
            rec["status"] = "repair_failed"
            rec["message"] = "NameId not verified after re-add"
    except Exception as exc:
        rec["status"] = "repair_failed"
        rec["message"] = str(exc)
        log.exception("Repair failed: %s", site_url)
    return rec


def run_repairs(client: SharePointClient, rows: List[Dict[str, Any]],
                max_workers: int, log: logging.Logger) -> List[Dict[str, Any]]:
    todo = [r for r in rows if safe_str(r.get("status", "")).lower() == "mismatch"]
    log.info("Repairs to attempt: %d", len(todo))
    if not todo:
        return rows

    by_url = {r["site_url"]: r for r in rows}
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        futs = {ex.submit(repair_one, client, r, log): r for r in todo}
        for fut in as_completed(futs):
            src = futs[fut]
            try:
                rec = fut.result()
                by_url[rec["site_url"]] = rec
                log.info("%-15s %s | %s",
                         rec["status"], rec["site_url"], rec["message"])
            except Exception:
                log.exception("Repair future failed: %s", src.get("site_url"))
    return list(by_url.values())


# ============================================================================
# MAIN
# ============================================================================

def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="PUID repair for 'Not Found -> Enabled' owners (master-driven)")
    p.add_argument("--config", default=None)
    p.add_argument("--master-report", default=None)
    p.add_argument("--skip-mismatch", action="store_true",
                   help="Reuse current statuses; only run repairs.")
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
        log.error("Run master_report.py first.")
        return 1

    result_path = report_dir / RESULT_FILE
    existing = {r["site_url"]: r for r in read_csv(result_path)
                if r.get("site_url")}

    log.info("=" * 70)
    log.info("PUID Repair — Not Found -> Enabled")
    log.info("Mode  : %s", "APPLY" if apply_changes else "REPORT ONLY")
    log.info("Master: %s (read-only)", master_path)
    log.info("Output: %s", result_path)
    log.info("=" * 70)

    client = SharePointClient(cfg, log)
    max_workers = args.max_workers or cfg.get("max_workers", 5)

    try:
        # ---- Step 1: build worklist from master ----
        master_rows = read_csv(master_path)
        worklist = build_worklist(master_rows, existing, log)
        if not worklist:
            log.info("Nothing to do.")
            write_csv(result_path, [], RESULT_FIELDS)
            return 0

        # ---- Step 2: mismatch check (assigns terminal 'ok' / 'error_check') ----
        if args.skip_mismatch:
            log.info("Skipping mismatch check (--skip-mismatch)")
        else:
            worklist = run_mismatch(client, worklist, max_workers, log)
            write_csv(result_path, worklist, RESULT_FIELDS)  # checkpoint

        # ---- Step 3: repairs ----
        if not apply_changes:
            log.info("report_only mode — repairs skipped")
        else:
            mismatches = [r for r in worklist
                          if safe_str(r.get("status", "")).lower() == "mismatch"]
            if not mismatches:
                log.info("No mismatches to repair.")
            else:
                worklist = run_repairs(client, worklist, max_workers, log)

        # ---- Final write ----
        write_csv(result_path, worklist, RESULT_FIELDS)

        # ---- Summary ----
        counts: Dict[str, int] = {}
        for r in worklist:
            counts[r["status"]] = counts.get(r["status"], 0) + 1
        log.info("=" * 70)
        log.info("DONE — %d site(s)", len(worklist))
        for k in sorted(counts):
            log.info("  %-15s %d", k, counts[k])
        log.info("Output: %s", result_path)
        log.info("=" * 70)
        return 0

    except Exception:
        log.exception("Workflow failed")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
