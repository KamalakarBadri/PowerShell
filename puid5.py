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
from dataclasses import dataclass, asdict, field
from datetime import datetime, timedelta, timezone
from pathlib import Path
from threading import Lock
from typing import Any, Dict, List, Optional

import requests
from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate


DEFAULT_CONFIG: Dict[str, Any] = {
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "tenant_name": "geekbyteonline.onmicrosoft.com",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "client_secret": "REPLACE_ME",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    "repair_account": "edit@geekbyte.online",
    "new_id_site_url": "https://geekbyteonline.sharepoint.com/sites/2DayRetention",
    "onedrive_host": "https://geekbyteonline-my.sharepoint.com",
    "run_mode": "report_only",
    "report_root": "reports",
    "max_workers": 5,
    "sleep_after_remove_seconds": 2,
    "api_throttle_seconds": 0.25,          # NEW: wait 0.25s before every API call
    "readded_user_site_admin": True,
    "cleanup_reference_site_user": True,
    "request_timeout_seconds": 60,
    "scopes": {
        "graph": "https://graph.microsoft.com/.default",
        "sharepoint": "https://geekbyteonline.sharepoint.com/.default",
    },
}


# ---------- Config / Logging ----------

def load_config(config_path: Optional[str]) -> Dict[str, Any]:
    config = json.loads(json.dumps(DEFAULT_CONFIG))
    if config_path:
        with open(config_path, "r", encoding="utf-8") as handle:
            file_config = json.load(handle)
        merge_dict(config, file_config)
    return config


def merge_dict(target: Dict[str, Any], source: Dict[str, Any]) -> None:
    for key, value in source.items():
        if isinstance(value, dict) and isinstance(target.get(key), dict):
            merge_dict(target[key], value)
        else:
            target[key] = value


def setup_logger(log_file: Path) -> logging.Logger:
    logger = logging.getLogger("onedrive_puid_repair")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    return logger


def normalize_run_mode(value: Optional[str]) -> str:
    mode = (value or "report_only").strip().lower()
    if mode not in {"report_only", "apply"}:
        raise ValueError("run_mode must be either 'report_only' or 'apply'")
    return mode


# ---------- Record ----------

@dataclass
class RepairRecord:
    site_url: str = ""
    site_created: str = ""
    site_title: str = ""
    site_id: str = ""
    run_mode: str = "report_only"
    owner_upn: str = ""
    current_user_id: Optional[str] = None
    current_nameid: Optional[str] = None
    owner_login_name: Optional[str] = None
    owner_title: Optional[str] = None
    reference_nameid: Optional[str] = None
    reference_user_id: Optional[str] = None
    reference_cleanup_status: Optional[str] = None
    reference_cleanup_message: str = ""
    nameid_match: bool = False
    action: str = "skipped"
    action_status: str = "pending"
    readded_user_id: Optional[str] = None
    verified_nameid: Optional[str] = None
    verified_match: Optional[bool] = None
    repair_account_user_id: Optional[str] = None
    repair_account_cleanup_status: Optional[str] = None
    repair_account_cleanup_message: str = ""
    resolved: str = "no"                   # NEW: yes / no / error
    error: str = ""                        # NEW: error text if any
    message: str = ""


# ---------- Client ----------

class Microsoft365RepairClient:
    def __init__(self, config: Dict[str, Any], logger: logging.Logger):
        self.config = config
        self.logger = logger
        self.request_timeout = config.get("request_timeout_seconds", 60)
        self.throttle = config.get("api_throttle_seconds", 0.25)
        self._token_cache: Dict[str, str] = {}
        self._throttle_lock = Lock()

    def _throttle_wait(self) -> None:
        """Wait 0.25s before every API call (thread-safe)."""
        if self.throttle and self.throttle > 0:
            time.sleep(self.throttle)

    # ---- Token acquisition ----

    def get_token(self, scope_key: str) -> str:
        scope = self.config["scopes"][scope_key]
        if scope in self._token_cache:
            return self._token_cache[scope]

        token = self.get_token_with_certificate(scope)
        if not token:
            token = self.get_token_with_secret(scope)
        if not token:
            raise RuntimeError(f"Failed to obtain token for scope {scope}")

        self._token_cache[scope] = token
        return token

    def get_token_with_certificate(self, scope: str) -> Optional[str]:
        try:
            cert_path = self.config["certificate_path"]
            key_path = self.config["private_key_path"]
            if not os.path.exists(cert_path) or not os.path.exists(key_path):
                return None

            with open(cert_path, "rb") as cert_file:
                certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
            with open(key_path, "rb") as key_file:
                private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())

            now = int(time.time())
            jwt_header = {
                "alg": "RS256",
                "typ": "JWT",
                "x5t": base64.urlsafe_b64encode(certificate.fingerprint(hashes.SHA1())).decode().rstrip("="),
            }
            jwt_payload = {
                "aud": f"https://login.microsoftonline.com/{self.config['tenant_id']}/oauth2/v2.0/token",
                "exp": now + 300,
                "iss": self.config["app_id"],
                "jti": str(uuid.uuid4()),
                "nbf": now,
                "sub": self.config["app_id"],
            }

            encoded_header = base64.urlsafe_b64encode(json.dumps(jwt_header).encode()).decode().rstrip("=")
            encoded_payload = base64.urlsafe_b64encode(json.dumps(jwt_payload).encode()).decode().rstrip("=")
            jwt_unsigned = f"{encoded_header}.{encoded_payload}"
            signature = private_key.sign(jwt_unsigned.encode(), padding.PKCS1v15(), hashes.SHA256())
            encoded_signature = base64.urlsafe_b64encode(signature).decode().rstrip("=")
            client_assertion = f"{jwt_unsigned}.{encoded_signature}"

            token_url = f"https://login.microsoftonline.com/{self.config['tenant_id']}/oauth2/v2.0/token"
            self._throttle_wait()
            response = requests.post(
                token_url,
                data={
                    "client_id": self.config["app_id"],
                    "client_assertion": client_assertion,
                    "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                    "scope": scope,
                    "grant_type": "client_credentials",
                },
                timeout=self.request_timeout,
            )
            if response.status_code == 200:
                return response.json()["access_token"]
            self.logger.warning("Certificate auth failed: %s", response.text)
            return None
        except Exception:
            self.logger.exception("Certificate auth error")
            return None

    def get_token_with_secret(self, scope: str) -> Optional[str]:
        try:
            token_url = f"https://login.microsoftonline.com/{self.config['tenant_id']}/oauth2/v2.0/token"
            self._throttle_wait()
            response = requests.post(
                token_url,
                data={
                    "client_id": self.config["app_id"],
                    "client_secret": self.config["client_secret"],
                    "scope": scope,
                    "grant_type": "client_credentials",
                },
                timeout=self.request_timeout,
            )
            if response.status_code == 200:
                return response.json()["access_token"]
            self.logger.warning("Client secret auth failed: %s", response.text)
            return None
        except Exception:
            self.logger.exception("Client secret auth error")
            return None

    # ---- Headers / helpers ----

    def sp_headers(self, token: str, with_json: bool = True) -> Dict[str, str]:
        headers = {"Authorization": f"Bearer {token}"}
        if with_json:
            headers["Accept"] = "application/json;odata=verbose"
            headers["Content-Type"] = "application/json;odata=verbose"
        return headers

    def get_request_digest(self, site_url: str, token: str) -> str:
        url = f"{site_url.rstrip('/')}/_api/contextinfo"
        self._throttle_wait()
        response = requests.post(url, headers=self.sp_headers(token), timeout=self.request_timeout)
        response.raise_for_status()
        return response.json()["d"]["GetContextWebInformation"]["FormDigestValue"]

    def ensure_user(self, site_url: str, token: str, request_digest: str, user_upn: str) -> Dict[str, Any]:
        url = f"{site_url.rstrip('/')}/_api/web/ensureuser"
        body = {"logonName": user_upn}
        self._throttle_wait()
        response = requests.post(
            url,
            headers={**self.sp_headers(token), "X-RequestDigest": request_digest},
            json=body,
            timeout=self.request_timeout,
        )
        response.raise_for_status()
        return response.json()["d"]

    def set_site_admin(self, site_url: str, token: str, request_digest: str, user_id: str, is_admin: bool) -> None:
        url = f"{site_url.rstrip('/')}/_api/web/getuserbyid({user_id})"
        body = {"__metadata": {"type": "SP.User"}, "IsSiteAdmin": is_admin}
        self._throttle_wait()
        response = requests.post(
            url,
            headers={
                **self.sp_headers(token),
                "X-RequestDigest": request_digest,
                "X-HTTP-Method": "MERGE",
                "IF-MATCH": "*",
            },
            json=body,
            timeout=self.request_timeout,
        )
        if response.status_code not in (200, 204):
            raise RuntimeError(f"Failed to set site admin for user {user_id}: {response.text}")

    def remove_user_by_id(self, site_url: str, token: str, request_digest: str, user_id: str) -> None:
        url = f"{site_url.rstrip('/')}/_api/web/siteusers/removebyid({user_id})"
        self._throttle_wait()
        response = requests.post(
            url,
            headers={**self.sp_headers(token), "X-RequestDigest": request_digest},
            timeout=self.request_timeout,
        )
        if response.status_code not in (200, 204):
            raise RuntimeError(f"Failed to remove user {user_id}: {response.text}")

    def log_site_step(self, site_url: str, message: str) -> None:
        self.logger.info("[%s] %s", site_url, message)

    # ---- SharePoint queries ----

    def get_site_owner_info(self, site_url: str) -> Dict[str, Optional[str]]:
        token = self.get_token("sharepoint")
        url = f"{site_url.rstrip('/')}/_api/site/owner"
        self._throttle_wait()
        response = requests.get(
            url,
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json;odata=verbose"},
            timeout=self.request_timeout,
        )
        response.raise_for_status()

        owner = response.json().get("d", {})
        owner_upn = owner.get("UserPrincipalName") or owner.get("Email")
        owner_nameid = (owner.get("UserId") or {}).get("NameId")

        return {
            "user_id": str(owner.get("Id")) if owner.get("Id") is not None else None,
            "owner_upn": owner_upn.lower() if owner_upn else None,
            "current_nameid": owner_nameid,
            "login_name": owner.get("LoginName"),
            "title": owner.get("Title"),
            "email": owner.get("Email"),
            "user_principal_name": owner.get("UserPrincipalName"),
            "is_site_admin": owner.get("IsSiteAdmin", False),
        }

    def get_reference_site_nameid_and_cleanup(self, target_upn: str) -> Dict[str, Optional[str]]:
        token = self.get_token("sharepoint")
        digest = self.get_request_digest(self.config["new_id_site_url"], token)
        ensured = self.ensure_user(self.config["new_id_site_url"], token, digest, target_upn)

        reference_user_id = str(ensured.get("Id")) if ensured.get("Id") is not None else None
        cleanup_status = "skipped"
        cleanup_message = ""

        if self.config.get("cleanup_reference_site_user", True) and reference_user_id:
            try:
                self.remove_user_by_id(self.config["new_id_site_url"], token, digest, reference_user_id)
                cleanup_status = "removed"
            except Exception as exc:
                cleanup_status = "error"
                cleanup_message = str(exc)

        return {
            "nameid": (ensured.get("UserId") or {}).get("NameId"),
            "reference_user_id": reference_user_id,
            "cleanup_status": cleanup_status,
            "cleanup_message": cleanup_message,
        }

    # ---- Graph discovery (date range) ----

    def discover_onedrives_in_range(self, start_utc: datetime, end_utc: datetime) -> List[Dict[str, Any]]:
        sites: List[Dict[str, Any]] = []
        token = self.get_token("graph")
        url = "https://graph.microsoft.com/v1.0/sites?$select=id,name,webUrl,createdDateTime&$top=999"

        while url:
            self._throttle_wait()
            response = requests.get(
                url,
                headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
                timeout=self.request_timeout,
            )
            response.raise_for_status()
            payload = response.json()

            for item in payload.get("value", []):
                if not is_personal_site(item, self.config["onedrive_host"]):
                    continue

                created_raw = item.get("createdDateTime")
                site_url = item.get("webUrl")
                if not created_raw or not site_url:
                    continue

                created = parse_datetime(created_raw)
                if not created or not (start_utc <= created < end_utc):
                    continue

                sites.append(
                    {
                        "owner_upn": "",
                        "site_url": site_url.rstrip("/"),
                        "site_created": created.isoformat(),
                        "site_title": item.get("name") or "",
                        "site_id": item.get("id") or "",
                    }
                )

            url = payload.get("@odata.nextLink")

        return sorted(sites, key=lambda item: item["site_created"], reverse=True)

    # ---- Repair ----

    def repair_onedrive_owner(self, site: Dict[str, Any], apply_changes: bool) -> RepairRecord:
        site_url = site["site_url"]
        record = RepairRecord(
            site_url=site_url,
            site_created=site["site_created"],
            site_title=site["site_title"],
            site_id=site["site_id"],
            run_mode="apply" if apply_changes else "report_only",
            owner_upn=site.get("owner_upn", ""),
        )

        try:
            self.log_site_step(site_url, "Starting processing")
            owner_info = self.get_site_owner_info(site_url)
            owner_upn = owner_info.get("owner_upn")
            if not owner_upn:
                record.action_status = "not_found"
                record.error = "Could not resolve owner UPN."
                record.resolved = "no"
                record.message = record.error
                return record

            record.owner_upn = owner_upn
            record.current_user_id = owner_info.get("user_id")
            record.current_nameid = owner_info.get("current_nameid")
            record.owner_login_name = owner_info.get("login_name")
            record.owner_title = owner_info.get("title")

            reference_result = self.get_reference_site_nameid_and_cleanup(owner_upn)
            record.reference_nameid = reference_result.get("nameid")
            record.reference_user_id = reference_result.get("reference_user_id")
            record.reference_cleanup_status = reference_result.get("cleanup_status")
            record.reference_cleanup_message = reference_result.get("cleanup_message") or ""

            record.nameid_match = (
                bool(record.current_nameid)
                and bool(record.reference_nameid)
                and record.current_nameid == record.reference_nameid
            )

            if record.nameid_match:
                record.action = "none"
                record.action_status = "already_match"
                record.resolved = "yes"
                record.message = "NameId already matches."
                return record

            record.action = "remove_readd"
            if not apply_changes:
                record.action_status = "report_only"
                record.resolved = "no"
                record.message = "Mismatch found. Report-only mode."
                return record

            token = self.get_token("sharepoint")
            digest = self.get_request_digest(site_url, token)

            repair_user = self.ensure_user(site_url, token, digest, self.config["repair_account"])
            repair_user_id = str(repair_user.get("Id"))
            record.repair_account_user_id = repair_user_id
            self.set_site_admin(site_url, token, digest, repair_user_id, True)

            if not record.current_user_id:
                raise RuntimeError("Owner user ID missing.")

            self.set_site_admin(site_url, token, digest, record.current_user_id, False)
            self.remove_user_by_id(site_url, token, digest, record.current_user_id)
            time.sleep(self.config.get("sleep_after_remove_seconds", 2))

            readded_user = self.ensure_user(site_url, token, digest, owner_upn)
            record.readded_user_id = str(readded_user.get("Id"))

            if self.config.get("readded_user_site_admin", True):
                self.set_site_admin(site_url, token, digest, record.readded_user_id, True)

            # cleanup repair account
            try:
                self.set_site_admin(site_url, token, digest, repair_user_id, False)
                self.remove_user_by_id(site_url, token, digest, repair_user_id)
                record.repair_account_cleanup_status = "removed"
            except Exception as cleanup_exc:
                record.repair_account_cleanup_status = "error"
                record.repair_account_cleanup_message = str(cleanup_exc)

            verified = self.get_site_owner_info(site_url)
            record.verified_nameid = verified.get("current_nameid") if verified else None
            record.verified_match = record.verified_nameid == record.reference_nameid if record.reference_nameid else False

            if record.verified_match:
                record.action_status = "resolved"
                record.resolved = "yes"
                record.message = "Mismatch repaired and verified."
            else:
                record.action_status = "readd_complete_unverified"
                record.resolved = "no"
                record.message = "User re-added but NameId not verified."
            return record

        except Exception as exc:
            record.action_status = "error"
            record.resolved = "error"
            record.error = str(exc)
            record.message = str(exc)
            self.logger.exception("[%s] Repair failed", site_url)
            return record


# ---------- Utilities ----------

def parse_datetime(value: str) -> Optional[datetime]:
    try:
        if value.endswith("Z"):
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
        parsed = datetime.fromisoformat(value)
        if parsed.tzinfo is None:
            return parsed.replace(tzinfo=timezone.utc)
        return parsed.astimezone(timezone.utc)
    except ValueError:
        try:
            parsed = datetime.strptime(value, "%m/%d/%Y %I:%M:%S %p")
            return parsed.replace(tzinfo=timezone.utc)
        except ValueError:
            return None


def is_personal_site(site: Dict[str, Any], onedrive_host: str) -> bool:
    web_url = (site.get("webUrl") or "").lower()
    if not web_url:
        return False
    is_personal_flag = site.get("isPersonalSite")
    if isinstance(is_personal_flag, bool):
        return is_personal_flag
    normalized_host = onedrive_host.lower().rstrip("/")
    return web_url.startswith(normalized_host) and "/personal/" in web_url


def ensure_directory(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def write_csv(path: Path, rows: List[Dict[str, Any]], fieldnames: Optional[List[str]] = None) -> None:
    if not rows and not fieldnames:
        with open(path, "w", encoding="utf-8", newline="") as handle:
            handle.write("")
        return

    if not fieldnames:
        fieldnames = []
        seen = set()
        for row in rows:
            for key in row.keys():
                if key not in seen:
                    fieldnames.append(key)
                    seen.add(key)

    with open(path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def append_or_update_master(master_path: Path, records: List[RepairRecord]) -> None:
    """Master CSV — one row per site, updated on every run."""
    fieldnames = ["site_url", "owner_upn", "site_created", "old_nameid", "new_nameid", "resolved"]

    existing: Dict[str, Dict[str, Any]] = {}
    if master_path.exists():
        with open(master_path, "r", encoding="utf-8", newline="") as handle:
            reader = csv.DictReader(handle)
            for row in reader:
                existing[row["site_url"]] = row

    for rec in records:
        existing[rec.site_url] = {
            "site_url": rec.site_url,
            "owner_upn": rec.owner_upn,
            "site_created": rec.site_created,
            "old_nameid": rec.current_nameid or "",
            "new_nameid": rec.verified_nameid or rec.reference_nameid or "",
            "resolved": rec.resolved,
        }

    with open(master_path, "w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in sorted(existing.values(), key=lambda r: r.get("site_created", ""), reverse=True):
            writer.writerow(row)


# ---------- Processing ----------

def process_sites(
    client: Microsoft365RepairClient,
    sites: List[Dict[str, Any]],
    apply_changes: bool,
    max_workers: int,
) -> List[RepairRecord]:
    records: List[RepairRecord] = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(client.repair_onedrive_owner, site, apply_changes) for site in sites]
        for future in as_completed(futures):
            records.append(future.result())
    return sorted(records, key=lambda item: item.owner_upn.lower())


# ---------- CLI ----------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Detect PUID mismatches for OneDrive sites in a date range and optionally repair."
    )
    parser.add_argument("--config", help="Path to JSON config file.", default=None)
    parser.add_argument("--from-date", required=False, default=None,
                        help="Start date YYYY-MM-DD (inclusive). Default: yesterday UTC.")
    parser.add_argument("--to-date", required=False, default=None,
                        help="End date YYYY-MM-DD (inclusive). Default: --from-date.")
    parser.add_argument("--max-workers", type=int, default=None, help="Parallel workers.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = load_config(args.config)
    run_mode = normalize_run_mode(config.get("run_mode"))
    apply_changes = run_mode == "apply"

    # ---- Date range resolution ----
    if args.from_date:
        from_dt = datetime.fromisoformat(f"{args.from_date}T00:00:00+00:00")
    else:
        from_dt = (datetime.now(timezone.utc) - timedelta(days=1)).replace(
            hour=0, minute=0, second=0, microsecond=0
        )

    if args.to_date:
        to_dt = datetime.fromisoformat(f"{args.to_date}T00:00:00+00:00") + timedelta(days=1)
    else:
        to_dt = from_dt + timedelta(days=1)

    if to_dt <= from_dt:
        print("ERROR: --to-date must be >= --from-date")
        return 2

    report_dir = Path(config["report_root"])
    ensure_directory(report_dir)

    logger = setup_logger(report_dir / "run.log")
    logger.info("OneDrive PUID repair job")
    logger.info("Mode: %s", run_mode)
    logger.info("Date range: %s -> %s", from_dt.isoformat(), to_dt.isoformat())

    client = Microsoft365RepairClient(config, logger)

    try:
        sites = client.discover_onedrives_in_range(from_dt, to_dt)
        logger.info("Discovered %s OneDrive site(s)", len(sites))

        records = process_sites(
            client,
            sites,
            apply_changes=apply_changes,
            max_workers=args.max_workers or config.get("max_workers", 5),
        )

        # ---- Split into matched / not matched ----
        matched_rows: List[Dict[str, Any]] = []
        not_matched_rows: List[Dict[str, Any]] = []

        for rec in records:
            base = {
                "site_url": rec.site_url,
                "owner_upn": rec.owner_upn,
                "site_created": rec.site_created,
                "old_nameid": rec.current_nameid or "",
                "new_nameid": rec.verified_nameid or rec.reference_nameid or "",
                "resolved": rec.resolved,
                "status": rec.action_status,
                "error": rec.error,
                "message": rec.message,
            }
            if rec.nameid_match:
                matched_rows.append(base)
            else:
                not_matched_rows.append(base)

        matched_path = report_dir / "matched.csv"
        not_matched_path = report_dir / "not_matched.csv"
        master_path = report_dir / "master.csv"

        write_csv(matched_path, matched_rows, fieldnames=[
            "site_url", "owner_upn", "site_created", "old_nameid", "new_nameid",
            "resolved", "status", "error", "message",
        ])
        write_csv(not_matched_path, not_matched_rows, fieldnames=[
            "site_url", "owner_upn", "site_created", "old_nameid", "new_nameid",
            "resolved", "status", "error", "message",
        ])

        # ---- Master file (cumulative) ----
        append_or_update_master(master_path, records)

        logger.info("Matched: %s", len(matched_rows))
        logger.info("Not matched: %s", len(not_matched_rows))
        logger.info("Master file: %s", master_path)
        logger.info("Reports written to %s", report_dir)
        return 0

    except Exception:
        logger.exception("Job failed")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
