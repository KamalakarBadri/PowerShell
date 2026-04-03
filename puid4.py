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
from dataclasses import dataclass, asdict
from datetime import datetime, timedelta, timezone
from pathlib import Path
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
    "readded_user_site_admin": True,
    "cleanup_reference_site_user": True,
    "request_timeout_seconds": 60,
    "scopes": {
        "graph": "https://graph.microsoft.com/.default",
        "sharepoint": "https://geekbyteonline.sharepoint.com/.default",
    },
}


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


@dataclass
class RepairRecord:
    site_url: str
    site_created: str
    site_title: str
    site_id: str
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
    message: str = ""


class Microsoft365RepairClient:
    def __init__(self, config: Dict[str, Any], logger: logging.Logger):
        self.config = config
        self.logger = logger
        self.request_timeout = config.get("request_timeout_seconds", 60)
        self._token_cache: Dict[str, str] = {}

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
            token_body = {
                "client_id": self.config["app_id"],
                "client_assertion": "<redacted>",
                "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                "scope": scope,
                "grant_type": "client_credentials",
            }
            self.log_request("POST", token_url, token_body, "auth-certificate")
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
            self.log_response(response, "auth-certificate")
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
            token_body = {
                "client_id": self.config["app_id"],
                "client_secret": "<redacted>",
                "scope": scope,
                "grant_type": "client_credentials",
            }
            self.log_request("POST", token_url, token_body, "auth-secret")
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
            self.log_response(response, "auth-secret")
            if response.status_code == 200:
                return response.json()["access_token"]
            self.logger.warning("Client secret auth failed: %s", response.text)
            return None
        except Exception:
            self.logger.exception("Client secret auth error")
            return None

    def sp_headers(self, token: str, with_json: bool = True) -> Dict[str, str]:
        headers = {"Authorization": f"Bearer {token}"}
        if with_json:
            headers["Accept"] = "application/json;odata=verbose"
            headers["Content-Type"] = "application/json;odata=verbose"
        return headers

    def log_request(self, method: str, url: str, payload: Optional[Any] = None, context: Optional[str] = None) -> None:
        prefix = f"[{context}] " if context else ""
        self.logger.info("%sHTTP %s %s", prefix, method.upper(), url)
        if payload is not None:
            try:
                rendered = json.dumps(payload, ensure_ascii=True)
            except TypeError:
                rendered = str(payload)
            self.logger.info("%sRequest Body: %s", prefix, rendered)

    def log_response(self, response: requests.Response, context: Optional[str] = None) -> None:
        prefix = f"[{context}] " if context else ""
        self.logger.info("%sResponse Status: %s", prefix, response.status_code)

    def get_request_digest(self, site_url: str, token: str) -> str:
        url = f"{site_url.rstrip('/')}/_api/contextinfo"
        self.log_request("POST", url, context=site_url)
        response = requests.post(
            url,
            headers=self.sp_headers(token),
            timeout=self.request_timeout,
        )
        self.log_response(response, site_url)
        response.raise_for_status()
        return response.json()["d"]["GetContextWebInformation"]["FormDigestValue"]

    def ensure_user(self, site_url: str, token: str, request_digest: str, user_upn: str) -> Dict[str, Any]:
        url = f"{site_url.rstrip('/')}/_api/web/ensureuser"
        body = {"logonName": user_upn}
        self.log_request("POST", url, body, site_url)
        response = requests.post(
            url,
            headers={**self.sp_headers(token), "X-RequestDigest": request_digest},
            json=body,
            timeout=self.request_timeout,
        )
        self.log_response(response, site_url)
        response.raise_for_status()
        return response.json()["d"]

    def set_site_admin(self, site_url: str, token: str, request_digest: str, user_id: str, is_admin: bool) -> None:
        url = f"{site_url.rstrip('/')}/_api/web/getuserbyid({user_id})"
        body = {"__metadata": {"type": "SP.User"}, "IsSiteAdmin": is_admin}
        self.log_request("POST", url, body, site_url)
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
        self.log_response(response, site_url)
        if response.status_code not in (200, 204):
            raise RuntimeError(f"Failed to set site admin for user {user_id}: {response.text}")

    def remove_user_by_id(self, site_url: str, token: str, request_digest: str, user_id: str) -> None:
        url = f"{site_url.rstrip('/')}/_api/web/siteusers/removebyid({user_id})"
        self.log_request("POST", url, {"user_id": user_id}, site_url)
        response = requests.post(
            url,
            headers={**self.sp_headers(token), "X-RequestDigest": request_digest},
            timeout=self.request_timeout,
        )
        self.log_response(response, site_url)
        if response.status_code not in (200, 204):
            raise RuntimeError(f"Failed to remove user {user_id}: {response.text}")

    def log_site_step(self, site_url: str, message: str) -> None:
        self.logger.info("[%s] %s", site_url, message)

    def get_site_owner_info(self, site_url: str) -> Dict[str, Optional[str]]:
        token = self.get_token("sharepoint")
        url = f"{site_url.rstrip('/')}/_api/site/owner"
        self.log_request("GET", url, context=site_url)
        response = requests.get(
            url,
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json;odata=verbose"},
            timeout=self.request_timeout,
        )
        self.log_response(response, site_url)
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

    def get_new_site_nameid(self, target_upn: str) -> Optional[str]:
        token = self.get_token("sharepoint")
        digest = self.get_request_digest(self.config["new_id_site_url"], token)
        ensured = self.ensure_user(self.config["new_id_site_url"], token, digest, target_upn)
        return (ensured.get("UserId") or {}).get("NameId")

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

    def find_user_on_site(self, target_upn: str, site_url: str) -> Optional[Dict[str, Any]]:
        token = self.get_token("sharepoint")
        response = requests.get(
            f"{site_url.rstrip('/')}/_api/web/siteusers",
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json;odata=verbose"},
            timeout=self.request_timeout,
        )
        response.raise_for_status()
        data = response.json()
        users = data.get("d", {}).get("results", [])

        for user in users:
            email = (user.get("Email") or "").lower()
            login_name = (user.get("LoginName") or "").lower()
            user_principal_name = (user.get("UserPrincipalName") or "").lower()
            target = target_upn.lower()
            if email == target or target in login_name or user_principal_name == target:
                return {
                    "user_id": user.get("Id"),
                    "title": user.get("Title"),
                    "email": user.get("Email"),
                    "login_name": user.get("LoginName"),
                    "user_principal_name": user.get("UserPrincipalName"),
                    "is_site_admin": user.get("IsSiteAdmin", False),
                    "current_nameid": (user.get("UserId") or {}).get("NameId"),
                }
        return None

    def discover_recent_onedrives(self, start_utc: datetime, end_utc: datetime) -> List[Dict[str, Any]]:
        sites: List[Dict[str, Any]] = []
        token = self.get_token("graph")
        url = "https://graph.microsoft.com/v1.0/sites?$select=id,name,webUrl,createdDateTime&$top=999"

        while url:
            self.log_request("GET", url, context="graph-sites")
            response = requests.get(
                url,
                headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
                timeout=self.request_timeout,
            )
            self.log_response(response, "graph-sites")
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
                        "raw_graph_row": item,
                    }
                )

            url = payload.get("@odata.nextLink")

        return sorted(sites, key=lambda item: item["site_created"], reverse=True)

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
                record.message = "Could not resolve owner UPN from _api/site/owner."
                self.log_site_step(site_url, record.message)
                return record

            record.owner_upn = owner_upn
            record.current_user_id = owner_info.get("user_id")
            record.current_nameid = owner_info.get("current_nameid")
            record.owner_login_name = owner_info.get("login_name")
            record.owner_title = owner_info.get("title")

            self.log_site_step(site_url, f"Owner resolved: {record.owner_upn}")
            self.log_site_step(site_url, f"Current NameId: {record.current_nameid or 'N/A'}")

            reference_result = self.get_reference_site_nameid_and_cleanup(owner_upn)
            record.reference_nameid = reference_result.get("nameid")
            record.reference_user_id = reference_result.get("reference_user_id")
            record.reference_cleanup_status = reference_result.get("cleanup_status")
            record.reference_cleanup_message = reference_result.get("cleanup_message") or ""
            self.log_site_step(site_url, f"Reference NameId: {record.reference_nameid or 'N/A'}")
            self.log_site_step(
                site_url,
                f"Reference cleanup: {record.reference_cleanup_status}"
                + (f" ({record.reference_cleanup_message})" if record.reference_cleanup_message else ""),
            )
            record.nameid_match = (
                bool(record.current_nameid)
                and bool(record.reference_nameid)
                and record.current_nameid == record.reference_nameid
            )

            if record.nameid_match:
                record.action = "none"
                record.action_status = "already_match"
                record.message = "Owner NameId already matches reference site."
                self.log_site_step(site_url, "Status: ALREADY MATCHED")
                return record

            record.action = "remove_readd"
            self.log_site_step(site_url, "Mismatch detected")
            if not apply_changes:
                record.action_status = "report_only"
                record.message = "Mismatch found. Repair skipped because report-only mode is enabled."
                self.log_site_step(site_url, "Status: REPORT ONLY")
                return record

            token = self.get_token("sharepoint")
            digest = self.get_request_digest(site_url, token)

            repair_user = self.ensure_user(site_url, token, digest, self.config["repair_account"])
            repair_user_id = str(repair_user.get("Id"))
            record.repair_account_user_id = repair_user_id
            self.set_site_admin(site_url, token, digest, repair_user_id, True)
            self.log_site_step(site_url, f"Repair account ensured and elevated: {self.config['repair_account']}")

            if not record.current_user_id:
                raise RuntimeError("Owner user ID missing from _api/site/owner response.")

            self.set_site_admin(site_url, token, digest, record.current_user_id, False)
            self.log_site_step(site_url, f"Owner site collection admin set to false: {record.owner_upn}")

            self.remove_user_by_id(site_url, token, digest, record.current_user_id)
            self.log_site_step(site_url, f"Owner removed from OneDrive: {record.owner_upn}")
            time.sleep(self.config.get("sleep_after_remove_seconds", 2))

            readded_user = self.ensure_user(site_url, token, digest, owner_upn)
            record.readded_user_id = str(readded_user.get("Id"))
            self.log_site_step(site_url, f"Owner re-added to OneDrive: {record.owner_upn}")

            if self.config.get("readded_user_site_admin", True):
                self.set_site_admin(site_url, token, digest, record.readded_user_id, True)
                self.log_site_step(site_url, f"Owner granted site admin after re-add: {record.owner_upn}")

            try:
                self.set_site_admin(site_url, token, digest, repair_user_id, False)
                self.log_site_step(site_url, f"Repair account site collection admin set to false: {self.config['repair_account']}")
                self.remove_user_by_id(site_url, token, digest, repair_user_id)
                record.repair_account_cleanup_status = "removed"
                self.log_site_step(site_url, f"Repair account removed from OneDrive: {self.config['repair_account']}")
            except Exception as cleanup_exc:
                record.repair_account_cleanup_status = "error"
                record.repair_account_cleanup_message = str(cleanup_exc)
                self.log_site_step(
                    site_url,
                    f"Repair account cleanup failed: {record.repair_account_cleanup_message}",
                )

            verified_user = self.get_site_owner_info(site_url)
            record.verified_nameid = verified_user.get("current_nameid") if verified_user else None
            record.verified_match = record.verified_nameid == record.reference_nameid if record.reference_nameid else False
            record.action_status = "resolved" if record.verified_match else "readd_complete_unverified"
            record.message = (
                "Mismatch repaired and verified."
                if record.verified_match
                else "User was re-added, but post-check did not confirm the expected NameId."
            )
            self.log_site_step(site_url, f"Verified NameId: {record.verified_nameid or 'N/A'}")
            self.log_site_step(site_url, f"Status: {record.action_status.upper()}")
            return record
        except Exception as exc:
            record.action_status = "error"
            record.message = str(exc)
            self.logger.exception("[%s] Repair failed", site_url)
            return record


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


def write_json(path: Path, payload: Any) -> None:
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, indent=2)


def write_csv(path: Path, rows: List[Dict[str, Any]]) -> None:
    if not rows:
        with open(path, "w", encoding="utf-8", newline="") as handle:
            handle.write("")
        return

    fieldnames: List[str] = []
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


def build_report_dir(report_root: Path, report_date: str) -> Path:
    report_dir = report_root / report_date
    ensure_directory(report_dir)
    return report_dir


def summarize(records: List[RepairRecord]) -> Dict[str, int]:
    summary = {
        "total_sites": len(records),
        "mismatched_sites": 0,
        "resolved": 0,
        "report_only": 0,
        "already_match": 0,
        "not_found": 0,
        "errors": 0,
    }
    for record in records:
        if record.current_nameid and record.reference_nameid and not record.nameid_match:
            summary["mismatched_sites"] += 1
        if record.action_status == "resolved":
            summary["resolved"] += 1
        elif record.action_status == "report_only":
            summary["report_only"] += 1
        elif record.action_status == "already_match":
            summary["already_match"] += 1
        elif record.action_status == "not_found":
            summary["not_found"] += 1
        elif record.action_status == "error":
            summary["errors"] += 1
    return summary


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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Detect PUID mismatches for newly created OneDrive sites and optionally remove/re-add the owner."
    )
    parser.add_argument("--config", help="Path to JSON config file.", default=None)
    parser.add_argument("--date", help="Report date in YYYY-MM-DD. Default is yesterday in UTC.", default=None)
    parser.add_argument("--max-workers", type=int, default=None, help="Parallel workers for processing sites.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config = load_config(args.config)
    run_mode = normalize_run_mode(config.get("run_mode"))
    apply_changes = run_mode == "apply"

    report_date = args.date or (datetime.now(timezone.utc) - timedelta(days=1)).strftime("%Y-%m-%d")
    start_utc = datetime.fromisoformat(f"{report_date}T00:00:00+00:00")
    end_utc = start_utc + timedelta(days=1)

    report_dir = build_report_dir(Path(config["report_root"]), report_date)
    logger = setup_logger(report_dir / "run.log")

    logger.info("Starting OneDrive PUID repair job for %s", report_date)
    logger.info("Mode: %s", run_mode)

    client = Microsoft365RepairClient(config, logger)

    try:
        sites = client.discover_recent_onedrives(start_utc, end_utc)
        write_json(report_dir / "discovered_sites.json", sites)
        write_csv(
            report_dir / "discovered_sites.csv",
            [
                {
                    "run_mode": run_mode,
                    "owner_upn": site["owner_upn"],
                    "site_url": site["site_url"],
                    "site_created": site["site_created"],
                    "site_title": site["site_title"],
                    "site_id": site["site_id"],
                }
                for site in sites
            ],
        )

        logger.info("Discovered %s OneDrive sites created on %s", len(sites), report_date)

        records = process_sites(
            client,
            sites,
            apply_changes=apply_changes,
            max_workers=args.max_workers or config.get("max_workers", 5),
        )

        record_rows = [asdict(record) for record in records]
        write_json(report_dir / "repair_results.json", record_rows)
        write_csv(report_dir / "repair_results.csv", record_rows)

        puid_rows = [
            {
                "run_mode": record.run_mode,
                "owner_upn": record.owner_upn,
                "owner_login_name": record.owner_login_name,
                "owner_title": record.owner_title,
                "site_url": record.site_url,
                "site_created": record.site_created,
                "current_user_id": record.current_user_id,
                "current_nameid": record.current_nameid,
                "reference_nameid": record.reference_nameid,
                "reference_user_id": record.reference_user_id,
                "reference_cleanup_status": record.reference_cleanup_status,
                "reference_cleanup_message": record.reference_cleanup_message,
                "nameid_match": record.nameid_match,
                "verified_nameid": record.verified_nameid,
                "verified_match": record.verified_match,
                "repair_account_user_id": record.repair_account_user_id,
                "repair_account_cleanup_status": record.repair_account_cleanup_status,
                "repair_account_cleanup_message": record.repair_account_cleanup_message,
                "action_status": record.action_status,
            }
            for record in records
        ]
        write_csv(report_dir / "puid_comparison.csv", puid_rows)

        logger.info("Completed processing %s site(s)", len(records))

        summary = summarize(records)
        summary["run_mode"] = run_mode
        write_json(report_dir / "summary.json", summary)
        logger.info("Summary: %s", json.dumps(summary))
        logger.info("Reports written to %s", report_dir)
        return 0
    except Exception:
        logger.exception("Job failed")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
