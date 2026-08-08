#!/usr/bin/env python3
"""
graph_file_versions_report.py

Walk a OneDrive / SharePoint document library via Microsoft Graph API and
produce a report of every file with:
  - version count
  - space occupied by old versions
  - current file size
  - total space (current + versions)

Requires: pip install requests   (no msal — pure REST calls to the token endpoint)

--------------------------------------------------------------------------
AUTH
--------------------------------------------------------------------------
Two auth modes are supported:

1) App-only (client credentials) - best for unattended/admin scripts.
   Needs an Azure AD App Registration with an Application permission of
   Files.Read.All (OneDrive) and/or Sites.Read.All (SharePoint), admin-consented.

2) Delegated (device code) - runs under the signed-in user's own permissions,
   no admin consent needed for a "Files.Read.All" delegated scope on many
   tenants, but only sees what that user can see.

--------------------------------------------------------------------------
FINDING YOUR drive_id
--------------------------------------------------------------------------
- Your own OneDrive:      GET /me/drive                 -> id
- A SharePoint site drive: GET /sites/{site-id}/drives    -> pick the one you want
  (site-id can be resolved via GET /sites/{hostname}:/sites/{sitepath})

You can also just pass --user-id (UPN or object id) to hit that user's OneDrive
without knowing the drive id up front, or --site-url to resolve a SharePoint
site's default document library automatically.

--------------------------------------------------------------------------
USAGE EXAMPLES
--------------------------------------------------------------------------
# App-only, a specific drive id, output to CSV
python graph_file_versions_report.py \\
    --auth app \\
    --tenant-id <TENANT_ID> --client-id <CLIENT_ID> --client-secret <SECRET> \\
    --drive-id b!abc123... \\
    --output report.csv

# App-only, a user's OneDrive by UPN
python graph_file_versions_report.py \\
    --auth app \\
    --tenant-id <TENANT_ID> --client-id <CLIENT_ID> --client-secret <SECRET> \\
    --user-id someone@contoso.com \\
    --output report.csv

# Delegated device-code login, a SharePoint site by URL
python graph_file_versions_report.py \\
    --auth delegated \\
    --tenant-id <TENANT_ID> --client-id <CLIENT_ID> \\
    --site-url https://contoso.sharepoint.com/sites/Marketing \\
    --output report.csv
"""

import argparse
import csv
import sys
import time
from datetime import datetime

import requests

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
SCOPE_APP = "https://graph.microsoft.com/.default"
SCOPE_DELEGATED = "Files.Read.All Sites.Read.All offline_access"


# --------------------------------------------------------------------------
# Auth (raw REST calls to the Microsoft identity platform token endpoint —
# no msal dependency)
# --------------------------------------------------------------------------
def get_token_app(tenant_id, client_id, client_secret):
    """OAuth2 client credentials flow. Needs an Application permission
    (Files.Read.All / Sites.Read.All) admin-consented on the app registration."""
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": SCOPE_APP,
        "grant_type": "client_credentials",
    }
    resp = requests.post(url, data=data)
    body = resp.json()
    if resp.status_code != 200 or "access_token" not in body:
        raise RuntimeError(f"Auth failed: {body.get('error_description', body)}")
    return body["access_token"]


def get_token_delegated(tenant_id, client_id):
    """OAuth2 device code flow, done as raw REST calls (no msal).
    Requires the app registration to allow the 'public client' / mobile &
    desktop flow (Authentication -> Advanced settings -> Allow public
    client flows = Yes)."""
    device_code_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/devicecode"
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    resp = requests.post(
        device_code_url,
        data={"client_id": client_id, "scope": SCOPE_DELEGATED},
    )
    flow = resp.json()
    if "device_code" not in flow:
        raise RuntimeError(f"Failed to start device code flow: {flow}")

    print(flow["message"])  # tells user to browse to microsoft.com/devicelogin

    interval = flow.get("interval", 5)
    expires_in = flow.get("expires_in", 900)
    deadline = time.time() + expires_in

    while time.time() < deadline:
        time.sleep(interval)
        poll = requests.post(
            token_url,
            data={
                "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": client_id,
                "device_code": flow["device_code"],
            },
        )
        body = poll.json()
        if "access_token" in body:
            return body["access_token"]
        error = body.get("error")
        if error == "authorization_pending":
            continue
        if error == "slow_down":
            interval += 5
            continue
        raise RuntimeError(f"Auth failed: {body.get('error_description', body)}")

    raise RuntimeError("Device code flow timed out before user completed sign-in.")


# --------------------------------------------------------------------------
# Graph helpers (with basic throttling / retry handling)
# --------------------------------------------------------------------------
class GraphClient:
    def __init__(self, token):
        self.session = requests.Session()
        self.session.headers.update(
            {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        )

    def get(self, url, params=None):
        while True:
            resp = self.session.get(url, params=params)
            if resp.status_code == 429:
                retry_after = int(resp.headers.get("Retry-After", "5"))
                print(f"  Throttled, waiting {retry_after}s...", file=sys.stderr)
                time.sleep(retry_after)
                continue
            if resp.status_code >= 500:
                print(f"  Server error {resp.status_code}, retrying in 5s...", file=sys.stderr)
                time.sleep(5)
                continue
            resp.raise_for_status()
            return resp.json()

    def get_paged(self, url, params=None):
        """Yield items across all pages for a Graph collection endpoint."""
        next_url = url
        next_params = params
        while next_url:
            data = self.get(next_url, params=next_params)
            for item in data.get("value", []):
                yield item
            next_url = data.get("@odata.nextLink")
            next_params = None  # nextLink already has query params baked in


# --------------------------------------------------------------------------
# Drive resolution helpers
# --------------------------------------------------------------------------
def resolve_drive_id_for_user(client, user_id):
    data = client.get(f"{GRAPH_ROOT}/users/{user_id}/drive")
    return data["id"]


def resolve_drive_id_for_site(client, site_url):
    # site_url like https://contoso.sharepoint.com/sites/Marketing
    parts = site_url.replace("https://", "").split("/", 1)
    hostname = parts[0]
    site_path = "/" + parts[1] if len(parts) > 1 else ""
    data = client.get(f"{GRAPH_ROOT}/sites/{hostname}:{site_path}")
    site_id = data["id"]
    drive = client.get(f"{GRAPH_ROOT}/sites/{site_id}/drive")
    return drive["id"]


# --------------------------------------------------------------------------
# Core walk + version lookup
# --------------------------------------------------------------------------
def walk_drive_items(client, drive_id, item_id="root", path=""):
    """Recursively yield (item_dict, full_path) for every FILE in the drive."""
    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/children"
    params = {"$select": "id,name,file,folder,size,webUrl"}
    for item in client.get_paged(url, params=params):
        item_path = f"{path}/{item['name']}"
        if "folder" in item:
            yield from walk_drive_items(client, drive_id, item["id"], item_path)
        elif "file" in item:
            yield item, item_path


def get_versions_info(client, drive_id, item_id):
    """Return (version_count, versions_total_size) for older versions of a file."""
    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/versions"
    try:
        versions = list(client.get_paged(url))
    except requests.HTTPError as e:
        # Some file types (e.g. certain non-Office files) may not support versioning
        print(f"  Warning: could not fetch versions for item {item_id}: {e}", file=sys.stderr)
        return 0, 0
    total_size = sum(v.get("size", 0) for v in versions)
    return len(versions), total_size


def human_size(num_bytes):
    step = 1024.0
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if abs(num_bytes) < step:
            return f"{num_bytes:.1f}{unit}"
        num_bytes /= step
    return f"{num_bytes:.1f}PB"


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Report file version counts & storage via MS Graph")
    parser.add_argument("--auth", choices=["app", "delegated"], default="app")
    parser.add_argument("--tenant-id", required=True)
    parser.add_argument("--client-id", required=True)
    parser.add_argument("--client-secret", help="Required for --auth app")

    drive_group = parser.add_mutually_exclusive_group(required=True)
    drive_group.add_argument("--drive-id", help="Explicit drive id")
    drive_group.add_argument("--user-id", help="UPN or object id; uses that user's OneDrive")
    drive_group.add_argument("--site-url", help="SharePoint site URL; uses its default document library")

    parser.add_argument("--output", default="file_versions_report.csv")
    args = parser.parse_args()

    if args.auth == "app":
        if not args.client_secret:
            parser.error("--client-secret is required for --auth app")
        token = get_token_app(args.tenant_id, args.client_id, args.client_secret)
    else:
        token = get_token_delegated(args.tenant_id, args.client_id)

    client = GraphClient(token)

    if args.drive_id:
        drive_id = args.drive_id
    elif args.user_id:
        drive_id = resolve_drive_id_for_user(client, args.user_id)
    else:
        drive_id = resolve_drive_id_for_site(client, args.site_url)

    print(f"Using drive: {drive_id}")
    print("Walking drive and collecting version info (this can take a while for large drives)...")

    rows = []
    grand_total_current = 0
    grand_total_versions = 0
    file_count = 0

    for item, path in walk_drive_items(client, drive_id):
        file_count += 1
        current_size = item.get("size", 0)
        version_count, versions_size = get_versions_info(client, drive_id, item["id"])

        rows.append(
            {
                "Path": path,
                "Name": item["name"],
                "CurrentSize": current_size,
                "CurrentSizeHuman": human_size(current_size),
                "VersionCount": version_count,
                "VersionsSize": versions_size,
                "VersionsSizeHuman": human_size(versions_size),
                "TotalSize": current_size + versions_size,
                "TotalSizeHuman": human_size(current_size + versions_size),
                "WebUrl": item.get("webUrl", ""),
            }
        )
        grand_total_current += current_size
        grand_total_versions += versions_size

        if file_count % 25 == 0:
            print(f"  ...{file_count} files processed so far")

    # sort largest version overhead first — usually what people care about
    rows.sort(key=lambda r: r["VersionsSize"], reverse=True)

    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "Path",
                "Name",
                "CurrentSize",
                "CurrentSizeHuman",
                "VersionCount",
                "VersionsSize",
                "VersionsSizeHuman",
                "TotalSize",
                "TotalSizeHuman",
                "WebUrl",
            ],
        )
        writer.writeheader()
        writer.writerows(rows)

    print()
    print(f"Done. {file_count} files scanned.")
    print(f"Current-version storage total: {human_size(grand_total_current)}")
    print(f"Old-version storage total:     {human_size(grand_total_versions)}")
    print(f"Grand total storage:           {human_size(grand_total_current + grand_total_versions)}")
    print(f"Report written to: {args.output}  ({datetime.now().isoformat(timespec='seconds')})")


if __name__ == "__main__":
    main()
