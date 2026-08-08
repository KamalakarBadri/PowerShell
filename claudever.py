#!/usr/bin/env python3
"""
graph_file_versions_report.py

Walk one or many SharePoint sites (or a specific drive / a user's OneDrive)
via Microsoft Graph REST API and produce, per site/library, a CSV report of
every file with:
  - version count
  - space occupied by old versions
  - current file size / total size
  - estimated space freed if you trimmed version history down to the most
    recent N versions (configurable thresholds, default 50 and 100)

Auth: app-only (client credentials) ONLY. No delegated / device-code flow.
Dependency: pip install requests   (no msal)

--------------------------------------------------------------------------
AZURE AD APP REGISTRATION
--------------------------------------------------------------------------
Create an App Registration, add an Application permission (not Delegated):
  - Sites.Read.All   (for SharePoint sites)
  - Files.Read.All   (for OneDrive/user drives, if you use --user-id)
Grant admin consent. Create a client secret. Use tenant id / client id /
client secret with this script.

--------------------------------------------------------------------------
BULK SITES
--------------------------------------------------------------------------
Provide multiple SharePoint sites via:
  --site-urls "https://contoso.sharepoint.com/sites/Marketing,https://contoso.sharepoint.com/sites/HR"
or
  --site-urls-file sites.txt      (one site URL per line, '#' comments allowed)

Each site can have multiple document libraries (drives). By default the
script reports on ALL document libraries in each site. A separate CSV is
written per site+library, plus one overall summary.csv across everything.

--------------------------------------------------------------------------
FILTERS
--------------------------------------------------------------------------
  --extensions ".pdf,.docx,.xlsx"     only include files with these extensions
  --min-size 10MB                     only include files at/above this size
  --max-size 2GB                      only include files at/below this size
  (sizes accept plain bytes or suffixes: KB, MB, GB, TB)

--------------------------------------------------------------------------
VERSION RETENTION SAVINGS
--------------------------------------------------------------------------
  --retention-thresholds 50,100       (default) report estimated bytes freed
                                       if only the N most recent OLD versions
                                       were kept and everything older than
                                       that was deleted. One column per
                                       threshold is added to the report,
                                       e.g. SavingsIfKeep50, SavingsIfKeep100.
  This is an ESTIMATE from metadata only — the script does not delete
  anything.

--------------------------------------------------------------------------
USAGE EXAMPLES
--------------------------------------------------------------------------
# Bulk sites from a file, default filters, default thresholds
python graph_file_versions_report.py \\
    --tenant-id <TENANT_ID> --client-id <CLIENT_ID> --client-secret <SECRET> \\
    --site-urls-file sites.txt \\
    --output-dir ./reports

# Single site, only PDFs/Office docs over 5MB, custom thresholds
python graph_file_versions_report.py \\
    --tenant-id <TENANT_ID> --client-id <CLIENT_ID> --client-secret <SECRET> \\
    --site-urls "https://contoso.sharepoint.com/sites/Marketing" \\
    --extensions ".pdf,.docx,.pptx,.xlsx" --min-size 5MB \\
    --retention-thresholds 25,50,100 \\
    --output-dir ./reports

# A specific drive id (no site resolution)
python graph_file_versions_report.py \\
    --tenant-id <TENANT_ID> --client-id <CLIENT_ID> --client-secret <SECRET> \\
    --drive-id b!abc123... \\
    --output-dir ./reports

# A specific user's OneDrive
python graph_file_versions_report.py \\
    --tenant-id <TENANT_ID> --client-id <CLIENT_ID> --client-secret <SECRET> \\
    --user-id someone@contoso.com \\
    --output-dir ./reports
"""

import argparse
import csv
import os
import re
import sys
import time
from datetime import datetime

import requests

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
SCOPE_APP = "https://graph.microsoft.com/.default"

DEFAULT_RETENTION_THRESHOLDS = [50, 100]


# --------------------------------------------------------------------------
# Auth (app-only / client credentials, raw REST — no msal)
# --------------------------------------------------------------------------
def get_token_app(tenant_id, client_id, client_secret):
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


# --------------------------------------------------------------------------
# Graph helpers (paging + throttling/retry)
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
        next_url = url
        next_params = params
        while next_url:
            data = self.get(next_url, params=next_params)
            for item in data.get("value", []):
                yield item
            next_url = data.get("@odata.nextLink")
            next_params = None  # nextLink already has query params baked in


# --------------------------------------------------------------------------
# Size parsing / formatting
# --------------------------------------------------------------------------
SIZE_UNITS = {"B": 1, "KB": 1024, "MB": 1024**2, "GB": 1024**3, "TB": 1024**4}


def parse_size(size_str):
    """Parse '10MB', '500 KB', '2GB', or a plain integer (bytes) into bytes."""
    if size_str is None:
        return None
    size_str = str(size_str).strip().upper()
    match = re.match(r"^([\d.]+)\s*([A-Z]*)$", size_str)
    if not match:
        raise ValueError(f"Cannot parse size: {size_str}")
    number, unit = match.groups()
    unit = unit if unit else "B"
    if unit not in SIZE_UNITS:
        raise ValueError(f"Unknown size unit '{unit}' in '{size_str}'")
    return int(float(number) * SIZE_UNITS[unit])


def human_size(num_bytes):
    step = 1024.0
    n = float(num_bytes)
    for unit in ["B", "KB", "MB", "GB", "TB"]:
        if abs(n) < step:
            return f"{n:.1f}{unit}"
        n /= step
    return f"{n:.1f}PB"


def slugify(text):
    text = re.sub(r"[^\w\-]+", "_", text.strip())
    return re.sub(r"_+", "_", text).strip("_") or "unnamed"


# --------------------------------------------------------------------------
# Site / drive resolution
# --------------------------------------------------------------------------
def resolve_site(client, site_url):
    """Return (site_id, site_display_name) for a SharePoint site URL."""
    parts = site_url.replace("https://", "").replace("http://", "").split("/", 1)
    hostname = parts[0]
    site_path = "/" + parts[1] if len(parts) > 1 else ""
    data = client.get(f"{GRAPH_ROOT}/sites/{hostname}:{site_path}")
    return data["id"], data.get("displayName") or data.get("name") or site_path.strip("/") or hostname


def list_drives_for_site(client, site_id):
    """Return list of {id, name} document libraries (drives) for a site."""
    url = f"{GRAPH_ROOT}/sites/{site_id}/drives"
    return [{"id": d["id"], "name": d.get("name", "Documents")} for d in client.get_paged(url)]


def resolve_drive_id_for_user(client, user_id):
    data = client.get(f"{GRAPH_ROOT}/users/{user_id}/drive")
    return data["id"]


# --------------------------------------------------------------------------
# Core walk + version lookup
# --------------------------------------------------------------------------
def walk_drive_items(client, drive_id, item_id="root", path="", extensions=None, min_size=None, max_size=None):
    """Recursively yield (item_dict, full_path) for every FILE in the drive
    that passes the extension/size filters. Folders are always traversed
    regardless of filters (filters only apply to files)."""
    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/children"
    params = {"$select": "id,name,file,folder,size,webUrl"}
    for item in client.get_paged(url, params=params):
        item_path = f"{path}/{item['name']}"
        if "folder" in item:
            yield from walk_drive_items(
                client, drive_id, item["id"], item_path, extensions, min_size, max_size
            )
        elif "file" in item:
            if extensions:
                ext = os.path.splitext(item["name"])[1].lower()
                if ext not in extensions:
                    continue
            size = item.get("size", 0)
            if min_size is not None and size < min_size:
                continue
            if max_size is not None and size > max_size:
                continue
            yield item, item_path


def get_versions_sorted(client, drive_id, item_id):
    """Return older versions (excludes current) sorted newest-first, with size."""
    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/versions"
    try:
        versions = list(client.get_paged(url))
    except requests.HTTPError as e:
        print(f"  Warning: could not fetch versions for item {item_id}: {e}", file=sys.stderr)
        return []

    def sort_key(v):
        ts = v.get("lastModifiedDateTime")
        if ts:
            try:
                return datetime.fromisoformat(ts.replace("Z", "+00:00"))
            except ValueError:
                pass
        return datetime.min

    return sorted(versions, key=sort_key, reverse=True)


def compute_retention_savings(versions_sorted, thresholds):
    """For each threshold N, bytes freed by keeping only the N most recent
    old versions and deleting anything older than that."""
    savings = {}
    for n in thresholds:
        if len(versions_sorted) > n:
            savings[n] = sum(v.get("size", 0) for v in versions_sorted[n:])
        else:
            savings[n] = 0
    return savings


# --------------------------------------------------------------------------
# Per-drive report
# --------------------------------------------------------------------------
def run_report_for_drive(client, drive_id, label, output_dir, extensions, min_size, max_size, thresholds):
    """Walk one drive, write its own CSV, return a summary dict."""
    filename = os.path.join(output_dir, f"{slugify(label)}.csv")
    print(f"\n=== {label} ===")
    print(f"Drive id: {drive_id}")
    print(f"Writing incrementally to: {filename}")

    fieldnames = [
        "Path",
        "Name",
        "Extension",
        "CurrentSize",
        "CurrentSizeHuman",
        "VersionCount",
        "VersionsSize",
        "VersionsSizeHuman",
        "TotalSize",
        "TotalSizeHuman",
    ]
    for n in thresholds:
        fieldnames.append(f"SavingsIfKeep{n}")
        fieldnames.append(f"SavingsIfKeep{n}Human")
    fieldnames.append("WebUrl")

    rows = []
    totals = {
        "file_count": 0,
        "current_size": 0,
        "versions_size": 0,
        "savings": {n: 0 for n in thresholds},
    }

    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        f.flush()

        for item, path in walk_drive_items(
            client, drive_id, extensions=extensions, min_size=min_size, max_size=max_size
        ):
            totals["file_count"] += 1
            current_size = item.get("size", 0)
            versions_sorted = get_versions_sorted(client, drive_id, item["id"])
            version_count = len(versions_sorted)
            versions_size = sum(v.get("size", 0) for v in versions_sorted)
            savings = compute_retention_savings(versions_sorted, thresholds)

            row = {
                "Path": path,
                "Name": item["name"],
                "Extension": os.path.splitext(item["name"])[1].lower(),
                "CurrentSize": current_size,
                "CurrentSizeHuman": human_size(current_size),
                "VersionCount": version_count,
                "VersionsSize": versions_size,
                "VersionsSizeHuman": human_size(versions_size),
                "TotalSize": current_size + versions_size,
                "TotalSizeHuman": human_size(current_size + versions_size),
                "WebUrl": item.get("webUrl", ""),
            }
            for n in thresholds:
                row[f"SavingsIfKeep{n}"] = savings[n]
                row[f"SavingsIfKeep{n}Human"] = human_size(savings[n])

            writer.writerow(row)
            f.flush()
            rows.append(row)

            totals["current_size"] += current_size
            totals["versions_size"] += versions_size
            for n in thresholds:
                totals["savings"][n] += savings[n]

            if totals["file_count"] % 25 == 0:
                print(f"  ...{totals['file_count']} files processed so far (last: {path})")

    # Final pass: rewrite sorted by version storage overhead, largest first
    rows.sort(key=lambda r: r["VersionsSize"], reverse=True)
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Done: {totals['file_count']} files. "
          f"Current: {human_size(totals['current_size'])}, "
          f"Old versions: {human_size(totals['versions_size'])}")
    for n in thresholds:
        print(f"  Potential savings if keeping only last {n} old versions per file: "
              f"{human_size(totals['savings'][n])}")

    return {
        "label": label,
        "drive_id": drive_id,
        "csv_file": filename,
        **totals,
    }


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Report file version counts & storage via MS Graph (app-only auth)")
    parser.add_argument("--tenant-id", required=True)
    parser.add_argument("--client-id", required=True)
    parser.add_argument("--client-secret", required=True)

    target_group = parser.add_mutually_exclusive_group(required=True)
    target_group.add_argument("--site-urls", help="Comma-separated list of SharePoint site URLs")
    target_group.add_argument("--site-urls-file", help="Text file, one site URL per line ('#' comments allowed)")
    target_group.add_argument("--drive-id", help="Explicit single drive id")
    target_group.add_argument("--user-id", help="UPN or object id; uses that user's OneDrive")

    parser.add_argument("--extensions", help="Comma-separated file extensions to include, e.g. '.pdf,.docx'")
    parser.add_argument("--min-size", help="Minimum current file size, e.g. '5MB'")
    parser.add_argument("--max-size", help="Maximum current file size, e.g. '2GB'")
    parser.add_argument(
        "--retention-thresholds",
        default=",".join(str(n) for n in DEFAULT_RETENTION_THRESHOLDS),
        help="Comma-separated list of 'keep most recent N old versions' scenarios to report savings for (default: 50,100)",
    )
    parser.add_argument("--output-dir", default="./reports")

    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)

    extensions = None
    if args.extensions:
        extensions = {e.strip().lower() if e.strip().startswith(".") else f".{e.strip().lower()}"
                      for e in args.extensions.split(",") if e.strip()}

    min_size = parse_size(args.min_size) if args.min_size else None
    max_size = parse_size(args.max_size) if args.max_size else None
    thresholds = [int(t.strip()) for t in args.retention_thresholds.split(",") if t.strip()]

    token = get_token_app(args.tenant_id, args.client_id, args.client_secret)
    client = GraphClient(token)

    # Build the list of (label, drive_id) targets to report on
    targets = []

    if args.drive_id:
        targets.append((f"drive_{args.drive_id[:12]}", args.drive_id))

    elif args.user_id:
        drive_id = resolve_drive_id_for_user(client, args.user_id)
        targets.append((f"onedrive_{args.user_id}", drive_id))

    else:
        if args.site_urls:
            site_urls = [s.strip() for s in args.site_urls.split(",") if s.strip()]
        else:
            with open(args.site_urls_file, "r", encoding="utf-8") as f:
                site_urls = [line.strip() for line in f if line.strip() and not line.strip().startswith("#")]

        for site_url in site_urls:
            try:
                site_id, site_name = resolve_site(client, site_url)
            except requests.HTTPError as e:
                print(f"Skipping site (could not resolve): {site_url} -> {e}", file=sys.stderr)
                continue

            drives = list_drives_for_site(client, site_id)
            if not drives:
                print(f"No document libraries found for site: {site_url}", file=sys.stderr)
                continue

            for drive in drives:
                label = f"{site_name}__{drive['name']}"
                targets.append((label, drive["id"]))

    if not targets:
        print("No drives resolved — nothing to report on.", file=sys.stderr)
        sys.exit(1)

    print(f"Resolved {len(targets)} drive(s)/library(ies) to report on.")
    if extensions:
        print(f"Extension filter: {sorted(extensions)}")
    if min_size is not None:
        print(f"Min size filter: {human_size(min_size)}")
    if max_size is not None:
        print(f"Max size filter: {human_size(max_size)}")
    print(f"Retention savings thresholds: {thresholds}")

    summaries = []
    for label, drive_id in targets:
        try:
            summary = run_report_for_drive(
                client, drive_id, label, args.output_dir, extensions, min_size, max_size, thresholds
            )
            summaries.append(summary)
        except Exception as e:
            print(f"ERROR processing {label} ({drive_id}): {e}", file=sys.stderr)

    # Overall summary.csv across all sites/libraries
    summary_path = os.path.join(args.output_dir, "summary.csv")
    with open(summary_path, "w", newline="", encoding="utf-8") as f:
        fieldnames = ["Label", "DriveId", "CsvFile", "FileCount", "CurrentSize", "CurrentSizeHuman",
                      "VersionsSize", "VersionsSizeHuman"]
        for n in thresholds:
            fieldnames.append(f"SavingsIfKeep{n}")
            fieldnames.append(f"SavingsIfKeep{n}Human")
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for s in summaries:
            row = {
                "Label": s["label"],
                "DriveId": s["drive_id"],
                "CsvFile": s["csv_file"],
                "FileCount": s["file_count"],
                "CurrentSize": s["current_size"],
                "CurrentSizeHuman": human_size(s["current_size"]),
                "VersionsSize": s["versions_size"],
                "VersionsSizeHuman": human_size(s["versions_size"]),
            }
            for n in thresholds:
                row[f"SavingsIfKeep{n}"] = s["savings"][n]
                row[f"SavingsIfKeep{n}Human"] = human_size(s["savings"][n])
            writer.writerow(row)

    print(f"\nAll done. {len(summaries)} report(s) written to {args.output_dir}")
    print(f"Overall summary: {summary_path}")
    print(f"Finished: {datetime.now().isoformat(timespec='seconds')}")


if __name__ == "__main__":
    main()
