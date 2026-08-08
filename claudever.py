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
import itertools
import os
import re
import sqlite3
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

import requests

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
SCOPE_APP = "https://graph.microsoft.com/.default"

DEFAULT_RETENTION_THRESHOLDS = [50, 100]


# ==========================================================================
# ===== CONFIGURATION — EDIT THIS SECTION, THEN JUST RUN THE SCRIPT ======
# ==========================================================================
# Fill these in directly. Anything left as None/empty falls back to a
# --command-line-flag if you pass one (flags always win over the values
# below), but for normal day-to-day use you shouldn't need any flags at all —
# just edit here and run: python graph_file_versions_report.py

# --- Azure AD app registration (required) ---
TENANT_ID = ""          # e.g. "72f988bf-XXXX-XXXX-XXXX-XXXXXXXXXXXX"
CLIENT_ID = ""          # Application (client) ID
CLIENT_SECRET = ""      # Client secret VALUE (not the secret ID)

# --- Pick exactly ONE of these four target modes ---
# 1) Bulk SharePoint sites, inline list:
SITE_URLS = ""          # e.g. "https://contoso.sharepoint.com/sites/Marketing,https://contoso.sharepoint.com/sites/HR"
# 2) Bulk SharePoint sites, from a text file (one URL per line, '#' comments ok):
SITE_URLS_FILE = ""     # e.g. "sites.txt"
# 3) A single specific drive id:
DRIVE_ID = ""           # e.g. "b!abc123..."
# 4) A single user's OneDrive:
USER_ID = ""            # e.g. "someone@contoso.com"

# --- Filters (leave blank / None to disable) ---
EXTENSIONS = ""         # e.g. ".pdf,.docx,.xlsx"  (blank = no extension filter)
MIN_SIZE = ""           # e.g. "5MB"               (blank = no minimum)
MAX_SIZE = ""           # e.g. "2GB"                (blank = no maximum)

# --- Version retention savings scenarios ---
RETENTION_THRESHOLDS = "50,100"   # "keep only the N most recent old versions" scenarios

# --- Where reports get written ---
OUTPUT_DIR = "./reports"

# --- Speed / concurrency ---
# Fetching version history is 1 API call per file, which is the slow part on
# large drives. Two things speed this up:
#  1) Requests are bundled into Graph's $batch endpoint, GRAPH_BATCH_SIZE
#     files per HTTP call (20 is Graph's hard maximum per batch — don't raise it).
#  2) CONCURRENCY batch calls run in parallel via a thread pool, instead of
#     one at a time, to overlap network latency.
# Effective files "in flight" at once = CONCURRENCY * GRAPH_BATCH_SIZE.
# Raise CONCURRENCY for more speed; lower it if you see a lot of throttling
# (HTTP 429) messages in the output.
CONCURRENCY = 5            # number of parallel $batch HTTP calls
GRAPH_BATCH_SIZE = 20       # Graph's hard limit per $batch request — do not increase

# ==========================================================================
# ===== END CONFIGURATION — nothing below this needs editing ============
# ==========================================================================


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

    def batch_post(self, requests_payload):
        """POST to Graph's $batch endpoint. requests_payload is the list that
        goes under the "requests" key (max 20 entries per Graph's limit)."""
        url = f"{GRAPH_ROOT}/$batch"
        while True:
            resp = self.session.post(url, json={"requests": requests_payload})
            if resp.status_code == 429:
                retry_after = int(resp.headers.get("Retry-After", "5"))
                print(f"  Throttled (batch), waiting {retry_after}s...", file=sys.stderr)
                time.sleep(retry_after)
                continue
            if resp.status_code >= 500:
                print(f"  Server error {resp.status_code} (batch), retrying in 5s...", file=sys.stderr)
                time.sleep(5)
                continue
            resp.raise_for_status()
            return resp.json()


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
def walk_drive_items(client, drive_id, extensions=None, min_size=None, max_size=None):
    """Iteratively (NOT recursively) walk every FILE in the drive that passes
    the extension/size filters. Uses an explicit stack instead of function
    recursion so it can't hit Python's recursion limit on deep folder trees,
    and never holds more than one folder's listing in memory at a time."""
    params = {"$select": "id,name,file,folder,size,webUrl,lastModifiedDateTime"}
    stack = [("root", "")]
    while stack:
        item_id, path = stack.pop()
        url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/children"
        for item in client.get_paged(url, params=params):
            item_path = f"{path}/{item['name']}"
            if "folder" in item:
                stack.append((item["id"], item_path))
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


def sort_versions(versions):
    """Sort a list of Graph version objects newest-first by lastModifiedDateTime."""
    def sort_key(v):
        ts = v.get("lastModifiedDateTime")
        if ts:
            try:
                return datetime.fromisoformat(ts.replace("Z", "+00:00"))
            except ValueError:
                pass
        return datetime.min

    return sorted(versions, key=sort_key, reverse=True)


def get_versions_sorted(client, drive_id, item_id):
    """Fetch older versions (excludes current) for ONE item via a plain GET,
    sorted newest-first. Used as: (a) a fallback when a batched sub-request
    fails, and (b) to page past 200 versions on the rare file that has more
    than a single page of history (batch responses only return one page)."""
    url = f"{GRAPH_ROOT}/drives/{drive_id}/items/{item_id}/versions"
    try:
        versions = list(client.get_paged(url))
    except requests.HTTPError as e:
        print(f"  Warning: could not fetch versions for item {item_id}: {e}", file=sys.stderr)
        return []
    return sort_versions(versions)


def chunked(iterable, size):
    """Yield lists of up to `size` items from iterable, pulling lazily so the
    whole iterable is never materialized at once — keeps memory bounded even
    when walking a drive with a huge number of files."""
    it = iter(iterable)
    while True:
        chunk = list(itertools.islice(it, size))
        if not chunk:
            return
        yield chunk


def fetch_versions_for_chunk(client, drive_id, chunk):
    """chunk: list of (item, path) tuples, length <= GRAPH_BATCH_SIZE.
    Fetches version history for all of them in a SINGLE Graph $batch HTTP
    call instead of one call per file. Returns list of
    (item, path, versions_sorted) in the same order as the input chunk."""
    requests_payload = [
        {"id": str(i), "method": "GET", "url": f"/drives/{drive_id}/items/{item['id']}/versions"}
        for i, (item, _path) in enumerate(chunk)
    ]

    try:
        batch_result = client.batch_post(requests_payload)
    except requests.HTTPError as e:
        # Whole batch failed outright (rare) — fall back to per-item calls
        print(f"  Warning: batch request failed ({e}); falling back to individual calls for this chunk", file=sys.stderr)
        return [(item, path, get_versions_sorted(client, drive_id, item["id"])) for item, path in chunk]

    responses_by_id = {r["id"]: r for r in batch_result.get("responses", [])}

    results = []
    for i, (item, path) in enumerate(chunk):
        r = responses_by_id.get(str(i))

        if r is None:
            print(f"  Warning: no batch response for {path}; retrying individually", file=sys.stderr)
            results.append((item, path, get_versions_sorted(client, drive_id, item["id"])))
            continue

        status = r.get("status", 500)

        if status == 429:
            # Per-request throttling inside a batch — respect Retry-After if
            # present, then retry just this one file on its own.
            retry_after = int((r.get("headers") or {}).get("Retry-After", "5"))
            time.sleep(retry_after)
            results.append((item, path, get_versions_sorted(client, drive_id, item["id"])))
            continue

        if status >= 400:
            print(f"  Warning: versions fetch failed for {path} (status {status}); retrying individually", file=sys.stderr)
            results.append((item, path, get_versions_sorted(client, drive_id, item["id"])))
            continue

        body = r.get("body", {}) or {}
        versions = body.get("value", [])

        # Rare: a single file has more versions than fit in one page (~200).
        # Batch responses don't auto-follow nextLink, so page the remainder
        # with a normal (non-batched) call just for this one file.
        next_link = body.get("@odata.nextLink")
        if next_link:
            try:
                versions += list(client.get_paged(next_link))
            except requests.HTTPError as e:
                print(f"  Warning: could not page extra versions for {path}: {e}", file=sys.stderr)

        results.append((item, path, sort_versions(versions)))

    return results


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
def run_report_for_drive(client, drive_id, label, output_dir, extensions, min_size, max_size, thresholds, concurrency):
    """Walk one drive, write its own CSV, return a summary dict.

    Speed for large drives: fetching version history is the slow part (it's
    normally 1 API call per file). Files are processed in chunks of
    GRAPH_BATCH_SIZE and each chunk's version history is fetched with a
    SINGLE Graph $batch call instead of one call per file — up to a 20x cut
    in HTTP round-trips. On top of that, `concurrency` chunks are fetched in
    parallel via a thread pool to overlap network latency. Effective files
    "in flight" at once = concurrency * GRAPH_BATCH_SIZE.

    Memory safety for very large drives (tens/hundreds of thousands of items):
    rows are NOT accumulated in a Python list. Each row is (a) written
    immediately to the live CSV (so you always have an up-to-date, valid file
    on disk even if the run is interrupted) and (b) inserted into a temporary
    on-disk SQLite database. Only the SQLite file — not Python memory — grows
    with item count. After the walk finishes, the final "sorted by version
    overhead" CSV is produced by streaming an ORDER BY query out of SQLite
    row-by-row, so peak memory stays roughly constant regardless of how many
    files are in the drive.
    """
    slug = slugify(label)
    filename = os.path.join(output_dir, f"{slug}.csv")
    db_path = os.path.join(output_dir, f".{slug}_tmp.sqlite")
    if os.path.exists(db_path):
        os.remove(db_path)

    print(f"\n=== {label} ===")
    print(f"Drive id: {drive_id}")
    print(f"Writing incrementally to: {filename}")
    print(f"Fetching versions in batches of {GRAPH_BATCH_SIZE}, {concurrency} batch(es) in parallel "
          f"(~{GRAPH_BATCH_SIZE * concurrency} files in flight at a time)")

    fieldnames = [
        "Path",
        "Name",
        "Extension",
        "CurrentSize",
        "CurrentSizeHuman",
        "CurrentModified",
        "VersionCount",
        "FirstVersionDate",
        "LastVersionDate",
        "VersionsSize",
        "VersionsSizeHuman",
        "TotalSize",
        "TotalSizeHuman",
    ]
    for n in thresholds:
        fieldnames.append(f"SavingsIfKeep{n}")
        fieldnames.append(f"SavingsIfKeep{n}Human")
    fieldnames.append("WebUrl")

    # Text columns vs integer columns, for the SQLite schema
    int_cols = {"CurrentSize", "VersionCount", "VersionsSize", "TotalSize"}
    int_cols |= {f"SavingsIfKeep{n}" for n in thresholds}

    db = sqlite3.connect(db_path)
    db.execute("PRAGMA journal_mode=OFF")   # we don't need crash-safety on the temp db itself
    db.execute("PRAGMA synchronous=OFF")
    col_defs = ", ".join(f'"{c}" {"INTEGER" if c in int_cols else "TEXT"}' for c in fieldnames)
    db.execute(f"CREATE TABLE rows ({col_defs})")
    insert_sql = f'INSERT INTO rows ({", ".join(f"[{c}]" for c in fieldnames)}) VALUES ({", ".join("?" for _ in fieldnames)})'

    totals = {
        "file_count": 0,
        "current_size": 0,
        "versions_size": 0,
        "savings": {n: 0 for n in thresholds},
    }

    SQLITE_BATCH_SIZE = 200
    sqlite_batch = []

    file_iter = walk_drive_items(client, drive_id, extensions=extensions, min_size=min_size, max_size=max_size)
    graph_chunks = chunked(file_iter, GRAPH_BATCH_SIZE)  # each is <= 20 (item, path) tuples

    with open(filename, "w", newline="", encoding="utf-8") as f, \
         ThreadPoolExecutor(max_workers=concurrency) as executor:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        f.flush()

        # Process `concurrency` chunks (i.e. concurrency * GRAPH_BATCH_SIZE
        # files) at a time: submit them all in parallel, wait for the group,
        # write results, move on. This bounds memory to one group's worth of
        # data while still overlapping network latency across the group.
        for group in chunked(graph_chunks, concurrency):
            futures = [executor.submit(fetch_versions_for_chunk, client, drive_id, chunk) for chunk in group]

            for future in futures:
                chunk_results = future.result()  # list of (item, path, versions_sorted)

                for item, path, versions_sorted in chunk_results:
                    totals["file_count"] += 1
                    current_size = item.get("size", 0)
                    version_count = len(versions_sorted)
                    versions_size = sum(v.get("size", 0) for v in versions_sorted)
                    savings = compute_retention_savings(versions_sorted, thresholds)

                    # versions_sorted is newest-first: index 0 = most recent
                    # old version, index -1 = oldest old version on record
                    first_version_date = versions_sorted[-1].get("lastModifiedDateTime", "") if versions_sorted else ""
                    last_version_date = versions_sorted[0].get("lastModifiedDateTime", "") if versions_sorted else ""

                    row = {
                        "Path": path,
                        "Name": item["name"],
                        "Extension": os.path.splitext(item["name"])[1].lower(),
                        "CurrentSize": current_size,
                        "CurrentSizeHuman": human_size(current_size),
                        "CurrentModified": item.get("lastModifiedDateTime", ""),
                        "VersionCount": version_count,
                        "FirstVersionDate": first_version_date,
                        "LastVersionDate": last_version_date,
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

                    sqlite_batch.append(tuple(row[c] for c in fieldnames))
                    if len(sqlite_batch) >= SQLITE_BATCH_SIZE:
                        db.executemany(insert_sql, sqlite_batch)
                        db.commit()
                        sqlite_batch.clear()

                    totals["current_size"] += current_size
                    totals["versions_size"] += versions_size
                    for n in thresholds:
                        totals["savings"][n] += savings[n]

            f.flush()
            if totals["file_count"] and totals["file_count"] % 100 < GRAPH_BATCH_SIZE * concurrency:
                print(f"  ...{totals['file_count']} files processed so far")

        if sqlite_batch:
            db.executemany(insert_sql, sqlite_batch)
            db.commit()

    # Final pass: stream rows back out of SQLite sorted by version storage
    # overhead (largest first). This never loads the full dataset into
    # Python memory at once — the cursor yields one row at a time.
    print("Sorting final report (streaming from disk, not memory)...")
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        cursor = db.execute(f'SELECT {", ".join(f"[{c}]" for c in fieldnames)} FROM rows ORDER BY "VersionsSize" DESC')
        while True:
            chunk = cursor.fetchmany(1000)
            if not chunk:
                break
            for record in chunk:
                writer.writerow(dict(zip(fieldnames, record)))

    db.close()
    os.remove(db_path)  # clean up the temp file

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
    parser = argparse.ArgumentParser(
        description="Report file version counts & storage via MS Graph (app-only auth). "
                     "Edit the CONFIGURATION block at the top of this file to avoid typing flags each run — "
                     "any flag passed here overrides the corresponding value in that block."
    )
    parser.add_argument("--tenant-id", default=TENANT_ID or None)
    parser.add_argument("--client-id", default=CLIENT_ID or None)
    parser.add_argument("--client-secret", default=CLIENT_SECRET or None)

    parser.add_argument("--site-urls", default=SITE_URLS or None, help="Comma-separated list of SharePoint site URLs")
    parser.add_argument("--site-urls-file", default=SITE_URLS_FILE or None, help="Text file, one site URL per line ('#' comments allowed)")
    parser.add_argument("--drive-id", default=DRIVE_ID or None, help="Explicit single drive id")
    parser.add_argument("--user-id", default=USER_ID or None, help="UPN or object id; uses that user's OneDrive")

    parser.add_argument("--extensions", default=EXTENSIONS or None, help="Comma-separated file extensions to include, e.g. '.pdf,.docx'")
    parser.add_argument("--min-size", default=MIN_SIZE or None, help="Minimum current file size, e.g. '5MB'")
    parser.add_argument("--max-size", default=MAX_SIZE or None, help="Maximum current file size, e.g. '2GB'")
    parser.add_argument(
        "--retention-thresholds",
        default=RETENTION_THRESHOLDS or ",".join(str(n) for n in DEFAULT_RETENTION_THRESHOLDS),
        help="Comma-separated list of 'keep most recent N old versions' scenarios to report savings for (default: 50,100)",
    )
    parser.add_argument("--output-dir", default=OUTPUT_DIR)
    parser.add_argument("--concurrency", type=int, default=CONCURRENCY,
                         help=f"Parallel $batch HTTP calls for fetching version history (default: {CONCURRENCY}). "
                              f"Lower this if you see throttling (429) messages.")

    args = parser.parse_args()

    # Validate required auth fields (may have come from the config block or a flag)
    missing_auth = [name for name, val in [("--tenant-id", args.tenant_id), ("--client-id", args.client_id),
                                            ("--client-secret", args.client_secret)] if not val]
    if missing_auth:
        parser.error(
            f"Missing required auth value(s): {', '.join(missing_auth)}. "
            f"Set them in the CONFIGURATION block at the top of the script, or pass them as flags."
        )

    # Validate exactly one target mode is set
    target_modes = {
        "--site-urls": args.site_urls,
        "--site-urls-file": args.site_urls_file,
        "--drive-id": args.drive_id,
        "--user-id": args.user_id,
    }
    chosen = [name for name, val in target_modes.items() if val]
    if len(chosen) == 0:
        parser.error(
            "No target specified. Set exactly ONE of SITE_URLS / SITE_URLS_FILE / DRIVE_ID / USER_ID "
            "in the CONFIGURATION block (or pass one of --site-urls/--site-urls-file/--drive-id/--user-id)."
        )
    if len(chosen) > 1:
        parser.error(f"Multiple target modes set ({', '.join(chosen)}) — set only ONE.")

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
    print(f"Concurrency: {args.concurrency} parallel batch call(s) of {GRAPH_BATCH_SIZE} files each")

    summaries = []
    for label, drive_id in targets:
        try:
            summary = run_report_for_drive(
                client, drive_id, label, args.output_dir, extensions, min_size, max_size, thresholds, args.concurrency
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
