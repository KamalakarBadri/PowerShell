import requests
import json
from urllib.parse import urljoin, urlparse
import csv
from datetime import datetime, timezone
import pytz
import re
import uuid
import base64
import time
import os
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend

# ============ Configuration ============
OUTPUT_CSV = f"SharePointBulkContentWithCompliance_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"

TENANT_NAME = "test.onmicrosoft.com"
APP_ID = "73efa35d-6188-42d4-b258-838a977eb149"

CERTIFICATE_PATH = "certificate.pem"
PRIVATE_KEY_PATH = "private_key.pem"

UTC_TZ = timezone.utc
IST_TZ = pytz.timezone('Asia/Kolkata')

# Token cache: { site_host: {token, expires_at} }
_token_cache = {}
_token_lock = threading.Lock()  # Protect token cache during parallel access

# CSV Write lock
_csv_lock = threading.Lock()

# ============ Certificate / JWT / Auth ============
def load_certificate_and_key():
    try:
        with open(CERTIFICATE_PATH, "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
        with open(PRIVATE_KEY_PATH, "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())
        return certificate, private_key
    except Exception as e:
        print(f"Error loading certificate or private key: {e}")
        raise


def get_jwt_token(certificate, private_key, audience):
    now = int(time.time())
    expiration = now + 600
    thumbprint = certificate.fingerprint(hashes.SHA1())
    x5t = base64.urlsafe_b64encode(thumbprint).decode('utf-8').replace('=', '')
    jwt_header = {"alg": "RS256", "typ": "JWT", "x5t": x5t}
    jwt_payload = {"aud": audience, "exp": expiration, "iss": APP_ID, "jti": str(uuid.uuid4()), "nbf": now, "sub": APP_ID}

    encoded_header = base64.urlsafe_b64encode(json.dumps(jwt_header, separators=(',', ':')).encode()).decode().replace('=', '')
    encoded_payload = base64.urlsafe_b64encode(json.dumps(jwt_payload, separators=(',', ':')).encode()).decode().replace('=', '')
    jwt_unsigned = f"{encoded_header}.{encoded_payload}"

    signature = private_key.sign(jwt_unsigned.encode(), padding.PKCS1v15(), hashes.SHA256())
    encoded_signature = base64.urlsafe_b64encode(signature).decode().replace('=', '')
    return f"{jwt_unsigned}.{encoded_signature}"


def get_access_token_for_site(site_url):
    host = urlparse(site_url).netloc

    # Thread-safe cache check
    with _token_lock:
        cached = _token_cache.get(host)
        if cached and cached["expires_at"] > time.time() + 60:
            return cached["token"]

    scope = f"https://{host}/.default"
    audience = f"https://login.microsoftonline.com/{TENANT_NAME}/oauth2/v2.0/token"

    certificate, private_key = load_certificate_and_key()
    jwt = get_jwt_token(certificate, private_key, audience)

    data = {
        "client_id": APP_ID,
        "client_assertion": jwt,
        "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        "scope": scope,
        "grant_type": "client_credentials",
    }

    resp = requests.post(audience, data=data, headers={"Content-Type": "application/x-www-form-urlencoded"}, timeout=30)
    resp.raise_for_status()
    js = resp.json()

    with _token_lock:
        _token_cache[host] = {
            "token": js["access_token"],
            "expires_at": time.time() + int(js.get("expires_in", 3600)),
        }
    return js["access_token"]


# ============ SharePoint request helper ============
def make_sharepoint_request(url, access_token, method='GET', headers=None, retries=3):
    default_headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json",
    }
    if headers:
        default_headers.update(headers)

    delay = 2
    for attempt in range(retries):
        try:
            r = requests.request(method, url, headers=default_headers, timeout=90)
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", delay))
                # Don't print 429 warnings in parallel mode to avoid console spam, just wait
                time.sleep(wait)
                continue
            if 500 <= r.status_code < 600:
                time.sleep(delay)
                delay *= 2
                continue
            r.raise_for_status()
            return r.json()
        except requests.exceptions.RequestException as e:
            if hasattr(e, 'response') and e.response is not None and e.response.status_code == 404:
                return None
            time.sleep(delay)
            delay *= 2
    return None


# ============ Datetime helpers ============
def normalize_datetime(dt_str):
    if not dt_str or dt_str == "N/A": return None
    dt_str = re.sub(r'[+-]\d{2}:\d{2}$', '', dt_str)
    if not dt_str.endswith('Z'): dt_str += 'Z'
    return dt_str


def convert_utc_to_ist(utc_datetime_str):
    utc_datetime_str = normalize_datetime(utc_datetime_str)
    if not utc_datetime_str: return "N/A"
    try:
        if '.' in utc_datetime_str:
            utc_dt = datetime.strptime(utc_datetime_str, "%Y-%m-%dT%H:%M:%S.%fZ").replace(tzinfo=UTC_TZ)
        else:
            utc_dt = datetime.strptime(utc_datetime_str, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=UTC_TZ)
        return utc_dt.astimezone(IST_TZ).strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return "N/A"


# ============ SharePoint data retrieval ============
def get_all_lists(site_url, access_token):
    lists_url = urljoin(site_url + "/", "_api/web/lists")
    response = make_sharepoint_request(lists_url, access_token)
    if response and 'd' in response and 'results' in response['d']:
        return [lst for lst in response['d']['results'] if lst['BaseTemplate'] == 101]
    return []


def get_list_items(site_url, list_id, access_token):
    items_url = f"{site_url}/_api/web/lists(guid'{list_id}')/items?$expand=File,Folder&$top=5000"
    all_items, next_url, page = [], items_url, 0
    while next_url:
        response = make_sharepoint_request(next_url, access_token)
        if response is None:
            raise RuntimeError(f"Pagination failed on page {page} for list {list_id}")
        if 'd' in response and 'results' in response['d']:
            batch = response['d']['results']
            all_items.extend(batch)
            print(f"      page {page}: +{len(batch)} (total {len(all_items)})")
        page += 1
        next_url = response.get('d', {}).get('__next')
    return all_items


def get_file_properties(site_url, list_id, item_id, access_token):
    props_url = f"{site_url}/_api/web/lists(guid'{list_id}')/items({item_id})/file/properties"
    response = make_sharepoint_request(props_url, access_token)
    return response.get('d', {}) if response else {}


def get_folder_field_values(site_url, list_id, item_id, access_token):
    url = f"{site_url}/_api/web/lists(guid'{list_id}')/items({item_id})/FieldValuesAsText"
    response = make_sharepoint_request(url, access_token)
    return response.get('d', {}) if response else {}


def process_item(site_url, list_id, item, access_token):
    item_type = item['FileSystemObjectType']
    details = {
        'Type': 'File' if item_type == 0 else 'Folder',
        'ID': item['Id'],
        'Name': '', 'Path': '',
        'Size': 0 if item_type == 0 else 'N/A',
        'Created': item.get('Created', 'N/A'),
        'Modified': item.get('Modified', 'N/A'),
        'Author': item.get('Author', {}).get('Title', 'N/A'),
        'Editor': item.get('Editor', {}).get('Title', 'N/A'),
        'ComplianceTag': 'N/A', 'ComplianceTagWrittenTime_UTC': 'N/A', 'ComplianceTagWrittenTime_IST': 'N/A',
        'LastModified_UTC': 'N/A', 'LastModified_IST': 'N/A'
    }
    
    if item_type == 0:  # FILE
        if 'File' in item:
            file = item['File']
            details['Name'] = file.get('Name', '')
            details['Path'] = file.get('ServerRelativeUrl', '')
            details['Size'] = file.get('Length', 0)
            
            file_props = get_file_properties(site_url, list_id, item['Id'], access_token)
            if file_props:
                lm = file_props.get('vti_x005f_timelastmodified')
                details['LastModified_UTC'] = lm if lm else 'N/A'
                details['LastModified_IST'] = convert_utc_to_ist(lm)
                
                tag = file_props.get('vti_x005f_complianceTag')
                details['ComplianceTag'] = tag if tag else 'N/A'
                
                ctime = file_props.get('vti_x005f_complianceTagWrittenTime')
                details['ComplianceTagWrittenTime_UTC'] = ctime if ctime else 'N/A'
                details['ComplianceTagWrittenTime_IST'] = convert_utc_to_ist(ctime)
                
    else:  # FOLDER
        if 'Folder' in item:
            folder = item['Folder']
            details['Name'] = folder.get('Name', '')
            details['Path'] = folder.get('ServerRelativeUrl', '')
            
            fv = get_folder_field_values(site_url, list_id, item['Id'], access_token)
            if fv:
                tag = fv.get('OData__x005f_ComplianceTag')
                details['ComplianceTag'] = tag if tag else 'N/A'
                
                ctime = fv.get('OData__x005f_ComplianceTagWrittenTime')
                details['ComplianceTagWrittenTime_UTC'] = ctime if ctime else 'N/A'
                details['ComplianceTagWrittenTime_IST'] = convert_utc_to_ist(ctime)
    
    return details


# ============ Parallel Worker ============
def process_single_item(site_url, list_id, lib_title, item, token):
    """Worker function for threads"""
    details = process_item(site_url, list_id, item, token)
    details['Site'] = site_url
    details['Library'] = lib_title
    return details


# ============ Main bulk routine ============
def prompt_sites():
    print("=" * 70)
    print("BULK SHAREPOINT SCANNER (PARALLEL PROCESSING)")
    print("=" * 70)
    print("Enter SharePoint site URLs, one per line.")
    print("Type 'DONE' on an empty line when finished.\n")

    sites = []
    while True:
        line = input("Site URL (or DONE): ").strip()
        if not line: continue
        if line.upper() == "DONE": break
        if line.lower().endswith(".txt") and os.path.isfile(line):
            with open(line, "r", encoding="utf-8") as f:
                for u in f:
                    u = u.strip()
                    if u and not u.startswith("#"): sites.append(u)
            print(f"  Loaded {len(sites)} URLs from file.")
            continue
        for part in [p.strip() for p in line.split(",") if p.strip()]:
            sites.append(part)

    seen, unique = set(), []
    for s in sites:
        s = s.rstrip("/")
        if s not in seen:
            seen.add(s)
            unique.append(s)
    return unique


def main():
    sites = prompt_sites()
    if not sites:
        print("No sites provided. Exiting.")
        return

    print(f"\nProcessing {len(sites)} site(s) using parallel threads...")

    fieldnames = [
        'Site', 'Library', 'Type', 'ID', 'Name', 'Path', 'Size',
        'Created', 'Modified', 'Author', 'Editor',
        'LastModified_UTC', 'LastModified_IST',
        'ComplianceTag', 'ComplianceTagWrittenTime_UTC', 'ComplianceTagWrittenTime_IST'
    ]

    stats = {"sites_ok": 0, "sites_failed": 0, "libraries": 0, "items": 0, "errors": 0}
    failed_sites = []

    # Set max_workers to 5 or 10. Too high and SharePoint will throttle you to death.
    MAX_WORKERS = 5 

    with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()

        for site_idx, site_url in enumerate(sites, 1):
            print(f"\n{'='*70}")
            print(f"[{site_idx}/{len(sites)}] Site: {site_url}")
            print('='*70)

            try:
                token = get_access_token_for_site(site_url)
            except Exception as e:
                print(f"  Auth failed for site: {e}")
                stats["sites_failed"] += 1
                failed_sites.append((site_url, str(e)))
                continue

            try:
                libraries = get_all_lists(site_url, token)
            except Exception as e:
                print(f"  Could not list libraries: {e}")
                stats["sites_failed"] += 1
                failed_sites.append((site_url, str(e)))
                continue

            if not libraries:
                print("  No document libraries found.")
                stats["sites_ok"] += 1
                continue

            print(f"  Found {len(libraries)} document libraries.")

            for lib_idx, library in enumerate(libraries, 1):
                lib_title = library.get('Title', '(untitled)')
                print(f"\n  [{lib_idx}/{len(libraries)}] Library: {lib_title}")

                try:
                    items = get_list_items(site_url, library['Id'], token)
                except Exception as e:
                    print(f"    Failed to read items: {e}")
                    stats["errors"] += 1
                    continue

                if not items:
                    print("    (empty)")
                    continue

                stats["libraries"] += 1
                print(f"    Processing {len(items)} item(s) in parallel (Threads: {MAX_WORKERS})...")

                # Start the ThreadPool for this library's items
                with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                    # Submit all items to the pool
                    futures = {
                        executor.submit(process_single_item, site_url, library['Id'], lib_title, item, token): item 
                        for item in items
                    }
                    
                    # As they finish, write them to CSV
                    for future in as_completed(futures):
                        item = futures[future]
                        try:
                            details = future.result()
                            # Use Lock so threads don't corrupt the CSV file simultaneously
                            with _csv_lock:
                                writer.writerow(details)
                            stats["items"] += 1
                        except Exception as e:
                            stats["errors"] += 1
                            print(f"    Item {item.get('Id')} error: {e}")

            stats["sites_ok"] += 1

    # ===== Summary =====
    print(f"\n{'='*70}")
    print("SUMMARY")
    print('='*70)
    print(f"Sites processed OK : {stats['sites_ok']}")
    print(f"Sites failed       : {stats['sites_failed']}")
    print(f"Libraries scanned  : {stats['libraries']}")
    print(f"Items written      : {stats['items']}")
    print(f"Item errors        : {stats['errors']}")
    print(f"\nReport: {OUTPUT_CSV}")

    if failed_sites:
        print("\nFailed sites:")
        for s, err in failed_sites:
            print(f"  - {s} :: {err}")


if __name__ == "__main__":
    main()