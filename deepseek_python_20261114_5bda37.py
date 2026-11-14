import requests
import json
from urllib.parse import urljoin, urlparse
import csv
from datetime import datetime
import os
import threading
import time
import uuid
import base64
from concurrent.futures import ThreadPoolExecutor, as_completed
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend

# ============ Configuration ============
OUTPUT_CSV = f"SharePoint_Recycle_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"

TENANT_NAME = "test.onmicrosoft.com"
APP_ID = "73efa35d-6188-42d4-b258-838a977eb149"

CERTIFICATE_PATH = "certificate.pem"
PRIVATE_KEY_PATH = "private_key.pem"

# Token cache: { site_host: {token, expires_at} }
_token_cache = {}
_token_lock = threading.Lock()
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

# ============ SharePoint Request Helper ============
def make_sharepoint_request(url, access_token, method='GET', retries=3):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json",
    }
    
    delay = 2
    for attempt in range(retries):
        try:
            r = requests.request(method, url, headers=headers, timeout=60)
            if r.status_code == 429:
                wait = int(r.headers.get("Retry-After", delay))
                time.sleep(wait)
                continue
            if 500 <= r.status_code < 600:
                time.sleep(delay)
                delay *= 2
                continue
            
            if r.status_code == 404:
                return {"success": True, "msg": "Already deleted/missing"}
                
            r.raise_for_status()
            return {"success": True, "msg": "Recycled successfully"}
            
        except requests.exceptions.RequestException as e:
            time.sleep(delay)
            delay *= 2
            error_msg = str(e)
            if hasattr(e, 'response') and e.response is not None:
                error_msg = e.response.text
            return {"success": False, "msg": f"Error: {error_msg}"}
            
    return {"success": False, "msg": "Failed after retries"}

# ============ SharePoint Data Retrieval ============
def get_all_lists(site_url, access_token):
    lists_url = urljoin(site_url + "/", "_api/web/lists")
    response = requests.get(lists_url, headers={"Authorization": f"Bearer {access_token}", "Accept": "application/json;odata=verbose"})
    if response.status_code == 200:
        js = response.json()
        if 'd' in js and 'results' in js['d']:
            return [lst for lst in js['d']['results'] if lst['BaseTemplate'] == 101]
    return []

def get_list_items(site_url, list_id, access_token):
    items_url = f"{site_url}/_api/web/lists(guid'{list_id}')/items?$expand=File,Folder&$top=5000"
    all_items, next_url, page = [], items_url, 0
    
    while next_url:
        response = requests.get(next_url, headers={"Authorization": f"Bearer {access_token}", "Accept": "application/json;odata=verbose"})
        if response.status_code != 200:
            break
            
        js = response.json()
        if 'd' in js and 'results' in js['d']:
            batch = js['d']['results']
            all_items.extend(batch)
            print(f"      page {page}: fetched {len(batch)} (total {len(all_items)})")
        page += 1
        next_url = js.get('d', {}).get('__next')
        
    return all_items

# ============ Deletion Logic ============
def recycle_single_item(site_url, list_id, item, access_token):
    """Worker function for threads. Uses /recycle() endpoint."""
    item_type = 'File' if item['FileSystemObjectType'] == 0 else 'Folder'
    
    if item_type == 'File':
        name = item.get('File', {}).get('Name', '')
        path = item.get('File', {}).get('ServerRelativeUrl', '')
    else:
        name = item.get('Folder', {}).get('Name', '')
        path = item.get('Folder', {}).get('ServerRelativeUrl', '')

    # RECYCLE endpoint instead of delete
    recycle_url = f"{site_url}/_api/web/lists(guid'{list_id}')/items({item['Id']})/recycle()"
    
    # POST is required for the recycle endpoint
    result = make_sharepoint_request(recycle_url, access_token, method='POST')
    
    # Note: 'ID' is intentionally omitted here since it's not in the CSV fieldnames
    return {
        'Type': item_type,
        'Name': name,
        'Path': path,
        'Status': result['msg']
    }

# ============ Main Routine ============
def prompt_sites():
    print("=" * 70)
    print("BULK SHAREPOINT RECYCLE TOOL")
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

    return list(dict.fromkeys([s.rstrip("/") for s in sites]))

def main():
    sites = prompt_sites()
    if not sites:
        return

    # REMOVED 'ID' from here to match the dictionary returned by the thread
    fieldnames = ['Site', 'Library', 'Type', 'Name', 'Path', 'Status']
    stats = {"recycled": 0, "errors": 0}
    
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
                continue

            libraries = get_all_lists(site_url, token)
            if not libraries:
                print("  No document libraries found.")
                continue

            for lib_idx, library in enumerate(libraries, 1):
                lib_title = library.get('Title', '(untitled)')
                list_id = library['Id']
                print(f"\n  [{lib_idx}/{len(libraries)}] Library: {lib_title}")

                items = get_list_items(site_url, list_id, token)
                if not items:
                    print("    (empty)")
                    continue

                print(f"    Sorting items for safe recycling...")
                
                # Sort Files first (0), Folders last (1). Folders sorted by deepest path first.
                def sort_key(x):
                    is_folder = x['FileSystemObjectType']
                    path = x.get('File', {}).get('ServerRelativeUrl') or x.get('Folder', {}).get('ServerRelativeUrl') or ''
                    return (is_folder, -len(path))

                sorted_items = sorted(items, key=sort_key)
                print(f"    Recycling {len(sorted_items)} item(s) in parallel...")

                MAX_WORKERS = 5
                
                with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
                    futures = {
                        executor.submit(recycle_single_item, site_url, list_id, item, token): item 
                        for item in sorted_items
                    }
                    
                    for future in as_completed(futures):
                        try:
                            details = future.result()
                            details['Site'] = site_url
                            details['Library'] = lib_title
                            
                            if "successfully" in details['Status'] or "Already" in details['Status']:
                                stats["recycled"] += 1
                            else:
                                stats["errors"] += 1
                                
                            with _csv_lock:
                                writer.writerow(details)
                                
                        except Exception as e:
                            stats["errors"] += 1
                            print(f"    Thread Error: {e}")

    # ===== Summary =====
    print(f"\n{'='*70}")
    print("RECYCLE SUMMARY")
    print('='*70)
    print(f"Items successfully recycled : {stats['recycled']}")
    print(f"Items failed/errors         : {stats['errors']}")
    print(f"\nReport saved to: {OUTPUT_CSV}")

if __name__ == "__main__":
    main()