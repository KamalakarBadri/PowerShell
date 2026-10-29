import requests
import json
from urllib.parse import urljoin
import csv
from datetime import datetime
import re
import uuid
import base64
import time
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend

# ============================================================
# CONFIGURATION
# ============================================================

ONEDRIVE_URLS = [
    "https://geekbyteonline-my.sharepoint.com/personal/user1_geekbyteonline_onmicrosoft_com",
    # Add more OneDrive URLs here
]

OUTPUT_CSV = f"OneDrive_Deletion_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"

TENANT_NAME = "geekbyteonline.onmicrosoft.com"
APP_ID = "73efa35d-6188-42d4-b258-838a977eb149"
SCOPE_SHAREPOINT = "https://geekbyteonline.sharepoint.com/.default"

CERTIFICATE_PATH = "certificate.pem"
PRIVATE_KEY_PATH = "private_key.pem"

# ============================================================
# AUTHENTICATION
# ============================================================

def load_certificate_and_key():
    with open(CERTIFICATE_PATH, "rb") as cert_file:
        certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
    with open(PRIVATE_KEY_PATH, "rb") as key_file:
        private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())
    return certificate, private_key

def get_jwt_token(certificate, private_key):
    now = int(time.time())
    expiration = now + 300
    thumbprint = certificate.fingerprint(hashes.SHA1())
    x5t = base64.urlsafe_b64encode(thumbprint).decode('utf-8').replace('=', '')
    
    jwt_header = {"alg": "RS256", "typ": "JWT", "x5t": x5t}
    jwt_payload = {
        "aud": f"https://login.microsoftonline.com/{TENANT_NAME}/oauth2/v2.0/token",
        "exp": expiration, "iss": APP_ID, "jti": str(uuid.uuid4()),
        "nbf": now, "sub": APP_ID
    }
    
    encoded_header = base64.urlsafe_b64encode(json.dumps(jwt_header, separators=(',', ':')).encode('utf-8')).decode('utf-8').replace('=', '')
    encoded_payload = base64.urlsafe_b64encode(json.dumps(jwt_payload, separators=(',', ':')).encode('utf-8')).decode('utf-8').replace('=', '')
    jwt_unsigned = f"{encoded_header}.{encoded_payload}"
    
    signature = private_key.sign(jwt_unsigned.encode('utf-8'), padding.PKCS1v15(), hashes.SHA256())
    encoded_signature = base64.urlsafe_b64encode(signature).decode('utf-8').replace('=', '')
    return f"{jwt_unsigned}.{encoded_signature}"

def get_access_token(jwt):
    url = f"https://login.microsoftonline.com/{TENANT_NAME}/oauth2/v2.0/token"
    data = {
        "client_id": APP_ID, "client_assertion": jwt,
        "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        "scope": SCOPE_SHAREPOINT, "grant_type": "client_credentials"
    }
    response = requests.post(url, data=data)
    response.raise_for_status()
    return response.json()["access_token"]

def authenticate():
    print("Authenticating...")
    cert, key = load_certificate_and_key()
    jwt = get_jwt_token(cert, key)
    token = get_access_token(jwt)
    print("Authentication successful.")
    return token

def make_request(url, token, method='GET', max_retries=3):
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json;odata=verbose"}
    for attempt in range(max_retries):
        try:
            response = requests.request(method, url, headers=headers, timeout=60)
            if response.status_code == 429:
                time.sleep(5 * (attempt + 1)); continue
            response.raise_for_status()
            # For POST requests, return the response object itself if no JSON body is expected
            if method.upper() == 'POST':
                return response
            return response.json()
        except requests.exceptions.HTTPError:
            if response.status_code == 401 and attempt < max_retries - 1:
                token = authenticate(); headers["Authorization"] = f"Bearer {token}"; continue
            return None
        except Exception:
            if attempt < max_retries - 1: time.sleep(2 ** attempt); continue
            return None
    return None

# ============================================================
# UTILITY
# ============================================================

def safe_int(value):
    if value is None: return 0
    if isinstance(value, (int, float)): return int(value)
    if isinstance(value, str):
        cleaned = re.sub(r'[^\d.]', '', value)
        try: return int(float(cleaned)) if cleaned else 0
        except ValueError: return 0
    return 0

def bytes_to_mb(b):
    b = safe_int(b); return round(b / (1024 * 1024), 2) if b else 0.00

# ============================================================
# DATA RETRIEVAL
# ============================================================

def get_documents_library_id(site_url, token):
    url = f"{site_url}/_api/web/lists?$filter=Title eq 'Documents'"
    data = make_request(url, token)
    if data and 'd' in data and 'results' in data['d'] and len(data['d']['results']) > 0:
        return data['d']['results'][0]['Id']
    return None

def get_root_items(site_url, library_id, token):
    """Uses RootFolder endpoint to get direct children reliably."""
    print("  Fetching root folder contents...")
    url = f"{site_url}/_api/web/lists(guid'{library_id}')/RootFolder?$expand=Files,Folders"
    data = make_request(url, token)
    
    if not data or 'd' not in data:
        print("  Failed to retrieve root folder.")
        return []
    
    root = data['d']
    items = []
    
    for f in root.get('Files', {}).get('results', []):
        items.append({
            'Id': f.get('ListItemAllFields', {}).get('Id') or f.get('UniqueId'),
            'FileLeafRef': f.get('Name'), 'FileRef': f.get('ServerRelativeUrl'),
            'File_x005f_x0020_x005f_Size': f.get('Length'),
            'Created': f.get('TimeCreated'), 'Modified': f.get('TimeLastModified'),
            'FileSystemObjectType': 0,
            'File': {'Name': f.get('Name'), 'ServerRelativeUrl': f.get('ServerRelativeUrl'), 'Length': f.get('Length')}
        })
    
    for fo in root.get('Folders', {}).get('results', []):
        if fo.get('Name') == 'Forms': continue
        items.append({
            'Id': fo.get('ListItemAllFields', {}).get('Id') or fo.get('UniqueId'),
            'FileLeafRef': fo.get('Name'), 'FileRef': fo.get('ServerRelativeUrl'),
            'Created': fo.get('TimeCreated'), 'Modified': fo.get('TimeLastModified'),
            'FileSystemObjectType': 1,
            'Folder': {'Name': fo.get('Name'), 'ServerRelativeUrl': fo.get('ServerRelativeUrl')}
        })
    
    print(f"  Found {len(items)} root items.")
    return items

def get_item_metadata_from_versions(site_url, library_id, item_id, token):
    meta = {'created_by': 'N/A', 'modified_by': 'N/A', 'size': 0, 'version_count': 0}
    try:
        url = f"{site_url}/_api/Web/Lists(guid'{library_id}')/items({item_id})/versions"
        data = make_request(url, token)
        if data and 'd' in data and 'results' in data['d']:
            versions = data['d']['results']
            meta['version_count'] = len(versions)
            if versions:
                latest = versions[0]
                if 'Author' in latest and latest['Author']:
                    meta['created_by'] = latest['Author'].get('LookupValue', 'N/A')
                if 'Editor' in latest and latest['Editor']:
                    meta['modified_by'] = latest['Editor'].get('LookupValue', 'N/A')
                size_val = latest.get('File_x005f_x0020_x005f_Size', 0) or latest.get('SMTotalFileStreamSize', 0)
                meta['size'] = safe_int(size_val)
    except Exception as e:
        print(f"    Metadata error for {item_id}: {e}")
    return meta

def recycle_item(site_url, library_id, item_id, token):
    """Moves item to recycle bin."""
    url = f"{site_url}/_api/web/lists(guid'{library_id}')/items({item_id})/recycle()"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json;odata=verbose"}
    try:
        response = requests.post(url, headers=headers, timeout=60)
        if response.status_code in [200, 201, 204]:
            return "Success", ""
        return "Failed", f"HTTP {response.status_code}"
    except Exception as e:
        return "Failed", str(e)

# ============================================================
# MAIN
# ============================================================

def main():
    print("="*80)
    print("ONEDRIVE ROOT DELETION SCRIPT")
    print("="*80)
    
    # Confirmation
    confirm = input("⚠️ WARNING: This will delete ALL root-level files and folders from the provided OneDrive URLs.\n"
                    "They will be moved to the recycle bin and then to the second-stage recycle bin.\n"
                    "Type 'YES' to continue: ")
    if confirm.strip().upper() != 'YES':
        print("Aborted by user.")
        return
    
    token = authenticate()
    
    with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['OneDrive URL', 'Library', 'Type', 'ID', 'Name', 'Path',
                      'Size (MB)', 'Created', 'Modified', 'Author', 'Editor', 'Version Count',
                      'Deletion Status', 'Error Details']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        
        for onedrive_url in ONEDRIVE_URLS:
            print(f"\nProcessing: {onedrive_url}")
            site_url = onedrive_url.rstrip('/')
            
            lib_id = get_documents_library_id(site_url, token)
            if not lib_id:
                print("  Could not find Documents library.")
                continue
            
            items = get_root_items(site_url, lib_id, token)
            
            for item in items:
                item_id = item.get('Id')
                is_folder = item.get('FileSystemObjectType') == 1
                item_type = 'Folder' if is_folder else 'File'
                name = item.get('FileLeafRef', 'N/A')
                path = item.get('FileRef', 'N/A')
                
                print(f"  {item_type}: {name} (ID: {item_id})...", end="")
                
                meta = {'created_by': 'N/A', 'modified_by': 'N/A', 'size': 0, 'version_count': 0}
                if not is_folder and item_id:
                    meta = get_item_metadata_from_versions(site_url, lib_id, item_id, token)
                
                status, error = recycle_item(site_url, lib_id, item_id, token) if item_id else ("Failed", "No ID")
                print(f" {status}")
                
                writer.writerow({
                    'OneDrive URL': onedrive_url, 'Library': 'Documents', 'Type': item_type,
                    'ID': item_id, 'Name': name, 'Path': path, 'Size (MB)': bytes_to_mb(meta['size']),
                    'Created': item.get('Created', 'N/A'), 'Modified': item.get('Modified', 'N/A'),
                    'Author': meta['created_by'], 'Editor': meta['modified_by'],
                    'Version Count': meta['version_count'], 'Deletion Status': status, 'Error Details': error
                })
            
            # ✅ Move all items to second-stage recycle bin for this site
            print(f"  Moving all items to second-stage recycle bin for {site_url}...")
            move_url = f"{site_url}/_api/web/recycleBin/MoveAllToSecondStage"
            move_response = make_request(move_url, token, method='POST')
            if move_response and move_response.status_code in [200, 201, 204]:
                print("  ✅ Successfully moved to second-stage recycle bin.")
            else:
                print(f"  ❌ Failed to move to second-stage. Status: {move_response.status_code if move_response else 'No response'}")
    
    print(f"\nReport saved: {OUTPUT_CSV}")

if __name__ == "__main__":
    main()