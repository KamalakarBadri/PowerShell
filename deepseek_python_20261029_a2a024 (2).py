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

def get_user_from_lookup(lookup_field):
    """Extract user name from a lookup object (Author/Editor)."""
    if not lookup_field:
        return 'N/A'
    return lookup_field.get('LookupValue', lookup_field.get('Title', 'N/A'))

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
    # Expand Files and Folders, and include Author/Editor for metadata
    url = f"{site_url}/_api/web/lists(guid'{library_id}')/RootFolder?$expand=Files,Folders,Files/Author,Files/Editor"
    data = make_request(url, token)
    
    if not data or 'd' not in data:
        print("  Failed to retrieve root folder.")
        return []
    
    root = data['d']
    items = []
    
    for f in root.get('Files', {}).get('results', []):
        # Extract Author/Editor from the expanded file object
        author = get_user_from_lookup(f.get('Author'))
        editor = get_user_from_lookup(f.get('Editor'))
        size = f.get('Length', 0)
        
        items.append({
            'FileLeafRef': f.get('Name'), 
            'FileRef': f.get('ServerRelativeUrl'),
            'Size': safe_int(size),
            'Created': f.get('TimeCreated'), 
            'Modified': f.get('TimeLastModified'),
            'Author': author,
            'Editor': editor,
            'FileSystemObjectType': 0
        })
    
    for fo in root.get('Folders', {}).get('results', []):
        if fo.get('Name') == 'Forms': continue
        items.append({
            'FileLeafRef': fo.get('Name'), 
            'FileRef': fo.get('ServerRelativeUrl'),
            'Size': 0,
            'Created': fo.get('TimeCreated'), 
            'Modified': fo.get('TimeLastModified'),
            'Author': 'N/A',  # Folders don't have Author/Editor in the same way
            'Editor': 'N/A',
            'FileSystemObjectType': 1
        })
    
    print(f"  Found {len(items)} root items.")
    return items

def recycle_item_by_path(site_url, item, token):
    """Recycles using GetFile/FolderByServerRelativeUrl + /recycle endpoint."""
    is_folder = item.get('FileSystemObjectType') == 1
    server_relative_url = item.get('FileRef')
    
    if not server_relative_url:
        return "Failed", "No ServerRelativeUrl"
    
    encoded_path = server_relative_url.replace(' ', '%20')
    
    if is_folder:
        endpoint = f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{encoded_path}')/recycle"
    else:
        endpoint = f"{site_url}/_api/web/GetFileByServerRelativeUrl('{encoded_path}')/recycle"
    
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json;odata=verbose"}
    try:
        response = requests.post(endpoint, headers=headers, timeout=60)
        if response.status_code in [200, 201, 204]:
            return "Success", ""
        return "Failed", f"HTTP {response.status_code}: {response.text[:100]}"
    except Exception as e:
        return "Failed", str(e)

# ============================================================
# MAIN
# ============================================================

def main():
    print("="*80)
    print("ONEDRIVE ROOT DELETION SCRIPT")
    print("="*80)
    
    confirm = input("⚠️ WARNING: This will delete ALL root-level files and folders from the provided OneDrive URLs.\n"
                    "They will be moved to the recycle bin and then to the second-stage recycle bin.\n"
                    "Type 'YES' to continue: ")
    if confirm.strip().upper() != 'YES':
        print("Aborted by user.")
        return
    
    token = authenticate()
    
    with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['OneDrive URL', 'Library', 'Type', 'Name', 'Path',
                      'Size (MB)', 'Created', 'Modified', 'Author', 'Editor',
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
                is_folder = item.get('FileSystemObjectType') == 1
                item_type = 'Folder' if is_folder else 'File'
                name = item.get('FileLeafRef', 'N/A')
                path = item.get('FileRef', 'N/A')
                
                print(f"  {item_type}: {name}...", end="")
                
                status, error = recycle_item_by_path(site_url, item, token)
                print(f" {status}")
                
                writer.writerow({
                    'OneDrive URL': onedrive_url, 
                    'Library': 'Documents', 
                    'Type': item_type,
                    'Name': name, 
                    'Path': path, 
                    'Size (MB)': bytes_to_mb(item.get('Size', 0)),
                    'Created': item.get('Created', 'N/A'), 
                    'Modified': item.get('Modified', 'N/A'),
                    'Author': item.get('Author', 'N/A'), 
                    'Editor': item.get('Editor', 'N/A'),
                    'Deletion Status': status, 
                    'Error Details': error
                })
            
            # Move all items to second-stage recycle bin for this site
            print(f"  Moving all items to second-stage recycle bin for {site_url}...")
            move_url = f"{site_url}/_api/web/recycleBin/MoveAllToSecondStage"
            move_response = make_request(move_url, token, method='POST')
            if move_response and move_response.status_code in [200, 201, 204]:
                print("  ✅ Successfully moved to second-stage recycle bin.")
            else:
                status_code = move_response.status_code if move_response else 'No response'
                print(f"  ❌ Failed to move to second-stage. Status: {status_code}")
    
    print(f"\nReport saved: {OUTPUT_CSV}")

if __name__ == "__main__":
    main()