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

# List of OneDrive URLs you want to process
ONEDRIVE_URLS = [
    "https://geekbyteonline-my.sharepoint.com/personal/user1_geekbyteonline_onmicrosoft_com",
    # Add more OneDrive URLs here
]

OUTPUT_CSV = f"OneDrive_Deletion_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"

# Authentication Configuration
TENANT_NAME = "geekbyteonline.onmicrosoft.com"
APP_ID = "73efa35d-6188-42d4-b258-838a977eb149"
SCOPE_SHAREPOINT = "https://geekbyteonline.sharepoint.com/.default"

CERTIFICATE_PATH = "certificate.pem"
PRIVATE_KEY_PATH = "private_key.pem"

# ============================================================
# AUTHENTICATION FUNCTIONS
# ============================================================

def load_certificate_and_key():
    try:
        with open(CERTIFICATE_PATH, "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
        with open(PRIVATE_KEY_PATH, "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())
        return certificate, private_key
    except Exception as e:
        print(f"Error loading certificate: {str(e)}")
        raise

def get_jwt_token(certificate, private_key):
    try:
        now = int(time.time())
        expiration = now + 300
        thumbprint = certificate.fingerprint(hashes.SHA1())
        x5t = base64.urlsafe_b64encode(thumbprint).decode('utf-8').replace('=', '')
        
        jwt_header = {"alg": "RS256", "typ": "JWT", "x5t": x5t}
        jwt_payload = {
            "aud": f"https://login.microsoftonline.com/{TENANT_NAME}/oauth2/v2.0/token",
            "exp": expiration,
            "iss": APP_ID,
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": APP_ID
        }
        
        encoded_header = base64.urlsafe_b64encode(json.dumps(jwt_header, separators=(',', ':')).encode('utf-8')).decode('utf-8').replace('=', '')
        encoded_payload = base64.urlsafe_b64encode(json.dumps(jwt_payload, separators=(',', ':')).encode('utf-8')).decode('utf-8').replace('=', '')
        jwt_unsigned = f"{encoded_header}.{encoded_payload}"
        
        signature = private_key.sign(jwt_unsigned.encode('utf-8'), padding.PKCS1v15(), hashes.SHA256())
        encoded_signature = base64.urlsafe_b64encode(signature).decode('utf-8').replace('=', '')
        
        return f"{jwt_unsigned}.{encoded_signature}"
    except Exception as e:
        print(f"Error generating JWT: {str(e)}")
        raise

def get_access_token(jwt):
    url = f"https://login.microsoftonline.com/{TENANT_NAME}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    data = {
        "client_id": APP_ID,
        "client_assertion": jwt,
        "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        "scope": SCOPE_SHAREPOINT,
        "grant_type": "client_credentials"
    }
    try:
        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()
        return response.json()["access_token"]
    except Exception as e:
        print(f"Error getting access token: {str(e)}")
        raise

def authenticate():
    print("Authenticating...")
    cert, key = load_certificate_and_key()
    jwt = get_jwt_token(cert, key)
    token = get_access_token(jwt)
    print("Authentication successful.")
    return token

def make_request(url, token, method='GET', max_retries=3):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json"
    }
    for attempt in range(max_retries):
        try:
            response = requests.request(method, url, headers=headers, timeout=60)
            if response.status_code == 429:
                wait = 5 * (attempt + 1)
                print(f"  Rate limited. Waiting {wait}s...")
                time.sleep(wait)
                continue
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as e:
            if response.status_code == 401 and attempt < max_retries - 1:
                token = authenticate()
                headers["Authorization"] = f"Bearer {token}"
                continue
            print(f"Request failed: {str(e)}")
            return None
        except Exception as e:
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)
                continue
            return None
    return None

# ============================================================
# UTILITY FUNCTIONS
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
    b = safe_int(b)
    return round(b / (1024 * 1024), 2) if b else 0.00

def get_site_url_from_onedrive(onedrive_url):
    """Converts a OneDrive personal URL to a SharePoint site URL."""
    # Example: https://tenant-my.sharepoint.com/personal/user_tenant_onmicrosoft_com
    # -> https://tenant-my.sharepoint.com/personal/user_tenant_onmicrosoft_com
    # Actually, for OneDrive for Business, the site URL is the personal URL itself.
    return onedrive_url.rstrip('/')

# ============================================================
# DATA RETRIEVAL
# ============================================================

def get_documents_library_id(site_url, token):
    """Finds the 'Documents' library ID in the OneDrive site."""
    url = f"{site_url}/_api/web/lists?$filter=Title eq 'Documents'"
    data = make_request(url, token)
    if data and 'd' in data and 'results' in data['d'] and len(data['d']['results']) > 0:
        return data['d']['results'][0]['Id']
    print(f"  Could not find 'Documents' library at {site_url}")
    return None

def get_root_items(site_url, library_id, token):
    """Gets all items directly under the root of the library."""
    # Filter by parent path to ensure we only get root items
    # The root path for OneDrive Documents is typically '/personal/username/Documents'
    # We can use the FileDirRef field to match the root.
    # First, let's get the root folder info to know the exact path.
    root_url = f"{site_url}/_api/web/lists(guid'{library_id}')/rootFolder"
    root_data = make_request(root_url, token)
    if not root_data or 'd' not in root_data:
        print("  Could not retrieve root folder info.")
        return []
    
    root_path = root_data['d']['ServerRelativeUrl']
    print(f"  Root path: {root_path}")
    
    # Now get items where FileDirRef equals root_path
    # We need to expand File to get size, and Author/Editor for names (though we will fetch from versions for accuracy)
    filter_query = f"$filter=FileDirRef eq '{root_path}'"
    select_query = "$select=Id,FileLeafRef,FileRef,File_x005f_x0020_x005f_Size,Created,Modified,FileSystemObjectType"
    expand_query = "$expand=File"
    
    items_url = f"{site_url}/_api/web/lists(guid'{library_id}')/items?{select_query}&{expand_query}&{filter_query}"
    
    all_items = []
    next_url = items_url
    while next_url:
        data = make_request(next_url, token)
        if not data: break
        if 'd' in data and 'results' in data['d']:
            all_items.extend(data['d']['results'])
        next_url = data.get('d', {}).get('__next')
    
    print(f"  Found {len(all_items)} root items.")
    return all_items

def get_item_metadata_from_versions(site_url, library_id, item_id, token):
    """Fetches latest version details (Author, Editor, Size) and version count."""
    meta = {
        'created_by': 'N/A',
        'modified_by': 'N/A',
        'size': 0,
        'version_count': 0
    }
    try:
        # Version count
        count_url = f"{site_url}/_api/Web/Lists(guid'{library_id}')/items({item_id})/versions"
        versions_data = make_request(count_url, token)
        if versions_data and 'd' in versions_data and 'results' in versions_data['d']:
            meta['version_count'] = len(versions_data['d']['results'])
            
            # Latest version details (for Author/Editor/Size)
            if meta['version_count'] > 0:
                latest = versions_data['d']['results'][0]
                if 'Author' in latest and latest['Author']:
                    meta['created_by'] = latest['Author'].get('LookupValue', 'N/A')
                if 'Editor' in latest and latest['Editor']:
                    meta['modified_by'] = latest['Editor'].get('LookupValue', 'N/A')
                
                size_val = latest.get('File_x005f_x0020_x005f_Size', 0)
                if not size_val:
                    size_val = latest.get('SMTotalFileStreamSize', 0)
                meta['size'] = safe_int(size_val)
    except Exception as e:
        print(f"    Error fetching metadata for item {item_id}: {e}")
    
    return meta

def recycle_item(site_url, library_id, item_id, token):
    """Moves an item to the recycle bin using the recycle endpoint."""
    url = f"{site_url}/_api/web/lists(guid'{library_id}')/items({item_id})/recycle()"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json"
    }
    try:
        # The recycle endpoint is a POST request
        response = requests.post(url, headers=headers, timeout=60)
        if response.status_code in [200, 201, 204]:
            return "Success", ""
        else:
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
    
    token = authenticate()
    
    with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = [
            'OneDrive URL', 'Library', 'Type', 'ID', 'Name', 'Path',
            'Size (MB)', 'Created', 'Modified',
            'Author (Created By)', 'Editor (Modified By)', 'Version Count',
            'Deletion Status', 'Error Details'
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        
        for onedrive_url in ONEDRIVE_URLS:
            print(f"\nProcessing: {onedrive_url}")
            site_url = get_site_url_from_onedrive(onedrive_url)
            
            lib_id = get_documents_library_id(site_url, token)
            if not lib_id:
                continue
            
            items = get_root_items(site_url, lib_id, token)
            
            for item in items:
                item_id = item.get('Id')
                is_folder = item.get('FileSystemObjectType') == 1
                item_type = 'Folder' if is_folder else 'File'
                
                # Basic name/path
                name = item.get('FileLeafRef', 'N/A')
                path = item.get('FileRef', 'N/A')
                
                print(f"  Processing {item_type}: {name} (ID: {item_id})...", end="")
                
                # Get detailed metadata from versions (for files only, folders may not have versions)
                meta = {'created_by': 'N/A', 'modified_by': 'N/A', 'size': 0, 'version_count': 0}
                if not is_folder:
                    meta = get_item_metadata_from_versions(site_url, lib_id, item_id, token)
                
                # Attempt Recycle
                status, error = recycle_item(site_url, lib_id, item_id, token)
                print(f" {status}")
                
                # Write row
                writer.writerow({
                    'OneDrive URL': onedrive_url,
                    'Library': 'Documents',
                    'Type': item_type,
                    'ID': item_id,
                    'Name': name,
                    'Path': path,
                    'Size (MB)': bytes_to_mb(meta['size']),
                    'Created': item.get('Created', 'N/A'),
                    'Modified': item.get('Modified', 'N/A'),
                    'Author (Created By)': meta['created_by'],
                    'Editor (Modified By)': meta['modified_by'],
                    'Version Count': meta['version_count'],
                    'Deletion Status': status,
                    'Error Details': error
                })
    
    print(f"\nReport saved to: {OUTPUT_CSV}")

if __name__ == "__main__":
    main()