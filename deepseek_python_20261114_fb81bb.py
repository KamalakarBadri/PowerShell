import requests
import json
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

OUTPUT_CSV = f"OneDrive_ItemWise_Deletion_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"

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
    if not lookup_field:
        return 'N/A'
    return lookup_field.get('LookupValue', lookup_field.get('Title', 'N/A'))

# ============================================================
# GET DOCUMENTS LIBRARY
# ============================================================

def get_documents_library_id(site_url, token):
    """Get the Documents library ID by title."""
    print("  Finding Documents library...")
    url = f"{site_url}/_api/web/lists?$filter=Title eq 'Documents'"
    data = make_request(url, token)
    if data and 'd' in data and 'results' in data['d'] and len(data['d']['results']) > 0:
        lib = data['d']['results'][0]
        print(f"  Found: {lib['Title']} (ID: {lib['Id']})")
        return lib['Id']
    print("  Documents library not found.")
    return None

# ============================================================
# GET ALL ITEMS VIA /items ENDPOINT
# ============================================================

def get_all_items(site_url, library_id, token):
    """Get ALL items from the library via /items endpoint with pagination."""
    print("  Fetching all items via /items endpoint...")
    
    items_url = (
        f"{site_url}/_api/web/lists(guid'{library_id}')/items"
        f"?$expand=File,Folder,Author,Editor"
    )
    
    all_items = []
    next_url = items_url
    page = 0
    
    while next_url:
        page += 1
        print(f"    Fetching page {page}...", end="")
        data = make_request(next_url, token)
        if not data:
            print(" failed")
            break
        if 'd' in data and 'results' in data['d']:
            batch = data['d']['results']
            all_items.extend(batch)
            print(f" got {len(batch)} items (total: {len(all_items)})")
        next_url = data.get('d', {}).get('__next')
    
    print(f"  Total items fetched: {len(all_items)}")
    return all_items

# ============================================================
# SPLIT ITEMS INTO FILES AND FOLDERS
# ============================================================

def split_items(items):
    """Split items into files and folders with metadata extracted."""
    files = []
    folders = []
    
    for item in items:
        is_folder = item.get('FileSystemObjectType') == 1
        
        path = item.get('FileRef', '')
        if not path:
            continue
        
        # Get name
        if is_folder:
            folder_obj = item.get('Folder') or {}
            name = folder_obj.get('Name') or item.get('FileLeafRef', '')
        else:
            file_obj = item.get('File') or {}
            name = file_obj.get('Name') or item.get('FileLeafRef', '')
        
        # Get size (files only)
        size = 0
        if not is_folder:
            file_obj = item.get('File') or {}
            size = file_obj.get('Length', 0)
            if not size:
                size = item.get('File_x005f_x0020_x005f_Size', 0)
        
        author = get_user_from_lookup(item.get('Author'))
        editor = get_user_from_lookup(item.get('Editor'))
        
        item_data = {
            'Id': item.get('Id'),
            'Name': name,
            'Path': path,
            'Size': safe_int(size),
            'Created': item.get('Created', 'N/A'),
            'Modified': item.get('Modified', 'N/A'),
            'Author': author,
            'Editor': editor,
            'IsFolder': is_folder
        }
        
        if is_folder:
            folders.append(item_data)
        else:
            files.append(item_data)
    
    return files, folders

# ============================================================
# RECYCLE ITEM USING items(id)/recycle
# ============================================================

def recycle_item_by_id(site_url, library_id, item, token):
    """Recycle a single file or folder using items(id)/recycle endpoint."""
    item_id = item.get('Id')
    
    if not item_id:
        return "Failed", "No Item ID"
    
    endpoint = f"{site_url}/_api/web/lists(guid'{library_id}')/items({item_id})/recycle()"
    
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
    print("ONEDRIVE ITEM-WISE DELETION SCRIPT")
    print("(Get Documents library -> Get items -> Delete files -> Delete folders)")
    print("="*80)
    
    confirm = input(
        "⚠️ WARNING: This will delete ALL items (files and folders) from the Documents library\n"
        "of the provided OneDrive URLs. Items will be moved to the recycle bin and then\n"
        "to the second-stage recycle bin.\n"
        "Type 'YES' to continue: "
    )
    if confirm.strip().upper() != 'YES':
        print("Aborted by user.")
        return
    
    token = authenticate()
    
    with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = [
            'OneDrive URL', 'Type', 'Item ID', 'Name', 'Path', 'Size (MB)',
            'Created', 'Modified', 'Author', 'Editor',
            'Deletion Status', 'Error Details'
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        
        for onedrive_url in ONEDRIVE_URLS:
            print(f"\n{'='*60}")
            print(f"Processing: {onedrive_url}")
            print(f"{'='*60}")
            site_url = onedrive_url.rstrip('/')
            
            # Step 1: Get Documents library
            lib_id = get_documents_library_id(site_url, token)
            if not lib_id:
                print("  Skipping this URL.")
                continue
            
            # Step 2: Get all items
            all_items = get_all_items(site_url, lib_id, token)
            if not all_items:
                print("  No items found. Skipping.")
                continue
            
            # Step 3: Split into files and folders
            files, folders = split_items(all_items)
            print(f"\n  Summary:")
            print(f"    Files:   {len(files)}")
            print(f"    Folders: {len(folders)}")
            
            # Step 4: Delete FILES first using items(id)/recycle
            if files:
                print(f"\n  --- Deleting {len(files)} FILES (items(id)/recycle) ---")
                for idx, file in enumerate(files, 1):
                    print(f"    [{idx}/{len(files)}] File: {file['Name']} (ID: {file['Id']})...", end="")
                    status, error = recycle_item_by_id(site_url, lib_id, file, token)
                    print(f" {status}")
                    
                    writer.writerow({
                        'OneDrive URL': onedrive_url,
                        'Type': 'File',
                        'Item ID': file['Id'],
                        'Name': file['Name'],
                        'Path': file['Path'],
                        'Size (MB)': bytes_to_mb(file.get('Size', 0)),
                        'Created': file.get('Created', 'N/A'),
                        'Modified': file.get('Modified', 'N/A'),
                        'Author': file.get('Author', 'N/A'),
                        'Editor': file.get('Editor', 'N/A'),
                        'Deletion Status': status,
                        'Error Details': error
                    })
                    csvfile.flush()
            
            # Step 5: Delete FOLDERS after files
            if folders:
                print(f"\n  --- Deleting {len(folders)} FOLDERS (items(id)/recycle) ---")
                for idx, folder in enumerate(folders, 1):
                    print(f"    [{idx}/{len(folders)}] Folder: {folder['Name']} (ID: {folder['Id']})...", end="")
                    status, error = recycle_item_by_id(site_url, lib_id, folder, token)
                    print(f" {status}")
                    
                    writer.writerow({
                        'OneDrive URL': onedrive_url,
                        'Type': 'Folder',
                        'Item ID': folder['Id'],
                        'Name': folder['Name'],
                        'Path': folder['Path'],
                        'Size (MB)': 'N/A',
                        'Created': folder.get('Created', 'N/A'),
                        'Modified': folder.get('Modified', 'N/A'),
                        'Author': folder.get('Author', 'N/A'),
                        'Editor': folder.get('Editor', 'N/A'),
                        'Deletion Status': status,
                        'Error Details': error
                    })
                    csvfile.flush()
            
            # Step 6: Move all to second-stage recycle bin
            print(f"\n  Moving all items to second-stage recycle bin...")
            move_url = f"{site_url}/_api/web/recycleBin/MoveAllToSecondStage"
            move_response = make_request(move_url, token, method='POST')
            if move_response and move_response.status_code in [200, 201, 204]:
                print("  ✅ Successfully moved to second-stage recycle bin.")
            else:
                status_code = move_response.status_code if move_response else 'No response'
                print(f"  ❌ Failed to move to second-stage. Status: {status_code}")
    
    print(f"\n{'='*80}")
    print(f"✅ DONE. Report saved: {OUTPUT_CSV}")
    print(f"{'='*80}")

if __name__ == "__main__":
    main()