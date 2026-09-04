import requests
import json
import csv
import uuid
import base64
import time
import os
import sys
import re
from datetime import datetime
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend

# ============================================================
# CONFIGURATION - UPDATE THESE VALUES
# ============================================================

CONFIG = {
    # SharePoint Authentication
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "scope": "https://geekbyteonline.sharepoint.com/.default",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    
    # Site and Library Details
    "site_url": "https://geekbyteonline.sharepoint.com/sites/Team_",
    "library_name": "Documents",
    
    # Report Settings
    "include_folders": True,
    "include_files": True,
    "include_versions": True,
    "include_created_by": True,
    "include_modified_by": True,
    
    # Output Settings
    "output_dir": "library_reports",
    "report_filename": "library_inventory_report.csv"
}

# ============================================================
# GLOBAL VARIABLES
# ============================================================

TOKEN_CACHE = {"token": None, "expires": 0}
csv_writer = None
csv_file = None
total_files = 0
total_folders = 0

# ============================================================
# AUTHENTICATION FUNCTIONS
# ============================================================

def load_certificate_and_key():
    try:
        if not os.path.exists(CONFIG['certificate_path']) or not os.path.exists(CONFIG['private_key_path']):
            raise Exception(f"Certificate files not found.")
        
        with open(CONFIG['certificate_path'], "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
        
        with open(CONFIG['private_key_path'], "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())
        
        return certificate, private_key
    except Exception as e:
        print(f"❌ Error loading certificate: {str(e)}")
        raise

def get_jwt_token(certificate, private_key):
    try:
        now = int(time.time())
        expiration = now + 300
        
        thumbprint = certificate.fingerprint(hashes.SHA1())
        x5t = base64.urlsafe_b64encode(thumbprint).decode('utf-8').replace('=', '')
        
        jwt_header = {"alg": "RS256", "typ": "JWT", "x5t": x5t}
        
        jwt_payload = {
            "aud": f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            "exp": expiration,
            "iss": CONFIG['app_id'],
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": CONFIG['app_id']
        }
        
        encoded_header = base64.urlsafe_b64encode(
            json.dumps(jwt_header, separators=(',', ':')).encode('utf-8')
        ).decode('utf-8').replace('=', '')
        
        encoded_payload = base64.urlsafe_b64encode(
            json.dumps(jwt_payload, separators=(',', ':')).encode('utf-8')
        ).decode('utf-8').replace('=', '')
        
        jwt_unsigned = f"{encoded_header}.{encoded_payload}"
        
        signature = private_key.sign(
            jwt_unsigned.encode('utf-8'),
            padding.PKCS1v15(),
            hashes.SHA256()
        )
        encoded_signature = base64.urlsafe_b64encode(signature).decode('utf-8').replace('=', '')
        
        return f"{jwt_unsigned}.{encoded_signature}"
    except Exception as e:
        print(f"❌ Error generating JWT: {str(e)}")
        raise

def get_access_token(jwt):
    print("  🔑 Getting access token...")
    url = f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token"
    
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    
    data = {
        "client_id": CONFIG['app_id'],
        "client_assertion": jwt,
        "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        "scope": CONFIG['scope'],
        "grant_type": "client_credentials"
    }
    
    try:
        response = requests.post(url, headers=headers, data=data, timeout=120)
        response.raise_for_status()
        result = response.json()
        print("  ✅ Token obtained")
        return result["access_token"]
    except Exception as e:
        print(f"  ❌ Error: {str(e)}")
        raise

def get_cached_token(force_refresh=False):
    cache = TOKEN_CACHE
    
    if not force_refresh and cache["token"] and cache["expires"] > time.time() + 300:
        return cache["token"]
    
    try:
        certificate, private_key = load_certificate_and_key()
        jwt = get_jwt_token(certificate, private_key)
        token = get_access_token(jwt)
        
        if token:
            cache["token"] = token
            cache["expires"] = time.time() + 3600
            return token
        return None
    except Exception as e:
        print(f"  ❌ Authentication failed: {str(e)}")
        return None

def get_current_token():
    return get_cached_token()

def make_sharepoint_request(url, max_retries=5):
    """Make request with retry logic"""
    
    for attempt in range(max_retries + 1):
        try:
            token = get_current_token()
            if not token:
                return None
            
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json"
            }
            
            response = requests.get(url, headers=headers, timeout=120)
            
            if response.status_code == 429:
                wait_time = 3 * (attempt + 1)
                print(f"\n    ⚠️ 429: Waiting {wait_time}s...")
                time.sleep(wait_time)
                continue
            
            if response.status_code == 401 and attempt < max_retries:
                print(f"\n    ⚠️ Token expired, refreshing...")
                TOKEN_CACHE["token"] = None
                TOKEN_CACHE["expires"] = 0
                continue
            
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 429:
                continue
            if attempt < max_retries:
                time.sleep(2 ** attempt)
                continue
            print(f"  ❌ Request failed: {str(e)}")
            return None
        except Exception as e:
            if attempt < max_retries:
                time.sleep(2 ** attempt)
                continue
            print(f"  ❌ Error: {str(e)}")
            return None
    
    return None

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def safe_int_conversion(value):
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        cleaned = re.sub(r'[^\d.]', '', value)
        try:
            return int(float(cleaned)) if cleaned else 0
        except ValueError:
            return 0
    return 0

def bytes_to_mb(bytes_value):
    bytes_value = safe_int_conversion(bytes_value)
    if bytes_value == 0:
        return 0.00
    return round(bytes_value / (1024 * 1024), 2)

def bytes_to_gb(bytes_value):
    bytes_value = safe_int_conversion(bytes_value)
    if bytes_value == 0:
        return 0.00
    return round(bytes_value / (1024 * 1024 * 1024), 2)

def format_datetime(datetime_str):
    if not datetime_str or datetime_str == "N/A" or datetime_str == "0":
        return "N/A"
    try:
        if 'T' in datetime_str:
            if '.' in datetime_str:
                dt = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S.%fZ")
            else:
                dt = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%SZ")
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        return datetime_str
    except:
        return datetime_str

def get_user_name(user_data):
    """Extract user name from user object"""
    if not user_data:
        return ''
    return user_data.get('Title', user_data.get('LookupValue', ''))

def get_user_email(user_data):
    """Extract user email from user object"""
    if not user_data:
        return ''
    return user_data.get('Email', '')

# ============================================================
# SHAREPOINT DATA RETRIEVAL
# ============================================================

def get_library_id(site_url, library_name):
    """Get library ID by name"""
    print(f"\n🔍 Finding library: {library_name}")
    
    url = f"{site_url}/_api/web/lists?$filter=Title eq '{library_name}'"
    response = make_sharepoint_request(url)
    
    if not response or 'd' not in response or 'results' not in response['d']:
        print(f"❌ Library '{library_name}' not found")
        return None
    
    if len(response['d']['results']) == 0:
        print(f"❌ Library '{library_name}' not found")
        return None
    
    library = response['d']['results'][0]
    print(f"✅ Found library: {library['Title']} (ID: {library['Id']})")
    return library['Id']

def get_all_items_from_library(site_url, library_id, library_name):
    """Get all items from library with pagination"""
    print(f"\n📁 Fetching items from library: {library_name}")
    
    # Include CreatedBy and ModifiedBy fields
    items_url = f"{site_url}/_api/web/lists(guid'{library_id}')/items?$expand=CreatedBy,ModifiedBy&$select=Id,Title,FileLeafRef,FileRef,File_x005f_x0020_x005f_Size,Created,Modified,FileSystemObjectType,CreatedBy/Title,CreatedBy/Email,ModifiedBy/Title,ModifiedBy/Email&$top=5000"
    
    all_items = []
    next_url = items_url
    page_count = 0
    
    while next_url:
        page_count += 1
        print(f"  Fetching page {page_count}...", end="")
        response = make_sharepoint_request(next_url)
        
        if not response or 'd' not in response:
            print(" ✗ Failed")
            break
        
        if 'results' in response['d']:
            items_in_page = len(response['d']['results'])
            all_items.extend(response['d']['results'])
            print(f" ✓ Got {items_in_page} items")
        
        next_url = None
        if '__next' in response.get('d', {}):
            next_url = response['d']['__next']
    
    print(f"  Total items fetched: {len(all_items)}")
    return all_items

def get_file_versions_count(site_url, list_id, item_id):
    """Get version count for a specific item"""
    try:
        versions_url = f"{site_url}/_api/Web/Lists(guid'{list_id}')/items({item_id})/versions?$select=VersionId&$top=1"
        
        response = make_sharepoint_request(versions_url)
        
        if not response or 'd' not in response:
            return 0
        
        if 'results' not in response['d']:
            return 0
        
        # Get total count from metadata
        if '__count' in response['d']:
            return safe_int_conversion(response['d']['__count'])
        
        return len(response['d']['results'])
        
    except Exception as e:
        return 0

def get_item_details(site_url, list_id, item):
    """Extract all details from item including Created By and Modified By"""
    
    item_id = item.get('Id', 0)
    fsob_type = item.get('FileSystemObjectType', 0)
    is_folder = fsob_type == 1
    
    # Get file/folder name
    name = item.get('FileLeafRef', '')
    if not name:
        name = item.get('Title', f"Item_{item_id}")
    
    # Get path
    path = item.get('FileRef', '')
    
    # Get file size (only for files)
    file_size = 0
    if not is_folder:
        file_size = item.get('File_x005f_x0020_x005f_Size', 0)
        if not file_size or file_size == 0:
            file_size = item.get('FileSize', 0)
    
    # Get Created By
    created_by_data = item.get('CreatedBy', {})
    created_by = get_user_name(created_by_data)
    created_by_email = get_user_email(created_by_data)
    
    # Get Modified By
    modified_by_data = item.get('ModifiedBy', {})
    modified_by = get_user_name(modified_by_data)
    modified_by_email = get_user_email(modified_by_data)
    
    # Get dates
    created = item.get('Created', 'N/A')
    modified = item.get('Modified', 'N/A')
    
    # Get version count (only for files)
    version_count = 0
    if not is_folder and CONFIG.get('include_versions', True):
        version_count = get_file_versions_count(site_url, list_id, item_id)
    
    # Build URL
    url = f"{site_url}{path}" if path else ""
    
    return {
        'id': item_id,
        'name': name,
        'path': path,
        'url': url,
        'is_folder': is_folder,
        'size': file_size,
        'size_mb': bytes_to_mb(file_size),
        'size_gb': bytes_to_gb(file_size),
        'created': created,
        'created_formatted': format_datetime(created),
        'modified': modified,
        'modified_formatted': format_datetime(modified),
        'created_by': created_by,
        'created_by_email': created_by_email,
        'modified_by': modified_by,
        'modified_by_email': modified_by_email,
        'version_count': version_count
    }

# ============================================================
# CSV REPORT FUNCTIONS
# ============================================================

def initialize_csv():
    """Initialize CSV file with headers"""
    global csv_writer, csv_file
    
    os.makedirs(CONFIG['output_dir'], exist_ok=True)
    
    filename = os.path.join(CONFIG['output_dir'], CONFIG['report_filename'])
    
    # Add timestamp if file exists
    if os.path.exists(filename):
        timestamp = datetime.now().strftime('_%Y-%m-%d_%H-%M-%S')
        base, ext = os.path.splitext(filename)
        filename = f"{base}{timestamp}{ext}"
    
    csv_file = open(filename, 'w', newline='', encoding='utf-8-sig')
    csv_writer = csv.writer(csv_file)
    
    # Write headers
    headers = [
        'Item ID',
        'Type',
        'Name',
        'Path',
        'URL',
        'Size (MB)',
        'Size (GB)',
        'Created Date',
        'Modified Date',
        'Created By',
        'Created By Email',
        'Modified By',
        'Modified By Email',
        'Version Count'
    ]
    csv_writer.writerow(headers)
    csv_file.flush()
    
    print(f"✅ CSV initialized: {filename}")
    return filename

def append_to_csv(data):
    """Append row to CSV"""
    global csv_writer, csv_file
    
    try:
        row = [
            data.get('id', 0),
            'Folder' if data.get('is_folder') else 'File',
            data.get('name', ''),
            data.get('path', ''),
            data.get('url', ''),
            f"{data.get('size_mb', 0):.2f}",
            f"{data.get('size_gb', 0):.6f}",
            data.get('created_formatted', 'N/A'),
            data.get('modified_formatted', 'N/A'),
            data.get('created_by', ''),
            data.get('created_by_email', ''),
            data.get('modified_by', ''),
            data.get('modified_by_email', ''),
            data.get('version_count', 0)
        ]
        csv_writer.writerow(row)
        csv_file.flush()
        return True
    except Exception as e:
        print(f"❌ Error writing to CSV: {str(e)}")
        return False

def close_csv():
    """Close CSV file"""
    global csv_file
    if csv_file:
        csv_file.close()
        print(f"✅ CSV file closed")

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    """Main function to generate library inventory report"""
    global total_files, total_folders
    
    print("="*80)
    print("📊 SHAREPOINT LIBRARY INVENTORY REPORT")
    print("Generate detailed report of all files and folders")
    print("="*80)
    
    site_url = CONFIG['site_url']
    library_name = CONFIG['library_name']
    
    print(f"\n📌 Site: {site_url}")
    print(f"📌 Library: {library_name}")
    print(f"📌 Output Directory: {CONFIG['output_dir']}")
    print("="*80)
    
    # Authenticate
    print("\n🔐 Authenticating to SharePoint...")
    access_token = get_cached_token()
    
    if not access_token:
        print("❌ Authentication failed!")
        return
    
    print("✅ Authentication successful\n")
    
    # Get library ID
    library_id = get_library_id(site_url, library_name)
    if not library_id:
        print("❌ Library not found!")
        return
    
    # Get all items
    items = get_all_items_from_library(site_url, library_id, library_name)
    
    if not items:
        print("❌ No items found in library!")
        return
    
    # Initialize CSV
    csv_filename = initialize_csv()
    
    print("\n📊 Processing items...")
    print("-"*80)
    
    total_items = 0
    total_size_bytes = 0
    total_versions = 0
    
    for item in items:
        # Get item details
        details = get_item_details(site_url, library_id, item)
        
        # Skip if not configured
        if details['is_folder'] and not CONFIG.get('include_folders', True):
            continue
        if not details['is_folder'] and not CONFIG.get('include_files', True):
            continue
        
        # Append to CSV
        append_to_csv(details)
        
        # Update totals
        total_items += 1
        if details['is_folder']:
            total_folders += 1
        else:
            total_files += 1
            total_size_bytes += details['size']
            total_versions += details['version_count']
        
        # Print progress
        if total_items % 100 == 0:
            print(f"  Processed {total_items} items...", end="\r")
    
    # Close CSV
    close_csv()
    
    # Print summary
    print("\n" + "="*80)
    print("📊 REPORT GENERATED SUCCESSFULLY")
    print("="*80)
    print(f"\n📄 Summary:")
    print(f"  Total Items: {total_items:,}")
    print(f"  Files: {total_files:,}")
    print(f"  Folders: {total_folders:,}")
    print(f"  Total Size: {bytes_to_gb(total_size_bytes):.2f} GB")
    print(f"  Total Versions: {total_versions:,}")
    print(f"\n📁 Report saved: {csv_filename}")
    print("="*80)

if __name__ == "__main__":
    main()