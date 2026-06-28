import requests
import json
import csv
import uuid
import base64
import time
import os
from datetime import datetime
import re
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend

# ============================================================
# CONFIGURATION - UPDATE THESE VALUES
# ============================================================

CONFIG = {
    "site_url": "https://geekbyteonline.sharepoint.com/sites/Team_",
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "scope": "https://geekbyteonline.sharepoint.com/.default",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    
    # ============================================================
    # VERSION HISTORY FILTER - Only check versions for files above this size (in MB)
    # ============================================================
    "min_file_size_mb": 200,  # Only check version history for files > 200 MB
    
    # Output filename is generated from the site name prefix if None.
    "output_csv": None
}

# File extension filter: set to None for all files, or use a single line like ["docx", "pdf"]
FILE_EXTENSIONS = None
FILE_EXTENSIONS = ["docx", "pdf", "xlsx"]

# Token cache
TOKEN_CACHE = {
    "token": None,
    "expires": 0
}

# Global CSV writers and file handles
csv_writers = {}
csv_files = {}

# ============================================================
# AUTHENTICATION FUNCTIONS
# ============================================================

def load_certificate_and_key():
    """Load certificate and private key from PEM files"""
    try:
        if not os.path.exists(CONFIG['certificate_path']) or not os.path.exists(CONFIG['private_key_path']):
            raise Exception(f"Certificate files not found.")
        
        with open(CONFIG['certificate_path'], "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
        
        with open(CONFIG['private_key_path'], "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())
        
        return certificate, private_key
    except Exception as e:
        print(f"Error loading certificate or private key: {str(e)}")
        raise

def get_jwt_token(certificate, private_key):
    """Generate JWT token using certificate and private key"""
    try:
        now = int(time.time())
        expiration = now + 300
        
        thumbprint = certificate.fingerprint(hashes.SHA1())
        x5t = base64.urlsafe_b64encode(thumbprint).decode('utf-8').replace('=', '')
        
        jwt_header = {
            "alg": "RS256",
            "typ": "JWT",
            "x5t": x5t
        }
        
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
        
        jwt = f"{jwt_unsigned}.{encoded_signature}"
        return jwt
    except Exception as e:
        print(f"Error generating JWT: {str(e)}")
        raise

def get_access_token(jwt, scope):
    """Get access token from Microsoft Identity Platform"""
    url = f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token"
    
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    
    data = {
        "client_id": CONFIG['app_id'],
        "client_assertion": jwt,
        "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        "scope": scope,
        "grant_type": "client_credentials"
    }
    
    try:
        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()
        return response.json()["access_token"]
    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error: {err}")
        raise
    except Exception as err:
        print(f"Error: {err}")
        raise

def get_cached_token(force_refresh=False):
    """Get cached token if it's still valid, otherwise get a new one"""
    cache = TOKEN_CACHE
    
    if not force_refresh and cache["token"] and cache["expires"] > time.time() + 300:
        return cache["token"]
    
    try:
        print(f"\n  🔄 {'Refreshing' if force_refresh else 'Getting'} access token...")
        certificate, private_key = load_certificate_and_key()
        jwt = get_jwt_token(certificate, private_key)
        token = get_access_token(jwt, CONFIG['scope'])
        
        if token:
            cache["token"] = token
            cache["expires"] = time.time() + 3600
            print(f"  ✓ Token obtained, valid for 1 hour")
            return token
        else:
            print("  ✗ Failed to get token")
            return None
    except Exception as e:
        print(f"  ✗ Authentication failed: {str(e)}")
        return None

def get_current_token():
    """Get current valid token, refreshing if needed"""
    return get_cached_token()

def make_sharepoint_request(url):
    """Make a request to SharePoint REST API with automatic token refresh"""
    max_retries = 2
    
    for attempt in range(max_retries + 1):
        try:
            token = get_current_token()
            if not token:
                print(f"  ✗ No valid token available")
                return None
            
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json"
            }
            
            response = requests.get(url, headers=headers)
            
            if response.status_code == 401 and attempt < max_retries:
                print(f"  ⚠️ Token expired, refreshing...")
                TOKEN_CACHE["token"] = None
                TOKEN_CACHE["expires"] = 0
                continue
            
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 401 and attempt < max_retries:
                print(f"  ⚠️ Token expired, refreshing...")
                TOKEN_CACHE["token"] = None
                TOKEN_CACHE["expires"] = 0
                continue
            print(f"  Request failed: {str(e)}")
            return None
        except requests.exceptions.RequestException as e:
            print(f"  Request failed: {str(e)}")
            if attempt < max_retries:
                time.sleep(2)
                continue
            return None
    
    return None

def get_site_url():
    """Get the site URL from CONFIG"""
    return CONFIG['site_url'].rstrip('/')


def get_site_prefix(site_url):
    """Extract site prefix from a SharePoint URL"""
    normalized = site_url.rstrip('/')
    parts = normalized.split('/')
    if 'sites' in parts:
        idx = parts.index('sites')
        if idx + 1 < len(parts):
            return parts[idx + 1]
    if parts:
        return parts[-1]
    return 'Site'


def get_report_filename(site_url):
    """Create output filename using the site prefix"""
    site_prefix = get_site_prefix(site_url)
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    return f"{site_prefix}_File_Version_History_Report_{timestamp}.csv"


def normalize_extensions(extensions):
    """Normalize configured file extensions for comparison"""
    if not extensions:
        return None
    if isinstance(extensions, str):
        extensions = [extensions]
    normalized = []
    for ext in extensions:
        if not ext:
            continue
        ext = ext.lower().strip()
        if ext.startswith('.'):
            ext = ext[1:]
        if ext:
            normalized.append(ext)
    return normalized or None


def should_process_file(file_name):
    """Decide whether a file matches the configured extension filter"""
    if not ALLOWED_FILE_EXTENSIONS:
        return True
    _, ext = os.path.splitext(file_name or '')
    ext = ext.lower().lstrip('.')
    return ext in ALLOWED_FILE_EXTENSIONS

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def safe_int_conversion(value):
    """Safely convert a value to integer"""
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
    """Convert bytes to MB with 2 decimal places"""
    bytes_value = safe_int_conversion(bytes_value)
    if bytes_value == 0:
        return 0.00
    return round(bytes_value / (1024 * 1024), 2)

def format_datetime(datetime_str):
    """Format datetime string to readable format"""
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

def should_check_versions(file_size_mb):
    """Check if file size meets the minimum threshold for version checking"""
    min_size = CONFIG.get('min_file_size_mb', 200)
    return file_size_mb > min_size

# ============================================================
# DYNAMIC CSV REPORT FUNCTIONS
# ============================================================

def initialize_reports(output_file):
    """Initialize CSV files with headers"""
    global csv_writers, csv_files
    
    # Main Report
    main_file = output_file
    main_fieldnames = [
        'Library', 'List ID', 'Item ID', 'File Name', 'File Path', 'Current File Size (MB)',
        'Version Count', 'First Version Date', 'Last Version Date', 
        'Total Versions Size (MB)', 'File Created', 'File Modified', 
        'Versions Checked', 'Processed At'
    ]
    
    csv_files['main'] = open(main_file, 'w', newline='', encoding='utf-8-sig')
    csv_writers['main'] = csv.DictWriter(csv_files['main'], fieldnames=main_fieldnames)
    csv_writers['main'].writeheader()
    csv_files['main'].flush()
    print(f"✓ Main report initialized: {main_file}")
    print(f"  Version history will only be checked for files > {CONFIG['min_file_size_mb']} MB")

def append_to_main_report(data):
    """Append a row to the main report"""
    global csv_writers, csv_files
    
    try:
        row = {
            'Library': data.get('library', ''),
            'List ID': data.get('list_id', ''),
            'Item ID': data.get('item_id', 0),
            'File Name': data.get('file_name', ''),
            'File Path': data.get('file_path', ''),
            'Current File Size (MB)': f"{data.get('current_file_size_mb', 0.00):.2f}",
            'Version Count': data.get('version_count', 0),
            'First Version Date': data.get('first_version_formatted', 'N/A'),
            'Last Version Date': data.get('last_version_formatted', 'N/A'),
            'Total Versions Size (MB)': f"{data.get('total_versions_size_mb', 0.00):.2f}",
            'File Created': data.get('created_formatted', 'N/A'),
            'File Modified': data.get('modified_formatted', 'N/A'),
            'Versions Checked': 'Yes' if data.get('versions_checked', False) else 'No',
            'Processed At': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        csv_writers['main'].writerow(row)
        csv_files['main'].flush()
        return True
    except Exception as e:
        print(f"Error appending to main report: {str(e)}")
        return False

def close_reports():
    """Close all CSV files"""
    global csv_files
    
    for key, file_handle in csv_files.items():
        try:
            file_handle.close()
        except:
            pass

# ============================================================
# SHAREPOINT DATA RETRIEVAL FUNCTIONS
# ============================================================

def get_all_libraries(site_url):
    """Get all document libraries from SharePoint site with pagination"""
    print("\nGetting document libraries...")
    lists_url = f"{site_url}/_api/web/lists"
    all_libraries = []
    next_url = lists_url
    
    while next_url:
        print(f"  Fetching libraries page...")
        response = make_sharepoint_request(next_url)
        
        if not response or 'd' not in response:
            break
        
        if 'results' in response['d']:
            for lst in response['d']['results']:
                if lst['BaseTemplate'] == 101:
                    all_libraries.append({
                        'id': lst['Id'],
                        'title': lst['Title']
                    })
        
        next_url = None
        if '__next' in response.get('d', {}):
            next_url = response['d']['__next']
    
    return all_libraries

def get_all_items_from_library(site_url, library_id):
    """Get all items from a library with pagination"""
    print(f"    Fetching items from library...")
    
    items_url = f"{site_url}/_api/web/lists(guid'{library_id}')/items?$expand=File"
    all_items = []
    next_url = items_url
    page_count = 0
    
    while next_url:
        page_count += 1
        print(f"    Fetching page {page_count}...", end="")
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
    
    print(f"    Total items fetched: {len(all_items)}")
    return all_items

def get_file_versions(site_url, list_id, item_id):
    """Get versions for a specific item"""
    try:
        versions_url = f"{site_url}/_api/Web/Lists(guid'{list_id}')/items({item_id})/versions"
        
        response = make_sharepoint_request(versions_url)
        
        if not response or 'd' not in response:
            return []
        
        if 'results' not in response['d']:
            return []
        
        versions = []
        for version in response['d']['results']:
            version_data = {
                'version_id': version.get('VersionId', 0),
                'version_label': version.get('VersionLabel', ''),
                'ui_version_string': version.get('OData__x005f_UIVersionString', ''),
                'created': version.get('Created', ''),
                'is_current': version.get('IsCurrentVersion', False),
                'size': safe_int_conversion(version.get('File_x005f_x0020_x005f_Size', '0')),
                'checkin_comment': version.get('OData__x005f_CheckinComment', ''),
                'author': version.get('Author', {}).get('LookupValue', '') if version.get('Author') else '',
                'editor': version.get('Editor', {}).get('LookupValue', '') if version.get('Editor') else ''
            }
            versions.append(version_data)
        
        return versions
        
    except Exception as e:
        print(f"    Error getting versions for item {item_id}: {str(e)}")
        return []

def get_file_details_from_item(item):
    """Extract file details from item with expanded File property"""
    file_obj = item.get('File', {})
    
    file_name = file_obj.get('Name', '')
    if not file_name:
        file_name = item.get('Title', f"Item_{item.get('Id', 0)}")
    
    file_path = file_obj.get('ServerRelativeUrl', '')
    if not file_path:
        file_path = item.get('FileRef', '')
    
    file_size = file_obj.get('Length', 0)
    if not file_size:
        file_size = item.get('File_x005f_x0020_x005f_Size', '0')
    
    return {
        'file_name': file_name,
        'file_path': file_path,
        'file_size': safe_int_conversion(file_size),
        'created': item.get('Created', 'N/A'),
        'modified': item.get('Modified', 'N/A')
    }

def process_file_item(site_url, list_id, item, library_title):
    """Process a single item and update reports dynamically"""
    try:
        item_id = item.get('Id')
        fsob_type = item.get('FileSystemObjectType', 1)
        
        if fsob_type == 1:
            return None
        
        file_details = get_file_details_from_item(item)
        file_size_mb = bytes_to_mb(file_details['file_size'])
        
        # Check if we should check versions based on file size
        check_versions = should_check_versions(file_size_mb)
        
        versions = []
        version_count = 0
        total_versions_size = 0
        first_version_date = 'N/A'
        last_version_date = 'N/A'
        
        if check_versions:
            # Only get versions for large files
            versions = get_file_versions(site_url, list_id, item_id)
            version_count = len(versions)
            
            if versions:
                sorted_versions = sorted(versions, key=lambda x: x.get('created', ''))
                first_version = sorted_versions[0]
                last_version = sorted_versions[-1]
                
                first_version_date = first_version.get('created', 'N/A')
                last_version_date = last_version.get('created', 'N/A')
                
                for version in versions:
                    total_versions_size += version.get('size', 0)
        
        current_file_size = file_details['file_size']
        if version_count == 0:
            total_versions_size = current_file_size
        
        file_data = {
            'library': library_title,
            'list_id': list_id,
            'item_id': item_id,
            'file_name': file_details['file_name'],
            'file_path': file_details['file_path'],
            'current_file_size': current_file_size,
            'current_file_size_mb': file_size_mb,
            'version_count': version_count,
            'first_version_date': first_version_date,
            'last_version_date': last_version_date,
            'total_versions_size': total_versions_size,
            'total_versions_size_mb': bytes_to_mb(total_versions_size),
            'versions': versions,
            'versions_checked': check_versions,
            'created_formatted': format_datetime(file_details['created']),
            'modified_formatted': format_datetime(file_details['modified']),
            'first_version_formatted': format_datetime(first_version_date),
            'last_version_formatted': format_datetime(last_version_date)
        }
        
        append_to_main_report(file_data)
        return file_data
        
    except Exception as e:
        print(f"    Error processing item: {str(e)}")
        return None

# ============================================================
# MAIN PROCESSING FUNCTIONS
# ============================================================

def process_files(site_url, output_file):
    """Process all files and update reports dynamically"""
    
    initialize_reports(output_file)
    
    libraries = get_all_libraries(site_url)
    
    if not libraries:
        print("No document libraries found.")
        close_reports()
        return []
    
    print(f"\nFound {len(libraries)} document libraries:")
    for lib in libraries:
        print(f"  - {lib['title']}")
    
    print(f"\nVersion history will be checked for files > {CONFIG['min_file_size_mb']} MB only")
    print(f"Smaller files will be reported with Version Count = 0\n")
    
    all_file_data = []
    total_files = 0
    processed = 0
    skipped_by_extension = 0
    
    for library in libraries:
        print(f"\n{'='*60}")
        print(f"Processing library: {library['title']}")
        print(f"{'='*60}")
        
        items = get_all_items_from_library(site_url, library['id'])
        
        if not items:
            print(f"  No items found in {library['title']}")
            continue
        
        files = [item for item in items if item.get('FileSystemObjectType') == 0]
        
        if not files:
            print(f"  No files found in {library['title']}")
            continue
        
        valid_files = [f for f in files if should_process_file(f.get('File', {}).get('Name', f"Item_{f.get('Id', 0)}"))]
        filtered_out = len(files) - len(valid_files)
        skipped_by_extension += filtered_out
        
        print(f"  Found {len(files)} files in {library['title']}")
        print(f"  - Files matching extension filter: {len(valid_files)}")
        print(f"  - Files skipped by extension filter: {filtered_out}")
        
        total_files += len(valid_files)
        
        for file_item in valid_files:
            processed += 1
            item_id = file_item.get('Id')
            file_obj = file_item.get('File', {})
            file_name = file_obj.get('Name', f'Item_{item_id}')
            file_size_mb = bytes_to_mb(file_obj.get('Length', 0))
            
            # Show size category in progress
            size_indicator = "🟢" if file_size_mb > CONFIG['min_file_size_mb'] else "⚪"
            
            print(f"\n  [{processed}/{total_files}] {size_indicator} Processing: {file_name} (ID: {item_id}) [{file_size_mb:.2f} MB]", end="")
            
            file_data = process_file_item(
                site_url, 
                library['id'], 
                file_item, 
                library['title']
            )
            
            if file_data:
                all_file_data.append(file_data)
                if file_data.get('versions_checked', False):
                    print(f" ✓ ({file_data['version_count']} versions, {file_data['total_versions_size_mb']:.2f} MB total)")
                else:
                    print(f" ✓ (Version check skipped - file ≤ {CONFIG['min_file_size_mb']} MB)")
            else:
                print(" ✗ (Failed to process)")
            
            time.sleep(0.3)
    
    print(f"\n{'='*60}")
    print(f"Processed {len(all_file_data)} files with version history.")
    
    close_reports()
    
    return all_file_data

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    """Main function"""
    print("="*80)
    print("FILE VERSION HISTORY REPORT GENERATOR")
    print("(Smart version checking - only for large files)")
    print("="*80)
    global ALLOWED_FILE_EXTENSIONS
    ALLOWED_FILE_EXTENSIONS = normalize_extensions(FILE_EXTENSIONS)
    site_url = get_site_url()
    output_file = CONFIG.get('output_csv') or get_report_filename(site_url)
    CONFIG['output_csv'] = output_file

    print(f"SharePoint Site: {CONFIG['site_url']}")
    print(f"Min File Size for Version Check: {CONFIG['min_file_size_mb']} MB")
    print(f"Output File: {output_file}")
    print("="*80)
    
    print("\nAuthenticating to SharePoint...")
    access_token = get_cached_token()
    
    if not access_token:
        print("✗ Authentication failed.")
        return
    
    print("✓ Authentication successful\n")
    
    site_url = get_site_url()
    output_file = CONFIG['output_csv']
    
    print("Starting file processing...")
    print(f"Version history will ONLY be checked for files > {CONFIG['min_file_size_mb']} MB")
    print("Smaller files will be reported with 'Version Count = 0'")
    start_time = time.time()
    
    file_data = process_files(site_url, output_file)
    
    elapsed_time = time.time() - start_time
    print(f"\nProcessing completed in {elapsed_time:.2f} seconds.")
    
    if not file_data:
        print("\nNo files were processed.")
        return
    
    # Print final summary
    print("\n" + "="*80)
    print("PROCESSING COMPLETED SUCCESSFULLY!")
    print("="*80)
    
    # Count how many files had versions checked
    checked_files = sum(1 for f in file_data if f.get('versions_checked', False))
    skipped_files = len(file_data) - checked_files
    
    print(f"Total files processed: {len(file_data)}")
    print(f"Files with version check: {checked_files}")
    print(f"Files skipped (≤ {CONFIG['min_file_size_mb']} MB): {skipped_files}")
    print(f"Total versions found: {sum(f['version_count'] for f in file_data)}")
    print("="*80)
    print(f"✓ Main Report: {output_file}")
    print("="*80)

if __name__ == "__main__":
    main()