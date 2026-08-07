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
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
from collections import defaultdict

# ============================================================
# CONFIGURATION - UPDATE THESE VALUES
# ============================================================

# Common authentication credentials for all sites
TENANT_ID = "0e439a1f-a497-462b-9e6b-4e582e203607"
APP_ID = "73efa35d-6188-42d4-b258-838a977eb149"
SCOPE = "https://geekbyteonline.sharepoint.com/.default"

# List of SharePoint sites to process (only site URLs needed)
SITES = [
    "https://geekbyteonline.sharepoint.com/sites/Team_",
    "https://geekbyteonline.sharepoint.com/sites/Site1",
    "https://geekbyteonline.sharepoint.com/sites/Site2",
    # Add more sites here
]

CONFIG = {
    "tenant_id": TENANT_ID,
    "app_id": APP_ID,
    "scope": SCOPE,
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    
    # ============================================================
    # VERSION HISTORY FILTER - Only check versions for files above this size (in MB)
    # ============================================================
    "min_file_size_mb": 200,  # Only check version history for files > 200 MB
    
    # ============================================================
    # VERSION RETENTION SETTINGS - Multiple policies
    # ============================================================
    "keep_versions_options": [20, 50, 100],  # Different version retention policies
    
    # ============================================================
    # PERFORMANCE SETTINGS
    # ============================================================
    "batch_size": 50,  # Number of files to process in parallel
    "max_workers": 10,  # Maximum concurrent threads
    "request_timeout": 60,  # Request timeout in seconds
    
    # Output directory for reports
    "output_dir": "reports"
}

# File extension filter: set to None for all files, or use a list like ["docx", "pdf", "xlsx"]
FILE_EXTENSIONS = ["docx", "pdf", "xlsx"]

# ============================================================
# GLOBAL VARIABLES
# ============================================================

TOKEN_CACHE = {"token": None, "expires": 0}
csv_writers = {}
csv_files = {}
csv_lock = threading.Lock()
ALLOWED_FILE_EXTENSIONS = None

# Global summary statistics for all sites
GLOBAL_SUMMARY = {
    "total_sites": 0,
    "total_files": 0,
    "total_files_checked": 0,
    "total_current_size_gb": 0.0,
    "total_versions_size_gb": 0.0,
    "total_versions_count": 0,
    "total_versions_to_delete_count": 0,
    "total_versions_to_delete_gb": 0.0,
    "total_versions_to_keep_gb": 0.0,
    "sites_data": []
}

summary_lock = threading.Lock()

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

def get_access_token(jwt):
    """Get access token from Microsoft Identity Platform"""
    print("  🔑 Requesting access token...")
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
        response = requests.post(url, headers=headers, data=data, timeout=CONFIG['request_timeout'])
        response.raise_for_status()
        result = response.json()
        print("  ✅ Access token obtained")
        return result["access_token"]
    except Exception as e:
        print(f"  ❌ Error getting token: {str(e)}")
        raise

def get_cached_token(force_refresh=False):
    """Get cached token if it's still valid, otherwise get a new one"""
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
    """Get current valid token, refreshing if needed"""
    return get_cached_token()

def make_sharepoint_request(site_url, url, max_retries=2):
    """Make a request to SharePoint REST API with automatic token refresh"""
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
            
            response = requests.get(url, headers=headers, timeout=CONFIG['request_timeout'])
            
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
            if attempt < max_retries:
                time.sleep(2 ** attempt)
                continue
            print(f"  Request failed: {str(e)}")
            return None
        except requests.exceptions.RequestException as e:
            if attempt < max_retries:
                time.sleep(2 ** attempt)
                continue
            print(f"  Request failed: {str(e)}")
            return None
    
    return None

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

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

def get_report_filename(site_url, keep_versions=None):
    """Create output filename using the site prefix and version policy"""
    site_prefix = get_site_prefix(site_url)
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    if keep_versions:
        return f"{site_prefix}_Keep{keep_versions}versions_Report_{timestamp}.csv"
    return f"{site_prefix}_File_Version_Report_{timestamp}.csv"

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

def bytes_to_gb(bytes_value):
    """Convert bytes to GB with 2 decimal places"""
    bytes_value = safe_int_conversion(bytes_value)
    if bytes_value == 0:
        return 0.00
    return round(bytes_value / (1024 * 1024 * 1024), 2)

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

def calculate_version_space_savings(versions, keep_last_n):
    """
    Calculate space savings if we keep only the last N versions
    """
    if not versions:
        return {
            'total_versions': 0,
            'keep_count': 0,
            'delete_count': 0,
            'keep_size_bytes': 0,
            'delete_size_bytes': 0,
            'space_saved_gb': 0.0,
            'space_saved_mb': 0.0,
            'delete_range': 'N/A'
        }
    
    sorted_versions = sorted(versions, key=lambda x: x.get('created', ''))
    total_versions = len(sorted_versions)
    
    if total_versions <= keep_last_n:
        keep_count = total_versions
        delete_count = 0
        keep_size_bytes = sum(v.get('size', 0) for v in sorted_versions)
        delete_size_bytes = 0
        delete_range = 'None (already within limit)'
    else:
        keep_count = keep_last_n
        delete_count = total_versions - keep_last_n
        keep_versions = sorted_versions[-keep_last_n:]
        delete_versions = sorted_versions[:-keep_last_n]
        
        keep_size_bytes = sum(v.get('size', 0) for v in keep_versions)
        delete_size_bytes = sum(v.get('size', 0) for v in delete_versions)
        
        first_deleted = delete_versions[0].get('version_label', 'unknown') if delete_versions else 'N/A'
        last_deleted = delete_versions[-1].get('version_label', 'unknown') if delete_versions else 'N/A'
        delete_range = f"{first_deleted} to {last_deleted}" if first_deleted != 'N/A' else 'N/A'
    
    return {
        'total_versions': total_versions,
        'keep_count': keep_count,
        'delete_count': delete_count,
        'keep_size_bytes': keep_size_bytes,
        'delete_size_bytes': delete_size_bytes,
        'space_saved_gb': bytes_to_gb(delete_size_bytes),
        'space_saved_mb': bytes_to_mb(delete_size_bytes),
        'delete_range': delete_range
    }

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
        response = make_sharepoint_request(site_url, next_url)
        
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
    
    items_url = f"{site_url}/_api/web/lists(guid'{library_id}')/items?$expand=File&$top=5000"
    all_items = []
    next_url = items_url
    page_count = 0
    
    while next_url:
        page_count += 1
        print(f"    Fetching page {page_count}...", end="")
        response = make_sharepoint_request(site_url, next_url)
        
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
        
        response = make_sharepoint_request(site_url, versions_url)
        
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

# ============================================================
# BATCH PROCESSING FUNCTIONS
# ============================================================

def process_file_batch(site_url, file_items, library_id, library_title, batch_id, total_batches, keep_versions, output_file):
    """
    Process a batch of files in parallel
    """
    results = []
    
    with ThreadPoolExecutor(max_workers=CONFIG['max_workers']) as executor:
        future_to_file = {
            executor.submit(
                process_single_file, 
                site_url, 
                library_id, 
                file_item, 
                library_title,
                keep_versions,
                output_file
            ): file_item
            for file_item in file_items
        }
        
        for future in as_completed(future_to_file):
            try:
                result = future.result()
                if result:
                    results.append(result)
            except Exception as e:
                pass
    
    return results

def process_single_file(site_url, list_id, item, library_title, keep_versions, output_file):
    """
    Process a single file item - thread-safe version
    """
    try:
        item_id = item.get('Id')
        fsob_type = item.get('FileSystemObjectType', 1)
        
        if fsob_type == 1:
            return None
        
        file_details = get_file_details_from_item(item)
        file_size_mb = bytes_to_mb(file_details['file_size'])
        
        check_versions = should_check_versions(file_size_mb)
        
        versions = []
        version_count = 0
        total_versions_size = 0
        first_version_date = 'N/A'
        last_version_date = 'N/A'
        versions_to_delete = 0
        space_saved_gb = 0.0
        delete_range = 'N/A'
        
        if check_versions:
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
                
                savings = calculate_version_space_savings(versions, keep_versions)
                
                versions_to_delete = savings['delete_count']
                space_saved_gb = savings['space_saved_gb']
                delete_range = savings['delete_range']
        
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
            'last_version_formatted': format_datetime(last_version_date),
            'versions_to_delete': versions_to_delete,
            'space_saved_gb': space_saved_gb,
            'delete_range': delete_range,
            'keep_versions': keep_versions
        }
        
        # Append to CSV
        append_to_report(output_file, file_data)
        
        return file_data
        
    except Exception as e:
        return None

# ============================================================
# CSV REPORT FUNCTIONS
# ============================================================

def initialize_report(output_file, keep_versions):
    """Initialize CSV file with headers"""
    global csv_writers, csv_files
    
    fieldnames = [
        'Library', 'List ID', 'Item ID', 'File Name', 'File Path', 'Current File Size (MB)',
        'Version Count', 'First Version Date', 'Last Version Date', 
        'Total Versions Size (MB)', 'File Created', 'File Modified', 
        'Versions Checked', 'Versions to Delete', 'Space Saved (GB)',
        'Deleted Version Range', f'Keep Last {keep_versions} Versions', 'Processed At'
    ]
    
    csv_files[output_file] = open(output_file, 'w', newline='', encoding='utf-8-sig')
    csv_writers[output_file] = csv.DictWriter(csv_files[output_file], fieldnames=fieldnames)
    csv_writers[output_file].writeheader()
    csv_files[output_file].flush()

def append_to_report(output_file, data):
    """Append a row to the report - thread-safe"""
    global csv_writers, csv_files
    
    with csv_lock:
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
                'Versions to Delete': data.get('versions_to_delete', 0),
                'Space Saved (GB)': f"{data.get('space_saved_gb', 0.00):.2f}",
                'Deleted Version Range': data.get('delete_range', 'N/A'),
                f'Keep Last {data.get("keep_versions", 50)} Versions': 'Applied',
                'Processed At': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            csv_writers[output_file].writerow(row)
            csv_files[output_file].flush()
            return True
        except Exception as e:
            print(f"Error appending to CSV: {str(e)}")
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
# SITE PROCESSING FUNCTIONS
# ============================================================

def process_site(site_url, keep_versions):
    """Process a single SharePoint site with a specific version retention policy"""
    site_prefix = get_site_prefix(site_url)
    
    print(f"\n{'='*80}")
    print(f"📊 Processing Site: {site_url}")
    print(f"📌 Version Retention Policy: Keep Last {keep_versions} Versions")
    print(f"{'='*80}")
    
    # Create output directory
    os.makedirs(CONFIG['output_dir'], exist_ok=True)
    
    # Generate output filename
    output_file = os.path.join(CONFIG['output_dir'], get_report_filename(site_url, keep_versions))
    
    # Initialize report
    initialize_report(output_file, keep_versions)
    
    # Get all libraries
    libraries = get_all_libraries(site_url)
    
    if not libraries:
        print("No document libraries found.")
        close_reports()
        return None
    
    print(f"\nFound {len(libraries)} document libraries:")
    for lib in libraries:
        print(f"  - {lib['title']}")
    
    print(f"\n⚡ Performance Settings:")
    print(f"  - Batch size: {CONFIG['batch_size']} files per batch")
    print(f"  - Max workers: {CONFIG['max_workers']} concurrent threads")
    print(f"  - Version check only for files > {CONFIG['min_file_size_mb']} MB")
    print(f"  - Keeping last {keep_versions} versions per file\n")
    
    all_file_data = []
    total_files = 0
    skipped_by_extension = 0
    batch_count = 0
    
    site_stats = {
        'site_url': site_url,
        'site_prefix': site_prefix,
        'keep_versions': keep_versions,
        'total_files': 0,
        'files_checked': 0,
        'total_current_size_bytes': 0,
        'total_versions_size_bytes': 0,
        'total_versions_to_delete_bytes': 0,
        'total_versions_to_keep_bytes': 0,
        'total_versions_count': 0,
        'total_versions_to_delete_count': 0,
        'files_with_more_versions': 0,
        'report_file': output_file
    }
    
    for library in libraries:
        print(f"\n{'='*60}")
        print(f"📁 Processing library: {library['title']}")
        print(f"{'='*60}")
        
        items = get_all_items_from_library(site_url, library['id'])
        
        if not items:
            print(f"  No items found in {library['title']}")
            continue
        
        files = [item for item in items if item.get('FileSystemObjectType') == 0]
        
        if not files:
            print(f"  No files found in {library['title']}")
            continue
        
        valid_files = []
        for f in files:
            file_name = f.get('File', {}).get('Name', f"Item_{f.get('Id', 0)}")
            if should_process_file(file_name):
                valid_files.append(f)
            else:
                skipped_by_extension += 1
        
        print(f"  Found {len(files)} files in {library['title']}")
        print(f"  - Files matching extension filter: {len(valid_files)}")
        print(f"  - Files skipped by extension filter: {len(files) - len(valid_files)}")
        
        total_files += len(valid_files)
        
        batch_size = CONFIG['batch_size']
        total_batches = (len(valid_files) + batch_size - 1) // batch_size
        
        for i in range(0, len(valid_files), batch_size):
            batch_count += 1
            batch = valid_files[i:i + batch_size]
            
            print(f"\n  🚀 Processing Batch {batch_count}/{total_batches} ({len(batch)} files)...")
            
            batch_results = process_file_batch(
                site_url, 
                batch, 
                library['id'], 
                library['title'],
                batch_count,
                total_batches,
                keep_versions,
                output_file
            )
            
            # Update site stats
            for result in batch_results:
                if result:
                    site_stats['total_files'] += 1
                    site_stats['total_current_size_bytes'] += result['current_file_size']
                    site_stats['total_versions_size_bytes'] += result['total_versions_size']
                    site_stats['total_versions_count'] += result['version_count']
                    site_stats['total_versions_to_delete_count'] += result['versions_to_delete']
                    
                    if result['versions_checked']:
                        site_stats['files_checked'] += 1
                    
                    # Calculate delete bytes from space_saved_gb
                    delete_bytes = result['space_saved_gb'] * 1024 * 1024 * 1024
                    site_stats['total_versions_to_delete_bytes'] += int(delete_bytes)
                    
                    if result['versions_to_delete'] > 0:
                        site_stats['files_with_more_versions'] += 1
            
            all_file_data.extend(batch_results)
            print(f"  ✅ Batch {batch_count} completed ({len(batch_results)} files processed)")
    
    print(f"\n{'='*60}")
    print(f"✅ Site processing completed: {site_prefix}")
    print(f"   Processed {len(all_file_data)} files")
    print(f"   Skipped {skipped_by_extension} files due to extension filter")
    print(f"   Report saved: {output_file}")
    
    close_reports()
    
    return site_stats

# ============================================================
# SUMMARY REPORT FUNCTIONS
# ============================================================

def create_summary_report(all_site_stats):
    """Create a comprehensive summary report for all sites"""
    summary_file = os.path.join(CONFIG['output_dir'], f"Summary_All_Sites_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv")
    
    fieldnames = [
        'Site URL', 'Site Prefix', 'Keep Versions', 'Total Files', 'Files Checked',
        'Current Size (GB)', 'Versions Size (GB)', 'Total Versions',
        'Versions to Delete', 'Versions to Keep', 'Space to Save (GB)',
        'Files with > Keep_Versions', 'Report File'
    ]
    
    with open(summary_file, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        
        for stats in all_site_stats:
            row = {
                'Site URL': stats['site_url'],
                'Site Prefix': stats['site_prefix'],
                'Keep Versions': stats['keep_versions'],
                'Total Files': stats['total_files'],
                'Files Checked': stats['files_checked'],
                'Current Size (GB)': f"{bytes_to_gb(stats['total_current_size_bytes']):.2f}",
                'Versions Size (GB)': f"{bytes_to_gb(stats['total_versions_size_bytes']):.2f}",
                'Total Versions': stats['total_versions_count'],
                'Versions to Delete': stats['total_versions_to_delete_count'],
                'Versions to Keep': stats['total_versions_count'] - stats['total_versions_to_delete_count'],
                'Space to Save (GB)': f"{bytes_to_gb(stats['total_versions_to_delete_bytes']):.2f}",
                'Files with > Keep_Versions': stats['files_with_more_versions'],
                'Report File': stats.get('report_file', '')
            }
            writer.writerow(row)
    
    return summary_file

def print_global_summary(all_site_stats):
    """Print a comprehensive global summary"""
    print("\n" + "="*100)
    print("🌍 GLOBAL SUMMARY - ALL SITES")
    print("="*100)
    
    # Group stats by site
    site_groups = defaultdict(list)
    for stats in all_site_stats:
        site_groups[stats['site_url']].append(stats)
    
    # Calculate totals
    total_files = sum(s['total_files'] for s in all_site_stats)
    total_files_checked = sum(s['files_checked'] for s in all_site_stats)
    total_current_size = sum(s['total_current_size_bytes'] for s in all_site_stats)
    total_versions_size = sum(s['total_versions_size_bytes'] for s in all_site_stats)
    total_versions = sum(s['total_versions_count'] for s in all_site_stats)
    total_versions_to_delete = sum(s['total_versions_to_delete_count'] for s in all_site_stats)
    total_space_to_save = sum(s['total_versions_to_delete_bytes'] for s in all_site_stats)
    total_files_with_more_versions = sum(s['files_with_more_versions'] for s in all_site_stats)
    
    print(f"\n📊 OVERALL STATISTICS:")
    print(f"  Total Sites Processed: {len(site_groups)}")
    print(f"  Total Version Policies: {len(CONFIG['keep_versions_options'])} per site")
    print(f"  Total Files Processed: {total_files}")
    print(f"  Files with Version Check: {total_files_checked}")
    
    print(f"\n💾 SIZE STATISTICS:")
    print(f"  Total Current File Size: {bytes_to_gb(total_current_size):.2f} GB")
    print(f"  Total Versions Size: {bytes_to_gb(total_versions_size):.2f} GB")
    if total_current_size > 0:
        print(f"  Version Overhead: {((total_versions_size / total_current_size) * 100):.1f}%")
    
    print(f"\n📄 VERSION STATISTICS:")
    print(f"  Total Versions Found: {total_versions}")
    print(f"  Versions to Delete: {total_versions_to_delete}")
    print(f"  Versions to Keep: {total_versions - total_versions_to_delete}")
    
    print(f"\n💰 SPACE SAVINGS ANALYSIS:")
    print(f"  Total Space to Save: {bytes_to_gb(total_space_to_save):.2f} GB")
    if total_versions_size > 0:
        savings_percentage = (total_space_to_save / total_versions_size) * 100
        print(f"  Savings Percentage: {savings_percentage:.1f}% of version history")
    
    if total_current_size > 0 and total_space_to_save > 0:
        print(f"\n  🎯 POTENTIAL SAVINGS:")
        print(f"     Total Space to Save: {bytes_to_gb(total_space_to_save):.2f} GB")
        print(f"     Equivalent to: {bytes_to_gb(total_space_to_save) / bytes_to_gb(total_current_size):.1f}x current total size")
    
    # Print per-site summary
    print(f"\n📋 PER-SITE SUMMARY:")
    print("-" * 100)
    print(f"{'Site':<30} {'Policy':<10} {'Files':<10} {'Current (GB)':<15} {'Versions':<12} {'To Delete':<12} {'Save (GB)':<12}")
    print("-" * 100)
    
    for site_url, site_stats_list in site_groups.items():
        site_prefix = get_site_prefix(site_url)
        for stats in site_stats_list:
            print(f"{site_prefix[:25]:<30} {stats['keep_versions']:<10} "
                  f"{stats['total_files']:<10} {bytes_to_gb(stats['total_current_size_bytes']):<15.2f} "
                  f"{stats['total_versions_count']:<12} {stats['total_versions_to_delete_count']:<12} "
                  f"{bytes_to_gb(stats['total_versions_to_delete_bytes']):<12.2f}")
    
    print("-" * 100)
    print(f"{'TOTAL':<30} {'':<10} {total_files:<10} {bytes_to_gb(total_current_size):<15.2f} "
          f"{total_versions:<12} {total_versions_to_delete:<12} {bytes_to_gb(total_space_to_save):<12.2f}")
    
    print("="*100)

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    """Main function to process all sites with multiple version policies"""
    print("="*100)
    print("📊 BULK SITE VERSION HISTORY REPORT GENERATOR")
    print("(Multiple Sites & Multiple Version Retention Policies)")
    print("="*100)
    
    global ALLOWED_FILE_EXTENSIONS
    ALLOWED_FILE_EXTENSIONS = normalize_extensions(FILE_EXTENSIONS)
    
    # Create output directory
    os.makedirs(CONFIG['output_dir'], exist_ok=True)
    
    print(f"\n📌 Configuration:")
    print(f"  Total Sites: {len(SITES)}")
    print(f"  Version Policies: {CONFIG['keep_versions_options']}")
    print(f"  Min File Size for Version Check: {CONFIG['min_file_size_mb']} MB")
    print(f"  Output Directory: {CONFIG['output_dir']}")
    print("="*100)
    
    # Authenticate once
    print("\n🔐 Authenticating to SharePoint...")
    access_token = get_cached_token()
    
    if not access_token:
        print("❌ Authentication failed. Please check your credentials.")
        return
    
    print("✅ Authentication successful\n")
    
    all_site_stats = []
    total_combinations = len(SITES) * len(CONFIG['keep_versions_options'])
    current_combination = 0
    
    start_time = time.time()
    
    # Process each site with each version policy
    for site_url in SITES:
        print(f"\n{'#'*100}")
        print(f"📍 SITE: {site_url}")
        print(f"{'#'*100}")
        
        for keep_versions in CONFIG['keep_versions_options']:
            current_combination += 1
            print(f"\n▶️  Processing combination {current_combination}/{total_combinations}")
            
            site_stats = process_site(site_url, keep_versions)
            
            if site_stats:
                all_site_stats.append(site_stats)
            
            # Small delay between policies
            time.sleep(1)
    
    elapsed_time = time.time() - start_time
    
    if not all_site_stats:
        print("\n❌ No sites were processed successfully.")
        return
    
    # Create summary report
    summary_file = create_summary_report(all_site_stats)
    
    # Print global summary
    print_global_summary(all_site_stats)
    
    print(f"\n⏱️ Total processing time: {elapsed_time:.2f} seconds ({elapsed_time/60:.2f} minutes)")
    print(f"\n📁 Output Files:")
    print(f"  - Summary Report: {summary_file}")
    
    # List all generated reports
    print(f"\n  - Individual Site Reports:")
    for stats in all_site_stats:
        if 'report_file' in stats:
            print(f"    • {os.path.basename(stats['report_file'])}")
    
    print("\n" + "="*100)
    print("✅ PROCESSING COMPLETED SUCCESSFULLY!")
    print("="*100)

if __name__ == "__main__":
    main()
