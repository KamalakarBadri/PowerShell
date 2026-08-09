import requests
import json
import csv
import uuid
import base64
import time
import os
from datetime import datetime
import re
import gc
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend

# ============================================================
# LOAD CONFIGURATION
# ============================================================

def load_config():
    """Load configuration from JSON file"""
    config_file = "config.json"
    
    if not os.path.exists(config_file):
        print(f"❌ Config file '{config_file}' not found!")
        print("📌 Please create config.json file")
        sys.exit(1)
    
    with open(config_file, 'r') as f:
        config = json.load(f)
    
    print(f"✅ Configuration loaded from {config_file}")
    return config

CONFIG = load_config()

# ============================================================
# GLOBAL VARIABLES
# ============================================================

TOKEN_CACHE = {"token": None, "expires": 0}
ALLOWED_FILE_EXTENSIONS = None
csv_writer = None
csv_file = None
csv_lock = threading.Lock()
current_csv_rows = 0
csv_file_counter = 1

# Statistics
STATS = {
    "total_files": 0,
    "total_files_checked": 0,
    "total_current_size_bytes": 0,
    "total_versions_size_bytes": 0,
    "total_versions_to_delete_bytes": 0,
    "total_versions_count": 0,
    "total_versions_to_delete_count": 0,
    "files_with_more_versions": 0,
    "rate_limit_errors": 0,
    "retry_count": 0
}
stats_lock = threading.Lock()

# ============================================================
# CSV ROTATION FUNCTIONS (For large datasets)
# ============================================================

def get_csv_filename(site_prefix, counter=1):
    """Generate CSV filename with rotation"""
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    return os.path.join(CONFIG['output']['output_dir'], f"{site_prefix}_Version_Report_Part{counter}_{timestamp}.csv")

def rotate_csv(site_prefix):
    """Rotate CSV file when max rows reached"""
    global csv_writer, csv_file, current_csv_rows, csv_file_counter
    
    with csv_lock:
        # Close current file if exists
        if csv_file:
            csv_file.close()
        
        # Create new file
        csv_file_counter += 1
        filename = get_csv_filename(site_prefix, csv_file_counter)
        csv_file = open(filename, 'w', newline='', encoding='utf-8-sig')
        csv_writer = csv.DictWriter(csv_file, fieldnames=get_csv_headers())
        csv_writer.writeheader()
        csv_file.flush()
        current_csv_rows = 0
        
        print(f"  📄 CSV rotated to: {filename}")

def get_csv_headers():
    """Get CSV headers"""
    base_headers = [
        'Site', 'Library', 'List ID', 'Item ID', 'File Name', 'File Path',
        'Current File Size (MB)', 'Version Count', 'Total Versions Size (MB)',
        'File Created', 'File Modified', 'Versions Checked'
    ]
    
    policy_headers = []
    for keep in CONFIG['version_settings']['keep_versions_options']:
        policy_headers.extend([
            f'Keep_{keep}_Versions_To_Delete',
            f'Keep_{keep}_Space_Saved_(GB)',
            f'Keep_{keep}_Deleted_Range'
        ])
    
    return base_headers + policy_headers + ['Processed At']

def initialize_csv(site_prefix):
    """Initialize CSV file"""
    global csv_writer, csv_file, current_csv_rows, csv_file_counter
    
    os.makedirs(CONFIG['output']['output_dir'], exist_ok=True)
    
    csv_file_counter = 1
    filename = get_csv_filename(site_prefix, csv_file_counter)
    csv_file = open(filename, 'w', newline='', encoding='utf-8-sig')
    csv_writer = csv.DictWriter(csv_file, fieldnames=get_csv_headers())
    csv_writer.writeheader()
    csv_file.flush()
    current_csv_rows = 0
    
    print(f"✅ CSV initialized: {filename}")
    return filename

def append_to_csv(data, site_prefix):
    """Append row to CSV with rotation"""
    global csv_writer, csv_file, current_csv_rows, csv_file_counter
    
    with csv_lock:
        try:
            # Check if rotation needed
            max_rows = CONFIG['output'].get('csv_max_rows', 100000)
            if current_csv_rows >= max_rows:
                rotate_csv(site_prefix)
            
            # Build row
            row = {
                'Site': data.get('site', ''),
                'Library': data.get('library', ''),
                'List ID': data.get('list_id', ''),
                'Item ID': data.get('item_id', 0),
                'File Name': data.get('file_name', ''),
                'File Path': data.get('file_path', ''),
                'Current File Size (MB)': f"{data.get('current_file_size_mb', 0.00):.2f}",
                'Version Count': data.get('version_count', 0),
                'Total Versions Size (MB)': f"{data.get('total_versions_size_mb', 0.00):.2f}",
                'File Created': data.get('created_formatted', 'N/A'),
                'File Modified': data.get('modified_formatted', 'N/A'),
                'Versions Checked': 'Yes' if data.get('versions_checked', False) else 'No'
            }
            
            # Add policy columns
            savings_data = data.get('savings_data', {})
            for keep in CONFIG['version_settings']['keep_versions_options']:
                policy_key = f'Keep_{keep}'
                policy_data = savings_data.get(policy_key, {})
                row[f'{policy_key}_Versions_To_Delete'] = policy_data.get('delete_count', 0)
                row[f'{policy_key}_Space_Saved_(GB)'] = f"{policy_data.get('space_saved_gb', 0.00):.2f}"
                row[f'{policy_key}_Deleted_Range'] = policy_data.get('delete_range', 'N/A')
            
            row['Processed At'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            csv_writer.writerow(row)
            csv_file.flush()
            current_csv_rows += 1
            
            return True
        except Exception as e:
            print(f"❌ Error writing to CSV: {str(e)}")
            return False

def close_csv():
    """Close CSV file"""
    global csv_file
    if csv_file:
        csv_file.close()

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
        response = requests.post(url, headers=headers, data=data, timeout=CONFIG['version_settings']['request_timeout'])
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

def make_sharepoint_request(site_url, url, max_retries=5):
    """Make request with retry logic for 429 errors"""
    global STATS
    
    retry_delay = CONFIG['version_settings'].get('retry_delay_seconds', 3)
    
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
            
            response = requests.get(url, headers=headers, timeout=CONFIG['version_settings']['request_timeout'])
            
            # Handle 429 Too Many Requests
            if response.status_code == 429:
                with stats_lock:
                    STATS["rate_limit_errors"] += 1
                
                wait_time = retry_delay * (attempt + 1)  # 3s, 6s, 9s, 12s, 15s
                print(f"    ⚠️ 429 Too Many Requests! Waiting {wait_time}s... (Attempt {attempt+1}/{max_retries})")
                time.sleep(wait_time)
                
                # Refresh token and retry
                TOKEN_CACHE["token"] = None
                TOKEN_CACHE["expires"] = 0
                continue
            
            if response.status_code == 401 and attempt < max_retries:
                print(f"    ⚠️ Token expired, refreshing...")
                TOKEN_CACHE["token"] = None
                TOKEN_CACHE["expires"] = 0
                with stats_lock:
                    STATS["retry_count"] += 1
                continue
            
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 429:
                continue
            if attempt < max_retries:
                time.sleep(2 ** attempt)
                with stats_lock:
                    STATS["retry_count"] += 1
                continue
            return None
        except Exception as e:
            if attempt < max_retries:
                time.sleep(2 ** attempt)
                with stats_lock:
                    STATS["retry_count"] += 1
                continue
            return None
    
    return None

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def get_site_prefix(site_url):
    normalized = site_url.rstrip('/')
    parts = normalized.split('/')
    if 'sites' in parts:
        idx = parts.index('sites')
        if idx + 1 < len(parts):
            return parts[idx + 1]
    if parts:
        return parts[-1]
    return 'Site'

def normalize_extensions(extensions):
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
    if not ALLOWED_FILE_EXTENSIONS:
        return True
    _, ext = os.path.splitext(file_name or '')
    ext = ext.lower().lstrip('.')
    return ext in ALLOWED_FILE_EXTENSIONS

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

def calculate_version_space_savings(versions, keep_last_n):
    if not versions:
        return {
            'total_versions': 0,
            'keep_count': 0,
            'delete_count': 0,
            'keep_size_bytes': 0,
            'delete_size_bytes': 0,
            'space_saved_gb': 0.0,
            'delete_range': 'N/A'
        }
    
    sorted_versions = sorted(versions, key=lambda x: x.get('created', ''))
    total_versions = len(sorted_versions)
    
    if total_versions <= keep_last_n:
        keep_count = total_versions
        delete_count = 0
        keep_size_bytes = sum(v.get('size', 0) for v in sorted_versions)
        delete_size_bytes = 0
        delete_range = 'None (within limit)'
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
        'delete_range': delete_range
    }

# ============================================================
# SHAREPOINT DATA RETRIEVAL
# ============================================================

def get_all_libraries(site_url):
    print("\n📁 Getting document libraries...")
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
    """Get all items from library with pagination"""
    print(f"    Fetching items from library...")
    
    items_url = f"{site_url}/_api/web/lists(guid'{library_id}')/items?$select=Id,Title,FileLeafRef,FileRef,Created,Modified,FileSystemObjectType&$top=5000"
    
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
    """Get ALL versions with selective fields"""
    try:
        versions_url = f"{site_url}/_api/Web/Lists(guid'{list_id}')/items({item_id})/versions?$select=VersionId,VersionLabel,Created,IsCurrentVersion,File_x005f_x0020_x005f_Size,OData__x005f_UIVersionString,OData__x005f_CheckinComment&$orderby=Created desc"
        
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
                'created': version.get('Created', ''),
                'is_current': version.get('IsCurrentVersion', False),
                'size': safe_int_conversion(version.get('File_x005f_x0020_x005f_Size', 0)),
                'checkin_comment': version.get('OData__x005f_CheckinComment', '')
            }
            versions.append(version_data)
        
        return versions
        
    except Exception as e:
        return []

# ============================================================
# BATCH PROCESSING
# ============================================================

def process_single_file(site_url, site_prefix, list_id, item, library_title):
    """Process a single file"""
    global STATS
    
    try:
        item_id = item.get('Id')
        fsob_type = item.get('FileSystemObjectType', 0)
        
        if fsob_type != 0:
            return None
        
        file_name = item.get('FileLeafRef', '')
        if not file_name:
            file_name = item.get('Title', f"Item_{item_id}")
        
        # Get versions
        versions = get_file_versions(site_url, list_id, item_id)
        version_count = len(versions)
        
        # Get current file size from latest version
        current_file_size = 0
        if versions:
            sorted_versions = sorted(versions, key=lambda x: x.get('created', ''))
            latest_version = sorted_versions[-1]
            current_file_size = latest_version.get('size', 0)
        
        file_size_mb = bytes_to_mb(current_file_size)
        
        # Skip small files
        if file_size_mb <= CONFIG['version_settings']['min_file_size_mb']:
            return None
        
        # Calculate totals
        total_versions_size = 0
        for version in versions:
            total_versions_size += version.get('size', 0)
        
        # Calculate savings for each policy
        savings_data = {}
        for keep_versions in CONFIG['version_settings']['keep_versions_options']:
            savings = calculate_version_space_savings(versions, keep_versions)
            savings_data[f'Keep_{keep_versions}'] = {
                'delete_count': savings['delete_count'],
                'space_saved_gb': savings['space_saved_gb'],
                'delete_range': savings['delete_range']
            }
        
        # Update statistics
        with stats_lock:
            STATS["total_files"] += 1
            STATS["total_current_size_bytes"] += current_file_size
            STATS["total_versions_size_bytes"] += total_versions_size
            STATS["total_versions_count"] += version_count
            
            if version_count > 0:
                STATS["total_files_checked"] += 1
            
            for keep, savings in savings_data.items():
                if savings['delete_count'] > 0:
                    STATS["total_versions_to_delete_count"] += savings['delete_count']
                    STATS["total_versions_to_delete_bytes"] += int(savings['space_saved_gb'] * 1024 * 1024 * 1024)
                    STATS["files_with_more_versions"] += 1
        
        # Prepare data for CSV
        file_data = {
            'site': site_prefix,
            'library': library_title,
            'list_id': list_id,
            'item_id': item_id,
            'file_name': file_name,
            'file_path': item.get('FileRef', ''),
            'current_file_size_mb': file_size_mb,
            'version_count': version_count,
            'total_versions_size_mb': bytes_to_mb(total_versions_size),
            'created_formatted': format_datetime(item.get('Created', 'N/A')),
            'modified_formatted': format_datetime(item.get('Modified', 'N/A')),
            'versions_checked': version_count > 0,
            'savings_data': savings_data
        }
        
        # Write to CSV
        append_to_csv(file_data, site_prefix)
        
        return file_data
        
    except Exception as e:
        return None

def process_batch(batch_items, site_url, site_prefix, library_id, library_title, batch_num):
    """Process a batch of files in parallel"""
    print(f"\n  📦 Processing Batch {batch_num} ({len(batch_items)} files)...")
    
    results = []
    batch_size = CONFIG['batch_settings']['max_workers']
    
    with ThreadPoolExecutor(max_workers=batch_size) as executor:
        futures = {
            executor.submit(
                process_single_file,
                site_url,
                site_prefix,
                library_id,
                item,
                library_title
            ): item
            for item in batch_items
        }
        
        processed = 0
        for future in as_completed(futures):
            try:
                result = future.result()
                if result:
                    results.append(result)
                processed += 1
                if processed % 10 == 0:
                    print(f"    Processed {processed}/{len(batch_items)} files...", end="\r")
            except Exception as e:
                pass
        
        print(f"    ✅ Completed {len(results)} files")
    
    return results

# ============================================================
# SITE PROCESSING
# ============================================================

def process_site(site_url):
    """Process a single SharePoint site"""
    site_prefix = get_site_prefix(site_url)
    
    print(f"\n{'#'*80}")
    print(f"📍 SITE: {site_url}")
    print(f"{'#'*80}")
    
    # Initialize CSV for this site
    initialize_csv(site_prefix)
    
    # Get libraries
    libraries = get_all_libraries(site_url)
    
    if not libraries:
        print("❌ No document libraries found.")
        return False
    
    print(f"\n📚 Found {len(libraries)} document libraries:")
    for lib in libraries:
        print(f"  - {lib['title']}")
    
    print(f"\n⚡ Settings:")
    print(f"  - Batch Size: {CONFIG['batch_settings']['batch_size']}")
    print(f"  - Max Workers: {CONFIG['batch_settings']['max_workers']}")
    print(f"  - Min File Size: {CONFIG['version_settings']['min_file_size_mb']} MB")
    print(f"  - Policies: {CONFIG['version_settings']['keep_versions_options']}")
    print(f"  - CSV Max Rows: {CONFIG['output']['csv_max_rows']}")
    print("="*80)
    
    batch_count = 0
    total_files_processed = 0
    
    for library in libraries:
        print(f"\n📁 Library: {library['title']}")
        
        items = get_all_items_from_library(site_url, library['id'])
        
        if not items:
            print("  No items found")
            continue
        
        files = [item for item in items if item.get('FileSystemObjectType') == 0]
        
        if not files:
            print("  No files found")
            continue
        
        # Filter by extension
        valid_files = []
        for f in files:
            file_name = f.get('FileLeafRef', f"Item_{f.get('Id', 0)}")
            if should_process_file(file_name):
                valid_files.append(f)
        
        print(f"  Found {len(files)} files, {len(valid_files)} matching extension filter")
        
        if not valid_files:
            continue
        
        # Process in batches
        batch_size = CONFIG['batch_settings']['batch_size']
        total_batches = (len(valid_files) + batch_size - 1) // batch_size
        
        for i in range(0, len(valid_files), batch_size):
            batch_count += 1
            batch = valid_files[i:i + batch_size]
            
            results = process_batch(
                batch,
                site_url,
                site_prefix,
                library['id'],
                library['title'],
                batch_count
            )
            
            total_files_processed += len(results)
            
            # Cleanup after each batch
            batch = None
            gc.collect()
    
    close_csv()
    
    print(f"\n✅ Site processing completed!")
    print(f"   Files processed: {total_files_processed}")
    print(f"   Batches: {batch_count}")
    
    return True

# ============================================================
# SUMMARY REPORT
# ============================================================

def print_summary_report():
    """Print final summary"""
    global STATS
    
    print("\n" + "="*80)
    print("📊 FINAL SUMMARY REPORT")
    print("="*80)
    
    print(f"\n📄 FILE STATISTICS:")
    print(f"  Total files processed: {STATS['total_files']}")
    print(f"  Files with version check: {STATS['total_files_checked']}")
    print(f"  Files with > versions than policy: {STATS['files_with_more_versions']}")
    
    print(f"\n💾 SIZE STATISTICS:")
    current_size_gb = bytes_to_gb(STATS['total_current_size_bytes'])
    versions_size_gb = bytes_to_gb(STATS['total_versions_size_bytes'])
    
    print(f"  Total current file size: {current_size_gb:.2f} GB")
    print(f"  Total versions size: {versions_size_gb:.2f} GB")
    if current_size_gb > 0:
        print(f"  Version overhead: {((versions_size_gb / current_size_gb) * 100):.1f}%")
    
    print(f"\n📄 VERSION STATISTICS:")
    print(f"  Total versions found: {STATS['total_versions_count']}")
    print(f"  Versions to delete: {STATS['total_versions_to_delete_count']}")
    print(f"  Versions to keep: {STATS['total_versions_count'] - STATS['total_versions_to_delete_count']}")
    
    print(f"\n💰 SPACE SAVINGS ANALYSIS:")
    versions_to_delete_gb = bytes_to_gb(STATS['total_versions_to_delete_bytes'])
    versions_to_keep_gb = bytes_to_gb(STATS['total_versions_size_bytes'] - STATS['total_versions_to_delete_bytes'])
    
    print(f"  Space occupied by versions to DELETE: {versions_to_delete_gb:.2f} GB")
    print(f"  Space occupied by versions to KEEP: {versions_to_keep_gb:.2f} GB")
    print(f"  🟢 TOTAL SPACE SAVED: {versions_to_delete_gb:.2f} GB")
    
    if STATS['total_versions_size_bytes'] > 0:
        savings_percentage = (STATS['total_versions_to_delete_bytes'] / STATS['total_versions_size_bytes']) * 100
        print(f"  Savings percentage: {savings_percentage:.1f}% of version history")
    
    print(f"\n🚀 PERFORMANCE:")
    print(f"  429 Errors encountered: {STATS['rate_limit_errors']}")
    print(f"  Retries: {STATS['retry_count']}")
    
    print("\n" + "="*80)
    print("✅ PROCESSING COMPLETED SUCCESSFULLY!")
    print("="*80)

def save_summary_csv():
    """Save summary to CSV"""
    global STATS
    
    summary_file = os.path.join(CONFIG['output']['output_dir'], CONFIG['output']['summary_file'])
    
    with open(summary_file, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        
        writer.writerow(['Summary Report'])
        writer.writerow(['Generated At', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
        writer.writerow([])
        
        writer.writerow(['Metric', 'Value'])
        writer.writerow(['Total Files Processed', STATS['total_files']])
        writer.writerow(['Files with Version Check', STATS['total_files_checked']])
        writer.writerow(['Files with > Versions than Policy', STATS['files_with_more_versions']])
        writer.writerow(['Total Current Size (GB)', f"{bytes_to_gb(STATS['total_current_size_bytes']):.2f}"])
        writer.writerow(['Total Versions Size (GB)', f"{bytes_to_gb(STATS['total_versions_size_bytes']):.2f}"])
        writer.writerow(['Total Versions', STATS['total_versions_count']])
        writer.writerow(['Versions to Delete', STATS['total_versions_to_delete_count']])
        writer.writerow(['Space to Save (GB)', f"{bytes_to_gb(STATS['total_versions_to_delete_bytes']):.2f}"])
        writer.writerow(['429 Errors', STATS['rate_limit_errors']])
        writer.writerow(['Retries', STATS['retry_count']])
    
    print(f"📁 Summary saved to: {summary_file}")

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    """Main function"""
    global ALLOWED_FILE_EXTENSIONS
    
    print("="*80)
    print("📊 SHAREPOINT VERSION HISTORY ANALYZER")
    print("(Handles millions of files with CSV rotation)")
    print("="*80)
    
    ALLOWED_FILE_EXTENSIONS = normalize_extensions(CONFIG['file_extensions'])
    
    print(f"\n📌 Configuration:")
    print(f"  Total Sites: {len(CONFIG['sites'])}")
    print(f"  Batch Size: {CONFIG['batch_settings']['batch_size']}")
    print(f"  Max Workers: {CONFIG['batch_settings']['max_workers']}")
    print(f"  CSV Max Rows: {CONFIG['output']['csv_max_rows']}")
    print(f"  Output Directory: {CONFIG['output']['output_dir']}")
    print("="*80)
    
    # Authenticate
    print("\n🔐 Authenticating to SharePoint...")
    access_token = get_cached_token()
    
    if not access_token:
        print("❌ Authentication failed!")
        return
    
    print("✅ Authentication successful\n")
    
    start_time = time.time()
    
    # Process each site
    for site_url in CONFIG['sites']:
        success = process_site(site_url)
        if not success:
            print(f"❌ Failed to process site: {site_url}")
            continue
        
        # Small delay between sites
        time.sleep(2)
    
    elapsed_time = time.time() - start_time
    
    # Print summary
    print_summary_report()
    save_summary_csv()
    
    print(f"\n⏱️ Total processing time: {elapsed_time:.2f} seconds ({elapsed_time/60:.2f} minutes)")
    print(f"📁 Reports saved in: {CONFIG['output']['output_dir']}")

if __name__ == "__main__":
    main()