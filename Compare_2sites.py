import requests
import json
import csv
import uuid
import base64
import time
import os
import sys
import re
import traceback
from datetime import datetime
from collections import defaultdict
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading

# ============================================================
# CONFIGURATION
# ============================================================

SHOW_TRACEBACK = False  # Set to True for debugging
BATCH_SIZE = 50  # Number of files to process in parallel
MAX_WORKERS = 10  # Max parallel threads

def log_error(message, error=None):
    """Log error with optional traceback"""
    print(f"  [ERROR] {message}")
    if SHOW_TRACEBACK and error:
        print("  " + "-"*60)
        traceback.print_exc()
        print("  " + "-"*60)

def print_success(message):
    print(f"[OK] {message}")

def print_error(message):
    print(f"[ERROR] {message}")

def print_warning(message):
    print(f"[WARNING] {message}")

def print_info(message):
    print(f"[INFO] {message}")

def print_header(message):
    print(f"\n{'='*80}")
    print(f" {message}")
    print(f"{'='*80}")

def print_subheader(message):
    print(f"\n{'-'*80}")
    print(f" {message}")
    print(f"{'-'*80}")

# ============================================================
# LOAD CONFIGURATION
# ============================================================

def load_config():
    """Load configuration from JSON file"""
    config_file = "compare_config.json"
    
    if not os.path.exists(config_file):
        print_error(f"Config file '{config_file}' not found!")
        print_info("Please create compare_config.json file")
        sys.exit(1)
    
    with open(config_file, 'r') as f:
        config = json.load(f)
    
    print_success(f"Configuration loaded from {config_file}")
    return config

CONFIG = load_config()

# ============================================================
# GLOBAL VARIABLES
# ============================================================

TOKEN_CACHE = {}
ALLOWED_FILE_EXTENSIONS = None
PROCESSED_COUNT = 0
PROCESSED_LOCK = threading.Lock()

# Progress tracking
PROGRESS = {
    'total_items': 0,
    'processed_items': 0,
    'current_library': '',
    'current_file': '',
    'start_time': time.time()
}

# Comparison results
COMPARISON_RESULTS = {
    "source": {
        "total_files": 0,
        "total_size_bytes": 0,
        "total_versions": 0,
        "files_with_versions": 0,
        "libraries": []
    },
    "destination": {
        "total_files": 0,
        "total_size_bytes": 0,
        "total_versions": 0,
        "files_with_versions": 0,
        "libraries": []
    },
    "differences": {
        "missing_in_destination": [],
        "missing_in_source": [],
        "size_mismatch": [],
        "modified_date_mismatch": [],
        "version_count_mismatch": [],
        "version_editor_mismatch": [],
        "file_name_mismatch": []
    }
}

# ============================================================
# PROGRESS DISPLAY FUNCTIONS
# ============================================================

def update_progress(file_name, library_name):
    """Update progress display"""
    global PROCESSED_COUNT
    with PROCESSED_LOCK:
        PROCESSED_COUNT += 1
        PROGRESS['processed_items'] = PROCESSED_COUNT
        PROGRESS['current_file'] = file_name
        
        if PROCESSED_COUNT % 10 == 0 or PROCESSED_COUNT == PROGRESS['total_items']:
            elapsed = time.time() - PROGRESS['start_time']
            if PROCESSED_COUNT > 0:
                items_per_sec = PROCESSED_COUNT / elapsed if elapsed > 0 else 0
                remaining = (PROGRESS['total_items'] - PROCESSED_COUNT) / items_per_sec if items_per_sec > 0 else 0
                
                progress_pct = (PROCESSED_COUNT / PROGRESS['total_items']) * 100 if PROGRESS['total_items'] > 0 else 0
                print(f"\r  [PROGRESS] {PROCESSED_COUNT}/{PROGRESS['total_items']} ({progress_pct:.1f}%) | "
                      f"Elapsed: {elapsed:.1f}s | ETA: {remaining:.1f}s | "
                      f"Current: {file_name[:30]}...", end="", flush=True)

def print_progress_summary():
    """Print final progress summary"""
    elapsed = time.time() - PROGRESS['start_time']
    print(f"\n\n  [OK] Processing complete in {elapsed:.1f} seconds")
    print(f"  [INFO] Processed {PROGRESS['processed_items']} items")

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
        log_error(f"Error loading certificate: {str(e)}", e)
        raise

def get_jwt_token(certificate, private_key, tenant_id, app_id):
    try:
        now = int(time.time())
        expiration = now + 300
        
        thumbprint = certificate.fingerprint(hashes.SHA1())
        x5t = base64.urlsafe_b64encode(thumbprint).decode('utf-8').replace('=', '')
        
        jwt_header = {"alg": "RS256", "typ": "JWT", "x5t": x5t}
        
        jwt_payload = {
            "aud": f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
            "exp": expiration,
            "iss": app_id,
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": app_id
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
        log_error(f"Error generating JWT: {str(e)}", e)
        raise

def get_access_token(jwt, tenant_id, app_id, scope):
    print("  [KEY] Getting access token...")
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    
    data = {
        "client_id": app_id,
        "client_assertion": jwt,
        "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        "scope": scope,
        "grant_type": "client_credentials"
    }
    
    try:
        response = requests.post(url, headers=headers, data=data, timeout=120)
        response.raise_for_status()
        result = response.json()
        print_success("Token obtained")
        return result["access_token"]
    except Exception as e:
        log_error(f"Error getting token: {str(e)}", e)
        raise

def get_cached_token(site_url, force_refresh=False):
    cache = TOKEN_CACHE
    
    if site_url not in cache:
        cache[site_url] = {"token": None, "expires": 0}
    
    if not force_refresh and cache[site_url]["token"] and cache[site_url]["expires"] > time.time() + 300:
        return cache[site_url]["token"]
    
    try:
        certificate, private_key = load_certificate_and_key()
        
        jwt = get_jwt_token(certificate, private_key, CONFIG['tenant_id'], CONFIG['app_id'])
        token = get_access_token(jwt, CONFIG['tenant_id'], CONFIG['app_id'], CONFIG['scope'])
        
        if token:
            cache[site_url]["token"] = token
            cache[site_url]["expires"] = time.time() + 3600
            return token
        return None
    except Exception as e:
        log_error(f"Authentication failed: {str(e)}", e)
        return None

def get_current_token(site_url):
    return get_cached_token(site_url)

def make_sharepoint_request(site_url, url, max_retries=5):
    """Make request with retry logic"""
    
    for attempt in range(max_retries + 1):
        try:
            token = get_current_token(site_url)
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
                print(f"\n    [WARNING] 429: Waiting {wait_time}s...")
                time.sleep(wait_time)
                continue
            
            if response.status_code == 401 and attempt < max_retries:
                print(f"\n    [WARNING] Token expired, refreshing...")
                TOKEN_CACHE[site_url]["token"] = None
                TOKEN_CACHE[site_url]["expires"] = 0
                continue
            
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 429:
                continue
            if attempt < max_retries:
                time.sleep(2 ** attempt)
                continue
            log_error(f"Request failed: {str(e)}", e)
            return None
        except Exception as e:
            if attempt < max_retries:
                time.sleep(2 ** attempt)
                continue
            log_error(f"Error: {str(e)}", e)
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
            if '.' in datetime_str and 'Z' in datetime_str:
                dt = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S.%fZ")
            elif 'Z' in datetime_str:
                dt = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%SZ")
            else:
                dt = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S")
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        return datetime_str
    except:
        return datetime_str

def get_file_name_without_extension(file_name):
    if not file_name:
        return ""
    return os.path.splitext(file_name)[0]

def get_file_extension(file_name):
    if not file_name:
        return ""
    return os.path.splitext(file_name)[1].lower()

# ============================================================
# SHAREPOINT DATA RETRIEVAL - CORRECTED
# ============================================================

def get_library_id(site_url, library_name):
    """Get library ID by EntityTypeName (not Title)"""
    print(f"\n[SEARCH] Finding library: {library_name}")
    
    # Get all document libraries with EntityTypeName
    url = f"{site_url}/_api/web/lists?$filter=BaseTemplate eq 101&$select=Id,Title,EntityTypeName"
    response = make_sharepoint_request(site_url, url)
    
    if response and 'd' in response and 'results' in response['d']:
        # First try: Match by EntityTypeName
        for lib in response['d']['results']:
            entity_type = lib.get('EntityTypeName', '')
            if entity_type.lower() == library_name.lower():
                print_success(f"Found library by EntityTypeName: '{lib['Title']}' (ID: {lib['Id']})")
                return lib['Id']
        
        # Second try: Match by Title
        for lib in response['d']['results']:
            if lib['Title'].lower() == library_name.lower():
                print_success(f"Found library by Title: '{lib['Title']}' (ID: {lib['Id']})")
                return lib['Id']
        
        # Third try: Partial match
        library_lower = library_name.lower()
        for lib in response['d']['results']:
            if library_lower in lib['Title'].lower() or library_lower in lib.get('EntityTypeName', '').lower():
                print_success(f"Found partial match: '{lib['Title']}' (EntityType: {lib.get('EntityTypeName', 'N/A')})")
                print_info(f"Did you mean: '{lib['Title']}'?")
                return lib['Id']
        
        # List available libraries
        print(f"\n  [INFO] Available document libraries:")
        for lib in response['d']['results'][:15]:
            print(f"    - Title: '{lib['Title']}', EntityType: '{lib.get('EntityTypeName', 'N/A')}'")
        if len(response['d']['results']) > 15:
            print(f"    ... and {len(response['d']['results']) - 15} more")
    
    print_error(f"Library '{library_name}' not found")
    print_info("Please check the exact library name in SharePoint")
    return None

def get_all_items_from_library(site_url, library_id, library_name):
    """Get all items from library with pagination - REMOVED size field"""
    print(f"\n[FOLDER] Fetching items from library: {library_name}")
    
    # CORRECTED: Removed File_x005f_x0020_x005f_Size - get size from versions endpoint
    items_url = f"{site_url}/_api/web/lists(guid'{library_id}')/items?$select=Id,FileLeafRef,FileRef,Created,Modified,FileSystemObjectType&$top=5000"
    
    all_items = []
    next_url = items_url
    page_count = 0
    
    while next_url:
        page_count += 1
        print(f"  Fetching page {page_count}...", end="")
        response = make_sharepoint_request(site_url, next_url)
        
        if not response or 'd' not in response:
            print(" [FAILED]")
            break
        
        if 'results' in response['d']:
            items_in_page = len(response['d']['results'])
            all_items.extend(response['d']['results'])
            print(f" [OK] Got {items_in_page} items")
        
        next_url = None
        if '__next' in response.get('d', {}):
            next_url = response['d']['__next']
    
    print(f"  Total items fetched: {len(all_items)}")
    return all_items

def get_file_versions_complete(site_url, list_id, item_id):
    """
    Get COMPLETE version information from versions endpoint
    This is where we get: Created, Modified, Size, Editor, Author
    """
    try:
        versions_url = (
            f"{site_url}/_api/Web/Lists(guid'{list_id}')/items({item_id})/versions"
            f"?$select=VersionId,VersionLabel,Created,Modified,IsCurrentVersion,"
            f"File_x005f_x0020_x005f_Size,CheckInComment,"
            f"Author/Title,Author/Email,Author/Id,"
            f"Editor/Title,Editor/Email,Editor/Id,"
            f"Created_x005f_x0020_x005f_By,Modified_x005f_x0020_x005f_By"
            f"&$expand=Author,Editor"
            f"&$orderby=Created desc"
        )
        
        response = make_sharepoint_request(site_url, versions_url)
        
        if not response or 'd' not in response:
            return []
        
        versions = []
        for version in response['d'].get('results', []):
            # Get Editor (who modified)
            editor = version.get('Editor', {})
            editor_name = editor.get('LookupValue', '') if isinstance(editor, dict) else ''
            editor_email = editor.get('Email', '') if isinstance(editor, dict) else ''
            editor_id = editor.get('LookupId', 0) if isinstance(editor, dict) else 0
            
            # Get Author/Created By
            author = version.get('Author', {})
            author_name = author.get('LookupValue', '') if isinstance(author, dict) else ''
            author_email = author.get('Email', '') if isinstance(author, dict) else ''
            author_id = author.get('LookupId', 0) if isinstance(author, dict) else 0
            
            # Get direct fields
            created_by = version.get('Created_x005f_x0020_x005f_By', '')
            modified_by = version.get('Modified_x005f_x0020_x005f_By', '')
            
            # Parse file size from versions endpoint (this is where size comes from)
            size_str = version.get('File_x005f_x0020_x005f_Size', '0')
            size = safe_int_conversion(size_str)
            
            version_data = {
                'version_id': version.get('VersionId', 0),
                'version_label': version.get('VersionLabel', ''),
                'is_current': version.get('IsCurrentVersion', False),
                'created': version.get('Created', ''),
                'created_formatted': format_datetime(version.get('Created', '')),
                'modified': version.get('Modified', ''),
                'modified_formatted': format_datetime(version.get('Modified', '')),
                'size': size,
                'size_mb': bytes_to_mb(size),
                'editor_name': editor_name or modified_by,
                'editor_email': editor_email,
                'editor_id': editor_id,
                'author_name': author_name or created_by,
                'author_email': author_email,
                'author_id': author_id,
                'check_in_comment': version.get('CheckInComment', ''),
                'is_current_version': version.get('IsCurrentVersion', False)
            }
            versions.append(version_data)
        
        return versions
        
    except Exception as e:
        log_error(f"Error getting versions for item {item_id}: {str(e)}", e)
        return []

def process_item_batch(site_url, list_id, items, library_name):
    """Process a batch of items in parallel"""
    results = []
    
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_item = {
            executor.submit(get_file_details_with_complete_versions, site_url, list_id, item): item 
            for item in items
        }
        
        for future in as_completed(future_to_item):
            try:
                file_details = future.result()
                if file_details:
                    results.append(file_details)
                update_progress(file_details.get('name', 'Unknown'), library_name)
            except Exception as e:
                log_error(f"Error processing item: {str(e)}", e)
    
    return results

def get_file_details_with_complete_versions(site_url, list_id, item):
    """
    Get file details - size comes from versions endpoint
    """
    # Get basic file info from items endpoint (NO size here)
    file_name = item.get('FileLeafRef', '')
    if not file_name:
        file_name = f"Item_{item.get('Id', 0)}"
    
    file_details = {
        'id': item.get('Id', 0),
        'name': file_name,
        'name_without_ext': get_file_name_without_extension(file_name),
        'extension': get_file_extension(file_name),
        'path': item.get('FileRef', ''),
        # Size will come from versions endpoint
        'size': 0,
        'size_mb': 0,
        'created': item.get('Created', 'N/A'),
        'modified': item.get('Modified', 'N/A'),
        'created_formatted': format_datetime(item.get('Created', 'N/A')),
        'modified_formatted': format_datetime(item.get('Modified', 'N/A')),
        'is_folder': item.get('FileSystemObjectType', 0) == 1
    }
    
    # Skip folders
    if file_details['is_folder']:
        file_details['version_count'] = 0
        file_details['versions'] = []
        file_details['last_modified_by'] = ''
        file_details['last_modified_by_email'] = ''
        file_details['last_modified_date'] = file_details['modified_formatted']
        file_details['version_stats'] = {}
        file_details['unique_editors'] = []
        file_details['editor_count'] = 0
        return file_details
    
    # Get ALL version details from versions endpoint (this includes size)
    versions = get_file_versions_complete(site_url, list_id, item.get('Id'))
    
    # Store complete version data
    file_details['versions'] = versions
    file_details['version_count'] = len(versions)
    
    # Version statistics from versions endpoint
    if versions:
        # Latest version (index 0 since sorted by Created desc)
        latest = versions[0]
        
        # Get size from latest version
        file_details['size'] = latest.get('size', 0)
        file_details['size_mb'] = latest.get('size_mb', 0)
        
        file_details['last_modified_by'] = latest.get('editor_name', '')
        file_details['last_modified_by_email'] = latest.get('editor_email', '')
        file_details['last_modified_date'] = latest.get('modified_formatted', '')
        file_details['last_modified_size'] = latest.get('size_mb', 0)
        
        # Version summary stats
        all_sizes = [v['size'] for v in versions]
        file_details['version_stats'] = {
            'min_size': bytes_to_mb(min(all_sizes)) if all_sizes else 0,
            'max_size': bytes_to_mb(max(all_sizes)) if all_sizes else 0,
            'avg_size': bytes_to_mb(sum(all_sizes) // len(all_sizes)) if all_sizes else 0,
            'total_versions_size': bytes_to_mb(sum(all_sizes)) if all_sizes else 0
        }
        
        # Get unique editors
        editors = set()
        for v in versions:
            if v['editor_name']:
                editors.add(v['editor_name'])
        file_details['unique_editors'] = list(editors)
        file_details['editor_count'] = len(editors)
    else:
        # No versions - size remains 0
        file_details['last_modified_by'] = ''
        file_details['last_modified_by_email'] = ''
        file_details['last_modified_date'] = file_details['modified_formatted']
        file_details['version_stats'] = {}
        file_details['unique_editors'] = []
        file_details['editor_count'] = 0
    
    return file_details

# ============================================================
# BUILD FILE MAP WITH BATCH PROCESSING
# ============================================================

def build_file_map(items, site_url, list_id, library_name):
    """Build a map of files by name with details using batch processing"""
    global PROCESSED_COUNT
    
    file_map = {}
    total_size = 0
    total_versions = 0
    files_with_versions = 0
    folders = 0
    files_count = 0
    
    print(f"\n[BUILD] Building file map for {library_name}...")
    print(f"[INFO] Total items: {len(items)}")
    print(f"[INFO] Processing in batches of {BATCH_SIZE} with {MAX_WORKERS} parallel threads")
    print(f"[INFO] File size will be retrieved from versions endpoint")
    
    # Reset progress
    PROCESSED_COUNT = 0
    PROGRESS['total_items'] = len(items)
    PROGRESS['processed_items'] = 0
    PROGRESS['current_library'] = library_name
    PROGRESS['start_time'] = time.time()
    
    # Process items in batches
    for i in range(0, len(items), BATCH_SIZE):
        batch = items[i:i+BATCH_SIZE]
        print(f"\n  [BATCH] Processing batch {i//BATCH_SIZE + 1}/{(len(items)-1)//BATCH_SIZE + 1} ({len(batch)} items)")
        
        # Process batch in parallel
        batch_results = process_item_batch(site_url, list_id, batch, library_name)
        
        # Add results to file map
        for file_details in batch_results:
            if file_details['is_folder']:
                folders += 1
                continue
            
            key = file_details['name_without_ext'].lower()
            
            if key not in file_map:
                file_map[key] = []
            
            file_map[key].append(file_details)
            files_count += 1
            
            total_size += file_details['size']
            total_versions += file_details['version_count']
            if file_details['version_count'] > 0:
                files_with_versions += 1
    
    print()
    
    print(f"  Files: {files_count}")
    print(f"  Folders: {folders}")
    print(f"  Total Size: {bytes_to_gb(total_size):.2f} GB")
    print(f"  Total Versions: {total_versions}")
    print(f"  Files with versions: {files_with_versions}")
    
    return file_map, total_size, total_versions, files_with_versions

# ============================================================
# COMPARISON FUNCTIONS
# ============================================================

def compare_libraries(source_map, dest_map, source_site, dest_site):
    """Compare source and destination file maps"""
    print_header("COMPARING SOURCE vs DESTINATION")
    
    differences = {
        "missing_in_destination": [],
        "missing_in_source": [],
        "size_mismatch": [],
        "modified_date_mismatch": [],
        "version_count_mismatch": [],
        "version_editor_mismatch": [],
        "file_name_mismatch": [],
        "matched_files": []
    }
    
    source_keys = set(source_map.keys())
    dest_keys = set(dest_map.keys())
    
    missing_in_dest = source_keys - dest_keys
    missing_in_source = dest_keys - source_keys
    
    print(f"\n[INFO] Files missing in Destination: {len(missing_in_dest)}")
    print(f"[INFO] Files missing in Source: {len(missing_in_source)}")
    
    common_keys = source_keys & dest_keys
    print(f"\n[INFO] Common files: {len(common_keys)}")
    
    print_subheader("DETAILED COMPARISON - COMMON FILES")
    
    total_common = len(common_keys)
    processed = 0
    
    for key in sorted(common_keys):
        processed += 1
        if processed % 50 == 0 or processed == total_common:
            print(f"\r  Comparing: {processed}/{total_common} files...", end="", flush=True)
        
        source_files = source_map[key]
        dest_files = dest_map[key]
        
        if len(source_files) == 1 and len(dest_files) == 1:
            source_file = source_files[0]
            dest_file = dest_files[0]
            
            if source_file['extension'] != dest_file['extension']:
                differences['file_name_mismatch'].append({
                    'name': key,
                    'source': source_file,
                    'destination': dest_file,
                    'issue': 'Extension mismatch'
                })
            
            if source_file['size'] != dest_file['size']:
                differences['size_mismatch'].append({
                    'name': key,
                    'source': source_file,
                    'destination': dest_file,
                    'source_size_mb': source_file['size_mb'],
                    'dest_size_mb': dest_file['size_mb'],
                    'diff_mb': source_file['size_mb'] - dest_file['size_mb']
                })
            
            if source_file['modified'] != dest_file['modified']:
                differences['modified_date_mismatch'].append({
                    'name': key,
                    'source': source_file,
                    'destination': dest_file,
                    'source_modified': source_file['modified_formatted'],
                    'dest_modified': dest_file['modified_formatted']
                })
            
            if source_file['version_count'] != dest_file['version_count']:
                differences['version_count_mismatch'].append({
                    'name': key,
                    'source': source_file,
                    'destination': dest_file,
                    'source_versions': source_file['version_count'],
                    'dest_versions': dest_file['version_count']
                })
            
            if CONFIG['comparison_settings'].get('check_version_editor', True):
                source_editors = set()
                dest_editors = set()
                
                for v in source_file.get('versions', []):
                    if v.get('editor_name'):
                        source_editors.add(v['editor_name'])
                
                for v in dest_file.get('versions', []):
                    if v.get('editor_name'):
                        dest_editors.add(v['editor_name'])
                
                if source_editors != dest_editors:
                    differences['version_editor_mismatch'].append({
                        'name': key,
                        'source': source_file,
                        'destination': dest_file,
                        'source_editors': list(source_editors),
                        'dest_editors': list(dest_editors),
                        'source_version_count': len(source_file.get('versions', [])),
                        'dest_version_count': len(dest_file.get('versions', []))
                    })
            
            if (source_file['size'] == dest_file['size'] and 
                source_file['modified'] == dest_file['modified'] and
                source_file['version_count'] == dest_file['version_count']):
                differences['matched_files'].append({
                    'name': key,
                    'source': source_file,
                    'destination': dest_file
                })
        
        else:
            differences['file_name_mismatch'].append({
                'name': key,
                'source_count': len(source_files),
                'dest_count': len(dest_files),
                'source_files': [f['name'] for f in source_files],
                'dest_files': [f['name'] for f in dest_files],
                'issue': 'Multiple files with same name'
            })
    
    print()
    
    for key in missing_in_dest:
        for file in source_map[key]:
            differences['missing_in_destination'].append(file)
    
    for key in missing_in_source:
        for file in dest_map[key]:
            differences['missing_in_source'].append(file)
    
    print_comparison_summary(differences)
    
    return differences

def print_comparison_summary(differences):
    """Print comparison summary"""
    print_header("COMPARISON SUMMARY")
    
    print(f"\n[OK] Matched Files: {len(differences['matched_files'])}")
    
    print(f"\n[ERROR] Missing in Destination: {len(differences['missing_in_destination'])}")
    if differences['missing_in_destination']:
        print("  (Files that exist in Source but not in Destination)")
    
    print(f"\n[ERROR] Missing in Source: {len(differences['missing_in_source'])}")
    if differences['missing_in_source']:
        print("  (Files that exist in Destination but not in Source)")
    
    print(f"\n[WARNING] Size Mismatch: {len(differences['size_mismatch'])}")
    if differences['size_mismatch']:
        print("  (Files with different sizes)")
    
    print(f"\n[WARNING] Modified Date Mismatch: {len(differences['modified_date_mismatch'])}")
    if differences['modified_date_mismatch']:
        print("  (Files with different modified dates)")
    
    print(f"\n[WARNING] Version Count Mismatch: {len(differences['version_count_mismatch'])}")
    if differences['version_count_mismatch']:
        print("  (Files with different version counts)")
    
    print(f"\n[WARNING] Version Editor Mismatch: {len(differences['version_editor_mismatch'])}")
    if differences['version_editor_mismatch']:
        print("  (Files where version editors are different)")
    
    print(f"\n[WARNING] File Name Issues: {len(differences['file_name_mismatch'])}")
    if differences['file_name_mismatch']:
        print("  (Files with extension mismatches or duplicates)")
    
    total_issues = (len(differences['missing_in_destination']) + 
                    len(differences['missing_in_source']) +
                    len(differences['size_mismatch']) +
                    len(differences['modified_date_mismatch']) +
                    len(differences['version_count_mismatch']) +
                    len(differences['version_editor_mismatch']) +
                    len(differences['file_name_mismatch']))
    
    print(f"\n[CRITICAL] TOTAL ISSUES FOUND: {total_issues}")
    print("="*80)

# ============================================================
# CSV REPORT FUNCTIONS (same as before - kept for brevity)
# ============================================================

def save_comparison_report(differences, source_site, dest_site):
    """Save comparison report to CSV"""
    os.makedirs(CONFIG['output']['output_dir'], exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    source_prefix = get_site_prefix(source_site)
    dest_prefix = get_site_prefix(dest_site)
    
    summary_file = os.path.join(CONFIG['output']['output_dir'], 
                                CONFIG['output']['summary_file'])
    
    with open(summary_file, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(['Migration Comparison Report'])
        writer.writerow(['Generated At', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
        writer.writerow(['Source Site', source_site])
        writer.writerow(['Destination Site', dest_site])
        writer.writerow([])
        
        writer.writerow(['Metric', 'Count'])
        writer.writerow(['Matched Files', len(differences['matched_files'])])
        writer.writerow(['Missing in Destination', len(differences['missing_in_destination'])])
        writer.writerow(['Missing in Source', len(differences['missing_in_source'])])
        writer.writerow(['Size Mismatch', len(differences['size_mismatch'])])
        writer.writerow(['Modified Date Mismatch', len(differences['modified_date_mismatch'])])
        writer.writerow(['Version Count Mismatch', len(differences['version_count_mismatch'])])
        writer.writerow(['Version Editor Mismatch', len(differences['version_editor_mismatch'])])
        writer.writerow(['File Name Issues', len(differences['file_name_mismatch'])])
        writer.writerow(['Total Issues', 
                        len(differences['missing_in_destination']) + 
                        len(differences['missing_in_source']) +
                        len(differences['size_mismatch']) +
                        len(differences['modified_date_mismatch']) +
                        len(differences['version_count_mismatch']) +
                        len(differences['version_editor_mismatch']) +
                        len(differences['file_name_mismatch'])])
    
    print(f"\n[FILE] Summary report saved: {summary_file}")
    
    detail_file = os.path.join(CONFIG['output']['output_dir'], 
                               f"{source_prefix}_vs_{dest_prefix}_comparison_details_{timestamp}.csv")
    
    with open(detail_file, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        
        writer.writerow([
            'File Name', 'Status', 
            'Source Size (MB)', 'Dest Size (MB)', 'Size Diff (MB)',
            'Source Modified', 'Dest Modified',
            'Source Versions', 'Dest Versions', 'Version Diff',
            'Source Last Modified By', 'Dest Last Modified By',
            'Source Path', 'Dest Path', 'Issue'
        ])
        
        for file in differences['matched_files']:
            writer.writerow([
                file['name'],
                'MATCHED',
                file['source']['size_mb'],
                file['destination']['size_mb'],
                0,
                file['source']['modified_formatted'],
                file['destination']['modified_formatted'],
                file['source']['version_count'],
                file['destination']['version_count'],
                0,
                file['source'].get('last_modified_by', ''),
                file['destination'].get('last_modified_by', ''),
                file['source']['path'],
                file['destination']['path'],
                ''
            ])
        
        for file in differences['missing_in_destination']:
            writer.writerow([
                file['name'],
                'MISSING IN DESTINATION',
                file['size_mb'],
                '',
                '',
                file['modified_formatted'],
                '',
                file['version_count'],
                '',
                '',
                file.get('last_modified_by', ''),
                '',
                file['path'],
                '',
                'File not found in destination'
            ])
        
        for file in differences['missing_in_source']:
            writer.writerow([
                file['name'],
                'MISSING IN SOURCE',
                '',
                file['size_mb'],
                '',
                '',
                file['modified_formatted'],
                '',
                file['version_count'],
                '',
                '',
                file.get('last_modified_by', ''),
                '',
                file['path'],
                'File not found in source'
            ])
        
        for item in differences['size_mismatch']:
            writer.writerow([
                item['name'],
                'SIZE MISMATCH',
                item['source_size_mb'],
                item['dest_size_mb'],
                item['diff_mb'],
                item['source']['modified_formatted'],
                item['destination']['modified_formatted'],
                item['source']['version_count'],
                item['destination']['version_count'],
                item['source']['version_count'] - item['destination']['version_count'],
                item['source'].get('last_modified_by', ''),
                item['destination'].get('last_modified_by', ''),
                item['source']['path'],
                item['destination']['path'],
                f'Size diff: {item["diff_mb"]:.2f} MB'
            ])
        
        for item in differences['modified_date_mismatch']:
            writer.writerow([
                item['name'],
                'MODIFIED DATE MISMATCH',
                item['source']['size_mb'],
                item['destination']['size_mb'],
                item['source']['size_mb'] - item['destination']['size_mb'],
                item['source_modified'],
                item['dest_modified'],
                item['source']['version_count'],
                item['destination']['version_count'],
                item['source']['version_count'] - item['destination']['version_count'],
                item['source'].get('last_modified_by', ''),
                item['destination'].get('last_modified_by', ''),
                item['source']['path'],
                item['destination']['path'],
                'Modified dates differ'
            ])
        
        for item in differences['version_count_mismatch']:
            writer.writerow([
                item['name'],
                'VERSION COUNT MISMATCH',
                item['source']['size_mb'],
                item['destination']['size_mb'],
                item['source']['size_mb'] - item['destination']['size_mb'],
                item['source']['modified_formatted'],
                item['destination']['modified_formatted'],
                item['source_versions'],
                item['dest_versions'],
                item['source_versions'] - item['dest_versions'],
                item['source'].get('last_modified_by', ''),
                item['destination'].get('last_modified_by', ''),
                item['source']['path'],
                item['destination']['path'],
                f'Version diff: {item["source_versions"] - item["dest_versions"]}'
            ])
        
        for item in differences['version_editor_mismatch']:
            writer.writerow([
                item['name'],
                'VERSION EDITOR MISMATCH',
                item['source']['size_mb'],
                item['destination']['size_mb'],
                item['source']['size_mb'] - item['destination']['size_mb'],
                item['source']['modified_formatted'],
                item['destination']['modified_formatted'],
                item['source_version_count'],
                item['dest_version_count'],
                item['source_version_count'] - item['dest_version_count'],
                ', '.join(item['source_editors']),
                ', '.join(item['dest_editors']),
                item['source']['path'],
                item['destination']['path'],
                'Version editors differ'
            ])
        
        for item in differences['file_name_mismatch']:
            writer.writerow([
                item['name'],
                'FILE NAME ISSUE',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                item.get('issue', 'File name mismatch')
            ])
    
    print(f"[FILE] Detailed report saved: {detail_file}")
    
    version_detail_file = os.path.join(CONFIG['output']['output_dir'], 
                                        CONFIG['output'].get('version_details_file', 'version_comparison_details.csv'))
    
    save_version_details(differences, source_prefix, dest_prefix, version_detail_file)
    
    return summary_file, detail_file, version_detail_file

def save_version_details(differences, source_prefix, dest_prefix, version_detail_file):
    """Save detailed version comparison including ALL information"""
    
    with open(version_detail_file, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        
        writer.writerow(['Complete Version Comparison with All Details'])
        writer.writerow(['Generated At', datetime.now().strftime('%Y-%m-%d %H:%M:%S')])
        writer.writerow(['Source', source_prefix])
        writer.writerow(['Destination', dest_prefix])
        writer.writerow([])
        
        writer.writerow([
            'File Name', 'Side', 
            'Version Label', 'Is Current',
            'Created Date', 'Modified Date',
            'Size (MB)', 
            'Editor Name', 'Editor Email',
            'Author Name', 'Author Email',
            'Check-in Comment',
            'Status'
        ])
        
        version_mismatch_files = []
        version_mismatch_files.extend(differences.get('version_count_mismatch', []))
        version_mismatch_files.extend(differences.get('version_editor_mismatch', []))
        
        if version_mismatch_files:
            for item in version_mismatch_files:
                source_file = item.get('source', {})
                dest_file = item.get('destination', {})
                
                for version in source_file.get('versions', []):
                    writer.writerow([
                        source_file.get('name', ''),
                        'SOURCE',
                        version.get('version_label', ''),
                        'Yes' if version.get('is_current') else 'No',
                        version.get('created_formatted', ''),
                        version.get('modified_formatted', ''),
                        version.get('size_mb', 0),
                        version.get('editor_name', ''),
                        version.get('editor_email', ''),
                        version.get('author_name', ''),
                        version.get('author_email', ''),
                        version.get('check_in_comment', ''),
                        'Source'
                    ])
                
                for version in dest_file.get('versions', []):
                    writer.writerow([
                        dest_file.get('name', ''),
                        'DESTINATION',
                        version.get('version_label', ''),
                        'Yes' if version.get('is_current') else 'No',
                        version.get('created_formatted', ''),
                        version.get('modified_formatted', ''),
                        version.get('size_mb', 0),
                        version.get('editor_name', ''),
                        version.get('editor_email', ''),
                        version.get('author_name', ''),
                        version.get('author_email', ''),
                        version.get('check_in_comment', ''),
                        'Destination'
                    ])
                
                writer.writerow([])
        else:
            writer.writerow(['No version mismatches found'])
    
    print(f"[FILE] Version details report saved: {version_detail_file}")

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    """Main function to compare two libraries"""
    print_header("MIGRATION COMPARISON TOOL (with Complete Version Info)")
    print("Compare two document libraries after migration")
    
    if SHOW_TRACEBACK:
        print("[WARNING] TRACEBACK MODE: ENABLED (detailed error output)")
    else:
        print("[INFO] TRACEBACK MODE: DISABLED (set SHOW_TRACEBACK=True for debugging)")
    
    print(f"[INFO] Batch Size: {BATCH_SIZE}, Max Workers: {MAX_WORKERS}")
    print(f"[INFO] File size will be retrieved from versions endpoint")
    
    source_site = CONFIG['source']['site_url']
    dest_site = CONFIG['destination']['site_url']
    source_library = CONFIG['source']['library_name']
    dest_library = CONFIG['destination']['library_name']
    
    print(f"\n[INFO] Source Site: {source_site}")
    print(f"[INFO] Source Library: {source_library}")
    print(f"[INFO] Destination Site: {dest_site}")
    print(f"[INFO] Destination Library: {dest_library}")
    print("="*80)
    
    print("\n[SEARCH] Finding libraries...")
    source_library_id = get_library_id(source_site, source_library)
    if not source_library_id:
        print_error("Source library not found!")
        return
    
    dest_library_id = get_library_id(dest_site, dest_library)
    if not dest_library_id:
        print_error("Destination library not found!")
        return
    
    source_items = get_all_items_from_library(source_site, source_library_id, source_library)
    dest_items = get_all_items_from_library(dest_site, dest_library_id, dest_library)
    
    if not source_items or not dest_items:
        print_error("No items found in one or both libraries!")
        return
    
    print_header("BUILDING FILE MAPS (with complete version details)")
    
    source_map, source_size, source_versions, source_versions_files = build_file_map(
        source_items, source_site, source_library_id, source_library
    )
    
    dest_map, dest_size, dest_versions, dest_versions_files = build_file_map(
        dest_items, dest_site, dest_library_id, dest_library
    )
    
    print_header("SOURCE LIBRARY SUMMARY")
    print(f"  Total Files: {len([f for files in source_map.values() for f in files])}")
    print(f"  Total Size: {bytes_to_gb(source_size):.2f} GB")
    print(f"  Total Versions: {source_versions}")
    print(f"  Files with Versions: {source_versions_files}")
    
    print_header("DESTINATION LIBRARY SUMMARY")
    print(f"  Total Files: {len([f for files in dest_map.values() for f in files])}")
    print(f"  Total Size: {bytes_to_gb(dest_size):.2f} GB")
    print(f"  Total Versions: {dest_versions}")
    print(f"  Files with Versions: {dest_versions_files}")
    
    differences = compare_libraries(source_map, dest_map, source_site, dest_site)
    
    summary_file, detail_file, version_detail_file = save_comparison_report(differences, source_site, dest_site)
    
    print_header("COMPARISON COMPLETED SUCCESSFULLY!")
    print(f"\n[FILE] Reports saved in: {CONFIG['output']['output_dir']}")
    print(f"  - Summary: {os.path.basename(summary_file)}")
    print(f"  - Details: {os.path.basename(detail_file)}")
    print(f"  - Version Details: {os.path.basename(version_detail_file)}")
    
    print_progress_summary()

if __name__ == "__main__":
    main()
