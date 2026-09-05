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

# ============================================================
# CONFIGURATION - TRACEBACK ON/OFF
# ============================================================

SHOW_TRACEBACK = False  # Set to True for debugging, False for clean output

def log_error(message, error=None):
    """Log error with optional traceback"""
    print(f"  ❌ {message}")
    if SHOW_TRACEBACK and error:
        print("  " + "-"*60)
        traceback.print_exc()
        print("  " + "-"*60)

# ============================================================
# LOAD CONFIGURATION
# ============================================================

def load_config():
    """Load configuration from JSON file"""
    config_file = "compare_config.json"
    
    if not os.path.exists(config_file):
        print(f"❌ Config file '{config_file}' not found!")
        print("📌 Please create compare_config.json file")
        sys.exit(1)
    
    with open(config_file, 'r') as f:
        config = json.load(f)
    
    print(f"✅ Configuration loaded from {config_file}")
    return config

CONFIG = load_config()

# ============================================================
# GLOBAL VARIABLES
# ============================================================

TOKEN_CACHE = {}
ALLOWED_FILE_EXTENSIONS = None

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
    PROGRESS['processed_items'] += 1
    PROGRESS['current_file'] = file_name
    
    if PROGRESS['processed_items'] % 10 == 0 or PROGRESS['processed_items'] == PROGRESS['total_items']:
        elapsed = time.time() - PROGRESS['start_time']
        if PROGRESS['processed_items'] > 0:
            items_per_sec = PROGRESS['processed_items'] / elapsed if elapsed > 0 else 0
            remaining = (PROGRESS['total_items'] - PROGRESS['processed_items']) / items_per_sec if items_per_sec > 0 else 0
            
            progress_pct = (PROGRESS['processed_items'] / PROGRESS['total_items']) * 100 if PROGRESS['total_items'] > 0 else 0
            print(f"\r  📊 Progress: {PROGRESS['processed_items']}/{PROGRESS['total_items']} ({progress_pct:.1f}%) | "
                  f"Elapsed: {elapsed:.1f}s | ETA: {remaining:.1f}s | "
                  f"Current: {file_name[:30]}...", end="", flush=True)

def print_progress_summary():
    """Print final progress summary"""
    elapsed = time.time() - PROGRESS['start_time']
    print(f"\n\n  ✅ Processing complete in {elapsed:.1f} seconds")
    print(f"  📊 Processed {PROGRESS['processed_items']} items")

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
    print("  🔑 Getting access token...")
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
        print("  ✅ Token obtained")
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
                print(f"\n    ⚠️ 429: Waiting {wait_time}s...")
                time.sleep(wait_time)
                continue
            
            if response.status_code == 401 and attempt < max_retries:
                print(f"\n    ⚠️ Token expired, refreshing...")
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
        # Remove non-numeric characters except decimal point
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
            # Handle different formats
            if '.' in datetime_str and 'Z' in datetime_str:
                dt = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S.%fZ")
            elif 'Z' in datetime_str:
                dt = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%SZ")
            else:
                # Handle format without Z: "2026-06-26T05:50:41"
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
# SHAREPOINT DATA RETRIEVAL
# ============================================================

def get_library_id(site_url, library_name):
    """Get library ID by name with proper URL encoding"""
    print(f"\n🔍 Finding library: {library_name}")
    
    # URL encode the library name for the filter
    encoded_name = urllib.parse.quote(library_name)
    
    # Try exact match with proper encoding
    url = f"{site_url}/_api/web/lists?$filter=Title eq '{encoded_name}'"
    response = make_sharepoint_request(site_url, url)
    
    if response and 'd' in response and 'results' in response['d'] and len(response['d']['results']) > 0:
        library = response['d']['results'][0]
        print(f"✅ Found library: {library['Title']} (ID: {library['Id']})")
        return library['Id']
    
    # If not found, try case-insensitive search
    print(f"  ⚠️ Exact match not found, trying case-insensitive search...")
    
    # Get all document libraries
    url = f"{site_url}/_api/web/lists?$filter=BaseTemplate eq 101"
    response = make_sharepoint_request(site_url, url)
    
    if response and 'd' in response and 'results' in response['d']:
        # Case-insensitive match
        for lib in response['d']['results']:
            if lib['Title'].lower() == library_name.lower():
                print(f"✅ Found library (case-insensitive): {lib['Title']} (ID: {lib['Id']})")
                return lib['Id']
        
        # If still not found, try partial match
        print(f"  ⚠️ Case-insensitive match not found, trying partial match...")
        library_lower = library_name.lower()
        for lib in response['d']['results']:
            if library_lower in lib['Title'].lower() or lib['Title'].lower() in library_lower:
                print(f"✅ Found partial match: '{lib['Title']}' (ID: {lib['Id']})")
                print(f"  💡 Did you mean: '{lib['Title']}'?")
                return lib['Id']
        
        # List available libraries for debugging
        print(f"\n  📋 Available document libraries:")
        for lib in response['d']['results'][:15]:
            print(f"    - '{lib['Title']}'")
        if len(response['d']['results']) > 15:
            print(f"    ... and {len(response['d']['results']) - 15} more")
    
    print(f"❌ Library '{library_name}' not found")
    print(f"  💡 Please check the exact library name in SharePoint")
    return None

def get_all_items_from_library(site_url, library_id, library_name):
    """Get all items from library with pagination"""
    print(f"\n📁 Fetching items from library: {library_name}")
    
    items_url = f"{site_url}/_api/web/lists(guid'{library_id}')/items?$select=Id,Title,FileLeafRef,FileRef,File_x005f_x0020_x005f_Size,Created,Modified,FileSystemObjectType&$top=5000"
    
    all_items = []
    next_url = items_url
    page_count = 0
    
    while next_url:
        page_count += 1
        print(f"  Fetching page {page_count}...", end="")
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
    
    print(f"  Total items fetched: {len(all_items)}")
    return all_items

def get_file_versions_complete(site_url, list_id, item_id):
    """
    Get COMPLETE version information including:
    - Version ID, Label
    - Created Date (when version was created)
    - Modified Date (when version was last modified)
    - Size (file size for that version)
    - Editor (who modified)
    - Check-in comments
    - Is current version
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
            
            # Get direct fields (as seen in your response)
            created_by = version.get('Created_x005f_x0020_x005f_By', '')
            modified_by = version.get('Modified_x005f_x0020_x005f_By', '')
            
            # Parse file size (it's a string in your response)
            size_str = version.get('File_x005f_x0020_x005f_Size', '0')
            size = safe_int_conversion(size_str)
            
            version_data = {
                # Version identifiers
                'version_id': version.get('VersionId', 0),
                'version_label': version.get('VersionLabel', ''),
                'is_current': version.get('IsCurrentVersion', False),
                
                # Dates
                'created': version.get('Created', ''),
                'created_formatted': format_datetime(version.get('Created', '')),
                'modified': version.get('Modified', ''),
                'modified_formatted': format_datetime(version.get('Modified', '')),
                
                # Size
                'size': size,
                'size_mb': bytes_to_mb(size),
                
                # Editor (who modified)
                'editor_name': editor_name or modified_by,
                'editor_email': editor_email,
                'editor_id': editor_id,
                
                # Author/Creator
                'author_name': author_name or created_by,
                'author_email': author_email,
                'author_id': author_id,
                
                # Additional info
                'check_in_comment': version.get('CheckInComment', ''),
                'is_current_version': version.get('IsCurrentVersion', False)
            }
            versions.append(version_data)
        
        return versions
        
    except Exception as e:
        log_error(f"Error getting versions for item {item_id}: {str(e)}", e)
        return []

def get_file_details(item):
    """Extract file details from item"""
    file_name = item.get('FileLeafRef', '')
    if not file_name:
        file_name = item.get('Title', f"Item_{item.get('Id', 0)}")
    
    file_size = item.get('File_x005f_x0020_x005f_Size', 0)
    if not file_size or file_size == 0:
        file_size = item.get('FileSize', 0)
    
    return {
        'id': item.get('Id', 0),
        'name': file_name,
        'name_without_ext': get_file_name_without_extension(file_name),
        'extension': get_file_extension(file_name),
        'path': item.get('FileRef', ''),
        'size': safe_int_conversion(file_size),
        'size_mb': bytes_to_mb(file_size),
        'created': item.get('Created', 'N/A'),
        'modified': item.get('Modified', 'N/A'),
        'created_formatted': format_datetime(item.get('Created', 'N/A')),
        'modified_formatted': format_datetime(item.get('Modified', 'N/A')),
        'is_folder': item.get('FileSystemObjectType', 0) == 1
    }

def get_file_details_with_complete_versions(site_url, list_id, item):
    """Get file details with COMPLETE version information"""
    file_details = get_file_details(item)
    
    if not file_details['is_folder']:
        # Get ALL version details
        versions = get_file_versions_complete(site_url, list_id, item.get('Id'))
        
        # Store complete version data
        file_details['versions'] = versions
        file_details['version_count'] = len(versions)
        
        # Version statistics
        if versions:
            # Latest version (index 0 since sorted by Created desc)
            latest = versions[0]
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
            file_details['last_modified_by'] = ''
            file_details['last_modified_by_email'] = ''
            file_details['last_modified_date'] = file_details['modified_formatted']
            file_details['version_stats'] = {}
            file_details['unique_editors'] = []
            file_details['editor_count'] = 0
    else:
        # Folder handling
        file_details['version_count'] = 0
        file_details['versions'] = []
        file_details['last_modified_by'] = ''
        file_details['last_modified_by_email'] = ''
        file_details['last_modified_date'] = file_details['modified_formatted']
        file_details['version_stats'] = {}
        file_details['unique_editors'] = []
        file_details['editor_count'] = 0
    
    return file_details

# ============================================================
# COMPARISON FUNCTIONS
# ============================================================

def build_file_map(items, site_url, list_id, library_name):
    """Build a map of files by name with details"""
    file_map = {}
    total_size = 0
    total_versions = 0
    files_with_versions = 0
    folders = 0
    files_count = 0
    
    print(f"\n📊 Building file map for {library_name}...")
    
    # Set up progress tracking
    PROGRESS['total_items'] = len(items)
    PROGRESS['processed_items'] = 0
    PROGRESS['current_library'] = library_name
    PROGRESS['start_time'] = time.time()
    
    for item in items:
        file_details = get_file_details_with_complete_versions(site_url, list_id, item)
        
        if file_details['is_folder']:
            folders += 1
            update_progress(file_details['name'] or "Folder", library_name)
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
        
        update_progress(file_details['name'], library_name)
    
    # Clear progress line
    print()
    
    print(f"  Files: {files_count}")
    print(f"  Folders: {folders}")
    print(f"  Total Size: {bytes_to_gb(total_size):.2f} GB")
    print(f"  Total Versions: {total_versions}")
    print(f"  Files with versions: {files_with_versions}")
    
    return file_map, total_size, total_versions, files_with_versions

def compare_libraries(source_map, dest_map, source_site, dest_site):
    """Compare source and destination file maps"""
    print("\n" + "="*80)
    print("🔍 COMPARING SOURCE vs DESTINATION")
    print("="*80)
    
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
    
    print(f"\n📄 Files missing in Destination: {len(missing_in_dest)}")
    print(f"📄 Files missing in Source: {len(missing_in_source)}")
    
    common_keys = source_keys & dest_keys
    print(f"\n📄 Common files: {len(common_keys)}")
    
    print("\n" + "-"*80)
    print("DETAILED COMPARISON - COMMON FILES")
    print("-"*80)
    
    # Progress for comparison
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
            
            # Check extension mismatch
            if source_file['extension'] != dest_file['extension']:
                differences['file_name_mismatch'].append({
                    'name': key,
                    'source': source_file,
                    'destination': dest_file,
                    'issue': 'Extension mismatch'
                })
            
            # Check size
            if source_file['size'] != dest_file['size']:
                differences['size_mismatch'].append({
                    'name': key,
                    'source': source_file,
                    'destination': dest_file,
                    'source_size_mb': source_file['size_mb'],
                    'dest_size_mb': dest_file['size_mb'],
                    'diff_mb': source_file['size_mb'] - dest_file['size_mb']
                })
            
            # Check modified date
            if source_file['modified'] != dest_file['modified']:
                differences['modified_date_mismatch'].append({
                    'name': key,
                    'source': source_file,
                    'destination': dest_file,
                    'source_modified': source_file['modified_formatted'],
                    'dest_modified': dest_file['modified_formatted']
                })
            
            # Check version count
            if source_file['version_count'] != dest_file['version_count']:
                differences['version_count_mismatch'].append({
                    'name': key,
                    'source': source_file,
                    'destination': dest_file,
                    'source_versions': source_file['version_count'],
                    'dest_versions': dest_file['version_count']
                })
            
            # Check version editors (Modified By)
            if CONFIG['comparison_settings'].get('check_version_editor', True):
                source_editors = set()
                dest_editors = set()
                
                for v in source_file.get('versions', []):
                    if v.get('editor_name'):
                        source_editors.add(v['editor_name'])
                
                for v in dest_file.get('versions', []):
                    if v.get('editor_name'):
                        dest_editors.add(v['editor_name'])
                
                # Check if editors are different
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
            
            # If all match, it's a match
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
    
    print()  # New line after progress
    
    # Store missing files details
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
    print("\n" + "="*80)
    print("📊 COMPARISON SUMMARY")
    print("="*80)
    
    print(f"\n✅ Matched Files: {len(differences['matched_files'])}")
    
    print(f"\n❌ Missing in Destination: {len(differences['missing_in_destination'])}")
    if differences['missing_in_destination']:
        print("  (Files that exist in Source but not in Destination)")
    
    print(f"\n❌ Missing in Source: {len(differences['missing_in_source'])}")
    if differences['missing_in_source']:
        print("  (Files that exist in Destination but not in Source)")
    
    print(f"\n⚠️ Size Mismatch: {len(differences['size_mismatch'])}")
    if differences['size_mismatch']:
        print("  (Files with different sizes)")
    
    print(f"\n⚠️ Modified Date Mismatch: {len(differences['modified_date_mismatch'])}")
    if differences['modified_date_mismatch']:
        print("  (Files with different modified dates)")
    
    print(f"\n⚠️ Version Count Mismatch: {len(differences['version_count_mismatch'])}")
    if differences['version_count_mismatch']:
        print("  (Files with different version counts)")
    
    print(f"\n⚠️ Version Editor Mismatch: {len(differences['version_editor_mismatch'])}")
    if differences['version_editor_mismatch']:
        print("  (Files where version editors are different)")
    
    print(f"\n⚠️ File Name Issues: {len(differences['file_name_mismatch'])}")
    if differences['file_name_mismatch']:
        print("  (Files with extension mismatches or duplicates)")
    
    total_issues = (len(differences['missing_in_destination']) + 
                    len(differences['missing_in_source']) +
                    len(differences['size_mismatch']) +
                    len(differences['modified_date_mismatch']) +
                    len(differences['version_count_mismatch']) +
                    len(differences['version_editor_mismatch']) +
                    len(differences['file_name_mismatch']))
    
    print(f"\n🔴 TOTAL ISSUES FOUND: {total_issues}")
    print("="*80)

# ============================================================
# CSV REPORT FUNCTIONS
# ============================================================

def save_comparison_report(differences, source_site, dest_site):
    """Save comparison report to CSV"""
    os.makedirs(CONFIG['output']['output_dir'], exist_ok=True)
    
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    source_prefix = get_site_prefix(source_site)
    dest_prefix = get_site_prefix(dest_site)
    
    # Summary file
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
    
    print(f"\n📊 Summary report saved: {summary_file}")
    
    # Detailed file
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
        
        # Matched files
        for file in differences['matched_files']:
            writer.writerow([
                file['name'],
                '✅ Matched',
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
        
        # Missing in destination
        for file in differences['missing_in_destination']:
            writer.writerow([
                file['name'],
                '❌ Missing in Destination',
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
        
        # Missing in source
        for file in differences['missing_in_source']:
            writer.writerow([
                file['name'],
                '❌ Missing in Source',
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
        
        # Size mismatches
        for item in differences['size_mismatch']:
            writer.writerow([
                item['name'],
                '⚠️ Size Mismatch',
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
        
        # Modified date mismatches
        for item in differences['modified_date_mismatch']:
            writer.writerow([
                item['name'],
                '⚠️ Modified Date Mismatch',
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
        
        # Version count mismatches
        for item in differences['version_count_mismatch']:
            writer.writerow([
                item['name'],
                '⚠️ Version Count Mismatch',
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
        
        # Version editor mismatches
        for item in differences['version_editor_mismatch']:
            writer.writerow([
                item['name'],
                '⚠️ Version Editor Mismatch',
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
        
        # File name issues
        for item in differences['file_name_mismatch']:
            writer.writerow([
                item['name'],
                '⚠️ File Name Issue',
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
    
    print(f"📄 Detailed report saved: {detail_file}")
    
    # Version details file
    version_detail_file = os.path.join(CONFIG['output']['output_dir'], 
                                        CONFIG['output'].get('version_details_file', 'version_comparison_details.csv'))
    
    save_version_details(differences, source_prefix, dest_prefix, version_detail_file)
    
    return summary_file, detail_file, version_detail_file

# ============================================================
# VERSION DETAILS WITH COMPLETE INFO
# ============================================================

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
        
        # Get files with version mismatches
        version_mismatch_files = []
        version_mismatch_files.extend(differences.get('version_count_mismatch', []))
        version_mismatch_files.extend(differences.get('version_editor_mismatch', []))
        
        if version_mismatch_files:
            for item in version_mismatch_files:
                # Source versions
                source_file = item.get('source', {})
                dest_file = item.get('destination', {})
                
                # Source versions
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
                
                # Destination versions
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
                
                # Empty line between files
                writer.writerow([])
        else:
            writer.writerow(['No version mismatches found'])
    
    print(f"📄 Version details report saved: {version_detail_file}")

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    """Main function to compare two libraries"""
    print("="*80)
    print("📊 MIGRATION COMPARISON TOOL (with Complete Version Info)")
    print("Compare two document libraries after migration")
    print("="*80)
    
    if SHOW_TRACEBACK:
        print("⚠️  TRACEBACK MODE: ENABLED (detailed error output)")
    else:
        print("ℹ️  TRACEBACK MODE: DISABLED (set SHOW_TRACEBACK=True for debugging)")
    
    source_site = CONFIG['source']['site_url']
    dest_site = CONFIG['destination']['site_url']
    source_library = CONFIG['source']['library_name']
    dest_library = CONFIG['destination']['library_name']
    
    print(f"\n📌 Source Site: {source_site}")
    print(f"📌 Source Library: {source_library}")
    print(f"📌 Destination Site: {dest_site}")
    print(f"📌 Destination Library: {dest_library}")
    print("="*80)
    
    # Get library IDs
    print("\n🔍 Finding libraries...")
    source_library_id = get_library_id(source_site, source_library)
    if not source_library_id:
        print("❌ Source library not found!")
        return
    
    dest_library_id = get_library_id(dest_site, dest_library)
    if not dest_library_id:
        print("❌ Destination library not found!")
        return
    
    # Get items from both libraries
    source_items = get_all_items_from_library(source_site, source_library_id, source_library)
    dest_items = get_all_items_from_library(dest_site, dest_library_id, dest_library)
    
    if not source_items or not dest_items:
        print("❌ No items found in one or both libraries!")
        return
    
    # Build file maps
    print("\n" + "="*80)
    print("📊 BUILDING FILE MAPS (with complete version details)")
    print("="*80)
    
    source_map, source_size, source_versions, source_versions_files = build_file_map(
        source_items, source_site, source_library_id, source_library
    )
    
    dest_map, dest_size, dest_versions, dest_versions_files = build_file_map(
        dest_items, dest_site, dest_library_id, dest_library
    )
    
    # Print summary
    print("\n" + "="*80)
    print("📊 SOURCE LIBRARY SUMMARY")
    print("="*80)
    print(f"  Total Files: {len([f for files in source_map.values() for f in files])}")
    print(f"  Total Size: {bytes_to_gb(source_size):.2f} GB")
    print(f"  Total Versions: {source_versions}")
    print(f"  Files with Versions: {source_versions_files}")
    
    print("\n" + "="*80)
    print("📊 DESTINATION LIBRARY SUMMARY")
    print("="*80)
    print(f"  Total Files: {len([f for files in dest_map.values() for f in files])}")
    print(f"  Total Size: {bytes_to_gb(dest_size):.2f} GB")
    print(f"  Total Versions: {dest_versions}")
    print(f"  Files with Versions: {dest_versions_files}")
    
    # Compare libraries
    differences = compare_libraries(source_map, dest_map, source_site, dest_site)
    
    # Save reports
    summary_file, detail_file, version_detail_file = save_comparison_report(differences, source_site, dest_site)
    
    print("\n" + "="*80)
    print("✅ COMPARISON COMPLETED SUCCESSFULLY!")
    print("="*80)
    print(f"\n📁 Reports saved in: {CONFIG['output']['output_dir']}")
    print(f"  - Summary: {os.path.basename(summary_file)}")
    print(f"  - Details: {os.path.basename(detail_file)}")
    print(f"  - Version Details: {os.path.basename(version_detail_file)}")
    
    print_progress_summary()

if __name__ == "__main__":
    main()
