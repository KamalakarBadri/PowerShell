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
# CONFIGURATION - UPDATE THESE VALUES
# ============================================================

SHAREPOINT_SITE = "https://test.sharepoint.com/sites/New365"
OUTPUT_CSV = f"SharePointContent_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"

# Authentication Configuration
TENANT_NAME = "test.onmicrosoft.com"
APP_ID = "73efa35d-6188-42d4-b258-838a977eb149"
SCOPE_SHAREPOINT = "https://test.sharepoint.com/.default"

# Certificate paths
CERTIFICATE_PATH = "certificate.pem"
PRIVATE_KEY_PATH = "private_key.pem"

# ============================================================
# LIBRARY SELECTION - Choose one option
# ============================================================

# Option 1: Process ALL libraries (set to True)
PROCESS_ALL_LIBRARIES = True

# Option 2: Process specific libraries (list their names)
# Set PROCESS_ALL_LIBRARIES = False and add library names below
SPECIFIC_LIBRARIES = [
    "Documents",
    "Reports",
    "Templates"
]

# ============================================================
# OPTIONAL: Skip specific libraries (if processing all)
# ============================================================

SKIP_LIBRARIES = [
    "Site Assets",
    "Site Pages",
    "Style Library",
    "Form Templates"
]

# ============================================================
# AUTHENTICATION FUNCTIONS
# ============================================================

def load_certificate_and_key():
    """Load certificate and private key from PEM files"""
    try:
        with open(CERTIFICATE_PATH, "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())

        with open(PRIVATE_KEY_PATH, "rb") as key_file:
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
            "aud": f"https://login.microsoftonline.com/{TENANT_NAME}/oauth2/v2.0/token",
            "exp": expiration,
            "iss": APP_ID,
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": APP_ID
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
    url = f"https://login.microsoftonline.com/{TENANT_NAME}/oauth2/v2.0/token"
    
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    
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
    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error: {err}")
        print(f"Response: {response.text}")
        raise
    except Exception as err:
        print(f"Error: {err}")
        raise

def authenticate_sharepoint():
    """Authenticate and get SharePoint access token using certificate"""
    try:
        print("Loading certificate and private key...")
        certificate, private_key = load_certificate_and_key()
        
        print("Generating JWT token...")
        jwt = get_jwt_token(certificate, private_key)
        
        print("Getting access token...")
        access_token = get_access_token(jwt)
        
        print("Successfully authenticated to SharePoint")
        return access_token
    except Exception as e:
        print(f"Authentication failed: {str(e)}")
        return None

def get_access_token_with_retry():
    """Get access token with retry logic"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            token = authenticate_sharepoint()
            if token:
                return token
            print(f"Authentication attempt {attempt + 1} failed. Retrying...")
            time.sleep(2)
        except Exception as e:
            print(f"Error during authentication: {str(e)}")
            if attempt < max_retries - 1:
                time.sleep(5)
    return None

def make_sharepoint_request(url, access_token, method='GET', headers=None, max_retries=3):
    """Make a request to SharePoint REST API with retry logic"""
    default_headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json"
    }
    
    if headers:
        default_headers.update(headers)
    
    for attempt in range(max_retries):
        try:
            response = requests.request(method, url, headers=default_headers, timeout=60)
            
            if response.status_code == 429:
                wait_time = 5 * (attempt + 1)
                print(f"  ⚠️ 429: Rate limited. Waiting {wait_time}s...")
                time.sleep(wait_time)
                continue
            
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.HTTPError as e:
            if response.status_code == 429:
                continue
            if response.status_code == 401 and attempt < max_retries - 1:
                print(f"  ⚠️ Token expired, refreshing...")
                access_token = authenticate_sharepoint()
                if access_token:
                    default_headers["Authorization"] = f"Bearer {access_token}"
                    continue
            print(f"Request failed: {str(e)}")
            return None
        except requests.exceptions.RequestException as e:
            print(f"Request error: {str(e)}")
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)
                continue
            return None
    
    return None

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def safe_int_conversion(value):
    """Safely convert value to integer"""
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
    """Convert bytes to MB"""
    bytes_value = safe_int_conversion(bytes_value)
    if bytes_value == 0:
        return 0.00
    return round(bytes_value / (1024 * 1024), 2)

# ============================================================
# SHAREPOINT DATA RETRIEVAL
# ============================================================

def get_all_lists(site_url, access_token):
    """Get all document libraries from SharePoint site"""
    lists_url = urljoin(site_url + "/", "_api/web/lists")
    response = make_sharepoint_request(lists_url, access_token)
    
    if response and 'd' in response and 'results' in response['d']:
        return [lst for lst in response['d']['results'] if lst['BaseTemplate'] == 101]
    return []

def filter_libraries(libraries):
    """Filter libraries based on configuration"""
    if PROCESS_ALL_LIBRARIES:
        filtered = [lib for lib in libraries if lib['Title'] not in SKIP_LIBRARIES]
        skipped = len(libraries) - len(filtered)
        if skipped > 0:
            print(f"  ⏭️ Skipped {skipped} libraries (configured in SKIP_LIBRARIES)")
        return filtered
    else:
        filtered = [lib for lib in libraries if lib['Title'] in SPECIFIC_LIBRARIES]
        not_found = [name for name in SPECIFIC_LIBRARIES if name not in [lib['Title'] for lib in libraries]]
        if not_found:
            print(f"  ⚠️ Libraries not found: {', '.join(not_found)}")
        return filtered

def get_list_items(site_url, list_id, access_token):
    """Get all items from a list with pagination"""
    items_url = f"{site_url}/_api/web/lists(guid'{list_id}')/items?$expand=File,Folder"
    all_items = []
    next_url = items_url
    
    while next_url:
        response = make_sharepoint_request(next_url, access_token)
        
        if not response:
            break
            
        if 'd' in response and 'results' in response['d']:
            all_items.extend(response['d']['results'])
            
        next_url = None
        if '__next' in response.get('d', {}):
            next_url = response['d']['__next']
            
    return all_items

def get_file_details_from_versions(site_url, list_id, item_id, access_token):
    """Get file details from versions endpoint"""
    try:
        versions_url = f"{site_url}/_api/Web/Lists(guid'{list_id}')/items({item_id})/versions?$top=1&$orderby=Created desc"
        
        response = make_sharepoint_request(versions_url, access_token)
        
        if not response or 'd' not in response:
            return None
        
        if 'results' not in response['d'] or len(response['d']['results']) == 0:
            return None
        
        latest_version = response['d']['results'][0]
        
        details = {
            'created_by': 'N/A',
            'modified_by': 'N/A',
            'file_size': 0,
            'file_size_mb': 0.00
        }
        
        if 'Author' in latest_version and latest_version['Author']:
            details['created_by'] = latest_version['Author'].get('LookupValue', 'N/A')
        
        if 'Editor' in latest_version and latest_version['Editor']:
            details['modified_by'] = latest_version['Editor'].get('LookupValue', 'N/A')
        
        file_size = latest_version.get('File_x005f_x0020_x005f_Size', 0)
        if not file_size:
            file_size = latest_version.get('SMTotalFileStreamSize', 0)
        
        details['file_size'] = safe_int_conversion(file_size)
        details['file_size_mb'] = bytes_to_mb(details['file_size'])
        
        return details
        
    except Exception as e:
        return None

def get_file_versions_count(site_url, list_id, item_id, access_token):
    """Get version count for a specific file item"""
    try:
        versions_url = f"{site_url}/_api/Web/Lists(guid'{list_id}')/items({item_id})/versions"
        
        response = make_sharepoint_request(versions_url, access_token)
        
        if not response or 'd' not in response:
            return 0
        
        if 'results' not in response['d']:
            return 0
        
        return len(response['d']['results'])
        
    except Exception as e:
        return 0

def process_item(site_url, list_id, item, access_token, item_index, total_items):
    """Extract relevant details from an item using versions endpoint"""
    item_type = item.get('FileSystemObjectType', 0)
    item_id = item.get('Id', 0)
    
    details = {
        'Type': 'File' if item_type == 0 else 'Folder',
        'ID': item_id,
        'Name': '',
        'Path': '',
        'Size': 0,
        'Size_MB': 0.00,
        'Created': item.get('Created', 'N/A'),
        'Modified': item.get('Modified', 'N/A'),
        'Author': 'N/A',
        'Editor': 'N/A',
        'Version_Count': 0
    }
    
    if item_type == 0:  # File
        if 'File' in item and item['File']:
            file = item['File']
            details['Name'] = file.get('Name', '')
            details['Path'] = file.get('ServerRelativeUrl', '')
        else:
            details['Name'] = item.get('FileLeafRef', '')
            details['Path'] = item.get('FileRef', '')
        
        version_details = get_file_details_from_versions(site_url, list_id, item_id, access_token)
        
        if version_details:
            details['Author'] = version_details.get('created_by', 'N/A')
            details['Editor'] = version_details.get('modified_by', 'N/A')
            details['Size'] = version_details.get('file_size', 0)
            details['Size_MB'] = version_details.get('file_size_mb', 0.00)
        
        details['Version_Count'] = get_file_versions_count(site_url, list_id, item_id, access_token)
        
    else:  # Folder
        if 'Folder' in item and item['Folder']:
            folder = item['Folder']
            details['Name'] = folder.get('Name', '')
            details['Path'] = folder.get('ServerRelativeUrl', '')
        else:
            details['Name'] = item.get('FileLeafRef', '')
            details['Path'] = item.get('FileRef', '')
    
    return details

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    print("="*80)
    print("📊 SHAREPOINT LIBRARY INVENTORY REPORT")
    print("(Uses Versions endpoint for Created By, Modified By, File Size)")
    print("="*80)
    
    print(f"\n📍 Site: {SHAREPOINT_SITE}")
    print(f"📁 Output: {OUTPUT_CSV}")
    
    if PROCESS_ALL_LIBRARIES:
        print(f"📚 Mode: ALL libraries (excluding {len(SKIP_LIBRARIES)} skipped)")
        if SKIP_LIBRARIES:
            print(f"   Skipped: {', '.join(SKIP_LIBRARIES)}")
    else:
        print(f"📚 Mode: SPECIFIC libraries ({len(SPECIFIC_LIBRARIES)} selected)")
        print(f"   Libraries: {', '.join(SPECIFIC_LIBRARIES)}")
    print("="*80)
    
    print("\n🔐 Authenticating to SharePoint...")
    access_token = get_access_token_with_retry()
    
    if not access_token:
        print("❌ Failed to authenticate to SharePoint. Exiting.")
        return
    
    print("✅ Authentication successful\n")
    
    print("📁 Retrieving document libraries...")
    all_libraries = get_all_lists(SHAREPOINT_SITE, access_token)
    
    if not all_libraries:
        print("❌ No document libraries found.")
        return
    
    print(f"\n📚 Found {len(all_libraries)} document libraries total")
    
    libraries = filter_libraries(all_libraries)
    
    if not libraries:
        print("❌ No libraries match your selection criteria.")
        return
    
    print(f"📚 Processing {len(libraries)} libraries:")
    for lib in libraries:
        print(f"  - {lib['Title']}")
    
    print(f"\n📊 Generating report...")
    print("="*80)
    
    with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = [
            'Library', 'Type', 'ID', 'Name', 'Path', 
            'Size (Bytes)', 'Size (MB)', 
            'Created', 'Modified',
            'Author (Created By)', 'Editor (Modified By)', 'Version Count'
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        
        total_files = 0
        total_folders = 0
        total_versions = 0
        processed_libraries = 0
        grand_total_items = 0
        
        for library in libraries:
            library_title = library['Title']
            library_id = library['Id']
            
            processed_libraries += 1
            print(f"\n📁 [{processed_libraries}/{len(libraries)}] Processing library: {library_title}")
            print("-"*60)
            
            items = get_list_items(SHAREPOINT_SITE, library_id, access_token)
            
            if not items:
                print("  No items found in this library")
                continue
            
            total_items_in_library = len(items)
            print(f"  Total items in library: {total_items_in_library}")
            print(f"  Processing items...")
            
            library_files = 0
            library_folders = 0
            library_versions = 0
            processed_count = 0
            
            for item in items:
                processed_count += 1
                grand_total_items += 1
                
                # ✅ Show progress for every item
                progress_percent = (processed_count / total_items_in_library) * 100
                print(f"\r  [{processed_count}/{total_items_in_library}] ({progress_percent:.1f}%) Processing item ID: {item.get('Id', 'N/A')}...", end="")
                
                try:
                    details = process_item(SHAREPOINT_SITE, library_id, item, access_token, processed_count, total_items_in_library)
                    details['Library'] = library_title
                    
                    row = {
                        'Library': details['Library'],
                        'Type': details['Type'],
                        'ID': details['ID'],
                        'Name': details['Name'],
                        'Path': details['Path'],
                        'Size (Bytes)': details['Size'],
                        'Size (MB)': f"{details['Size_MB']:.2f}",
                        'Created': details['Created'],
                        'Modified': details['Modified'],
                        'Author (Created By)': details['Author'],
                        'Editor (Modified By)': details['Editor'],
                        'Version Count': details['Version_Count']
                    }
                    writer.writerow(row)
                    
                    if details['Type'] == 'File':
                        library_files += 1
                        library_versions += details['Version_Count']
                    else:
                        library_folders += 1
                    
                except Exception as e:
                    print(f"\n  ❌ Error processing item {item.get('Id', 'unknown')}: {str(e)}")
            
            total_files += library_files
            total_folders += library_folders
            total_versions += library_versions
            
            # ✅ Show library completion summary
            print(f"\n  ✅ Library complete: {library_files} files, {library_folders} folders, {library_versions} versions")
            print(f"  📊 Grand Total Items Processed: {grand_total_items}")
    
    # ✅ Final summary
    print("\n" + "="*80)
    print("📊 REPORT GENERATED SUCCESSFULLY")
    print("="*80)
    print(f"\n📄 Summary:")
    print(f"  Libraries Processed: {processed_libraries}")
    print(f"  Total Files: {total_files:,}")
    print(f"  Total Folders: {total_folders:,}")
    print(f"  Total Items: {total_files + total_folders:,}")
    print(f"  Total Versions: {total_versions:,}")
    print(f"\n📁 Report saved: {OUTPUT_CSV}")
    print("="*80)

if __name__ == "__main__":
    main()
