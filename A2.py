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
    # SharePoint Site Configuration
    "site_url": "https://geekbyteonline.sharepoint.com/sites/Team_Site2",
    
    # Authentication Configuration
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "scope": "https://geekbyteonline.sharepoint.com/.default",
    
    # Certificate Paths
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    
    # Output Configuration
    "output_csv": f"File_Version_History_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"
}

# Token cache with expiration
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
            raise Exception(f"Certificate files not found. Check paths: {CONFIG['certificate_path']}, {CONFIG['private_key_path']}")
        
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
        expiration = now + 300  # 5 minutes
        
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
        print(f"Response: {response.text}")
        raise
    except Exception as err:
        print(f"Error: {err}")
        raise

def get_cached_token(force_refresh=False):
    """Get cached token if it's still valid, otherwise get a new one"""
    cache = TOKEN_CACHE
    
    # If token exists and hasn't expired (with 5 minute buffer), and not forcing refresh
    if not force_refresh and cache["token"] and cache["expires"] > time.time() + 300:
        return cache["token"]
    
    try:
        print(f"\n  🔄 {'Refreshing' if force_refresh else 'Getting'} access token...")
        certificate, private_key = load_certificate_and_key()
        jwt = get_jwt_token(certificate, private_key)
        token = get_access_token(jwt, CONFIG['scope'])
        
        if token:
            # Cache the token with expiration (assuming 1 hour lifetime)
            cache["token"] = token
            cache["expires"] = time.time() + 3600
            print(f"  ✓ Token obtained, expires at: {datetime.fromtimestamp(cache['expires']).strftime('%Y-%m-%d %H:%M:%S')}")
            return token
        else:
            print("  ✗ Failed to get token")
            return None
    except Exception as e:
        print(f"  ✗ Authentication failed: {str(e)}")
        return None

def make_sharepoint_request(url, access_token, max_retries=2):
    """Make a request to SharePoint REST API with token refresh on 401"""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json"
    }
    
    for attempt in range(max_retries + 1):
        try:
            response = requests.get(url, headers=headers)
            
            # If unauthorized (401), refresh token and retry
            if response.status_code == 401 and attempt < max_retries:
                print(f"  ⚠️ Token expired (401), refreshing...")
                new_token = get_cached_token(force_refresh=True)
                if new_token:
                    headers["Authorization"] = f"Bearer {new_token}"
                    access_token = new_token  # Update for next attempts
                    continue
                else:
                    print(f"  ✗ Failed to refresh token")
                    return None
            
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 401 and attempt < max_retries:
                print(f"  ⚠️ Token expired (401), refreshing...")
                new_token = get_cached_token(force_refresh=True)
                if new_token:
                    headers["Authorization"] = f"Bearer {new_token}"
                    access_token = new_token
                    continue
            print(f"  Request failed (attempt {attempt + 1}): {str(e)}")
            return None
        except requests.exceptions.RequestException as e:
            print(f"  Request failed (attempt {attempt + 1}): {str(e)}")
            if attempt < max_retries:
                time.sleep(2)  # Wait before retry
                continue
            return None
    
    return None

def get_site_url():
    """Get the site URL from CONFIG"""
    return CONFIG['site_url'].rstrip('/')

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

# ============================================================
# DYNAMIC CSV REPORT FUNCTIONS
# ============================================================

def initialize_reports(output_file):
    """Initialize CSV files with headers"""
    global csv_writers, csv_files
    
    # Main Report
    main_file = output_file
    main_fieldnames = [
        'Library',
        'Item ID',
        'File Name',
        'File Path',
        'Current File Size (MB)',
        'Version Count',
        'First Version Date',
        'Last Version Date',
        'Total Versions Size (MB)',
        'File Created',
        'File Modified',
        'Processed At'
    ]
    
    csv_files['main'] = open(main_file, 'w', newline='', encoding='utf-8-sig')
    csv_writers['main'] = csv.DictWriter(csv_files['main'], fieldnames=main_fieldnames)
    csv_writers['main'].writeheader()
    csv_files['main'].flush()
    print(f"✓ Main report initialized: {main_file}")
    
    # Detailed Report
    detailed_file = output_file.replace('.csv', '_Detailed_Versions.csv')
    detailed_fieldnames = [
        'Library',
        'Item ID',
        'File Name',
        'File Path',
        'Version ID',
        'Version Label',
        'UI Version',
        'Version Created',
        'Is Current Version',
        'Version Size (MB)',
        'Check-in Comment',
        'Author',
        'Editor',
        'Processed At'
    ]
    
    csv_files['detailed'] = open(detailed_file, 'w', newline='', encoding='utf-8-sig')
    csv_writers['detailed'] = csv.DictWriter(csv_files['detailed'], fieldnames=detailed_fieldnames)
    csv_writers['detailed'].writeheader()
    csv_files['detailed'].flush()
    print(f"✓ Detailed report initialized: {detailed_file}")
    
    # Summary Report
    summary_file = output_file.replace('.csv', '_Summary.csv')
    summary_fieldnames = [
        'Library',
        'Total Files',
        'Total Versions',
        'Total Current Size (MB)',
        'Total Versions Size (MB)',
        'Average Versions per File',
        'Last Updated'
    ]
    
    csv_files['summary'] = open(summary_file, 'w', newline='', encoding='utf-8-sig')
    csv_writers['summary'] = csv.DictWriter(csv_files['summary'], fieldnames=summary_fieldnames)
    csv_writers['summary'].writeheader()
    csv_files['summary'].flush()
    print(f"✓ Summary report initialized: {summary_file}")

def append_to_main_report(data):
    """Append a row to the main report"""
    global csv_writers, csv_files
    
    try:
        row = {
            'Library': data.get('library', ''),
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
            'Processed At': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }
        csv_writers['main'].writerow(row)
        csv_files['main'].flush()
        return True
    except Exception as e:
        print(f"Error appending to main report: {str(e)}")
        return False

def append_to_detailed_report(data):
    """Append version rows to the detailed report"""
    global csv_writers, csv_files
    
    try:
        if data.get('versions'):
            for version in data['versions']:
                row = {
                    'Library': data.get('library', ''),
                    'Item ID': data.get('item_id', 0),
                    'File Name': data.get('file_name', ''),
                    'File Path': data.get('file_path', ''),
                    'Version ID': version.get('version_id', 0),
                    'Version Label': version.get('version_label', ''),
                    'UI Version': version.get('ui_version_string', ''),
                    'Version Created': format_datetime(version.get('created', 'N/A')),
                    'Is Current Version': 'Yes' if version.get('is_current', False) else 'No',
                    'Version Size (MB)': f"{bytes_to_mb(version.get('size', 0)):.2f}",
                    'Check-in Comment': version.get('checkin_comment', ''),
                    'Author': version.get('author', ''),
                    'Editor': version.get('editor', ''),
                    'Processed At': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                csv_writers['detailed'].writerow(row)
        csv_files['detailed'].flush()
        return True
    except Exception as e:
        print(f"Error appending to detailed report: {str(e)}")
        return False

def update_summary_report(library_summary):
    """Update the summary report with current data"""
    global csv_writers, csv_files
    
    try:
        csv_files['summary'].close()
        csv_files['summary'] = open(csv_files['summary'].name, 'w', newline='', encoding='utf-8-sig')
        csv_writers['summary'] = csv.DictWriter(csv_files['summary'], 
                                               fieldnames=['Library', 'Total Files', 'Total Versions', 
                                                          'Total Current Size (MB)', 'Total Versions Size (MB)',
                                                          'Average Versions per File', 'Last Updated'])
        csv_writers['summary'].writeheader()
        
        for lib, stats in library_summary.items():
            avg_versions = stats['versions'] / stats['files'] if stats['files'] > 0 else 0
            row = {
                'Library': lib,
                'Total Files': stats['files'],
                'Total Versions': stats['versions'],
                'Total Current Size (MB)': f"{bytes_to_mb(stats['current_size']):.2f}",
                'Total Versions Size (MB)': f"{bytes_to_mb(stats['versions_size']):.2f}",
                'Average Versions per File': round(avg_versions, 2),
                'Last Updated': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            csv_writers['summary'].writerow(row)
        
        csv_files['summary'].flush()
        return True
    except Exception as e:
        print(f"Error updating summary report: {str(e)}")
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
# SHAREPOINT DATA RETRIEVAL FUNCTIONS WITH TOKEN REFRESH
# ============================================================

def get_all_libraries(site_url, access_token):
    """Get all document libraries from SharePoint site with pagination and token refresh"""
    print("\nGetting document libraries...")
    lists_url = f"{site_url}/_api/web/lists"
    all_libraries = []
    next_url = lists_url
    
    while next_url:
        print(f"  Fetching libraries page...")
        response = make_sharepoint_request(next_url, access_token)
        
        if not response or 'd' not in response:
            break
        
        if 'results' in response['d']:
            for lst in response['d']['results']:
                if lst['BaseTemplate'] == 101:
                    all_libraries.append({
                        'id': lst['Id'],
                        'title': lst['Title']
                    })
        
        # Check for next page
        next_url = None
        if '__next' in response.get('d', {}):
            next_url = response['d']['__next']
    
    return all_libraries

def get_all_items_from_library(site_url, library_id, access_token):
    """Get all items from a library with pagination and token refresh"""
    print(f"    Fetching items from library...")
    
    items_url = f"{site_url}/_api/web/lists(guid'{library_id}')/items?$expand=File"
    all_items = []
    next_url = items_url
    page_count = 0
    
    while next_url:
        page_count += 1
        print(f"    Fetching page {page_count}...", end="")
        response = make_sharepoint_request(next_url, access_token)
        
        if not response or 'd' not in response:
            print(" ✗ Failed")
            break
        
        if 'results' in response['d']:
            items_in_page = len(response['d']['results'])
            all_items.extend(response['d']['results'])
            print(f" ✓ Got {items_in_page} items")
        
        # Check for next page
        next_url = None
        if '__next' in response.get('d', {}):
            next_url = response['d']['__next']
    
    print(f"    Total items fetched: {len(all_items)}")
    return all_items

def get_file_versions(site_url, list_id, item_id, access_token):
    """Get versions for a specific item with token refresh"""
    try:
        versions_url = f"{site_url}/_api/Web/Lists(guid'{list_id}')/items({item_id})/versions"
        
        response = make_sharepoint_request(versions_url, access_token)
        
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

def process_file_item(site_url, list_id, item, access_token, library_title, library_summary):
    """Process a single item and update reports dynamically"""
    try:
        item_id = item.get('Id')
        fsob_type = item.get('FileSystemObjectType', 1)
        
        if fsob_type == 1:
            return None
        
        file_details = get_file_details_from_item(item)
        versions = get_file_versions(site_url, list_id, item_id, access_token)
        
        version_count = len(versions)
        total_versions_size = 0
        first_version_date = 'N/A'
        last_version_date = 'N/A'
        
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
            'library_id': list_id,
            'item_id': item_id,
            'file_name': file_details['file_name'],
            'file_path': file_details['file_path'],
            'current_file_size': current_file_size,
            'current_file_size_mb': bytes_to_mb(current_file_size),
            'version_count': version_count,
            'first_version_date': first_version_date,
            'last_version_date': last_version_date,
            'total_versions_size': total_versions_size,
            'total_versions_size_mb': bytes_to_mb(total_versions_size),
            'versions': versions,
            'created_formatted': format_datetime(file_details['created']),
            'modified_formatted': format_datetime(file_details['modified']),
            'first_version_formatted': format_datetime(first_version_date),
            'last_version_formatted': format_datetime(last_version_date)
        }
        
        if library_title not in library_summary:
            library_summary[library_title] = {
                'files': 0,
                'versions': 0,
                'current_size': 0,
                'versions_size': 0
            }
        
        library_summary[library_title]['files'] += 1
        library_summary[library_title]['versions'] += version_count
        library_summary[library_title]['current_size'] += current_file_size
        library_summary[library_title]['versions_size'] += total_versions_size
        
        append_to_main_report(file_data)
        append_to_detailed_report(file_data)
        update_summary_report(library_summary)
        
        return file_data
        
    except Exception as e:
        print(f"    Error processing item: {str(e)}")
        return None

# ============================================================
# MAIN PROCESSING FUNCTIONS
# ============================================================

def process_files(site_url, access_token, output_file):
    """Process all files and update reports dynamically with token refresh"""
    
    initialize_reports(output_file)
    
    libraries = get_all_libraries(site_url, access_token)
    
    if not libraries:
        print("No document libraries found.")
        close_reports()
        return []
    
    print(f"\nFound {len(libraries)} document libraries:")
    for lib in libraries:
        print(f"  - {lib['title']}")
    
    all_file_data = []
    total_files = 0
    processed = 0
    library_summary = {}
    
    for library in libraries:
        print(f"\n{'='*60}")
        print(f"Processing library: {library['title']}")
        print(f"{'='*60}")
        
        items = get_all_items_from_library(site_url, library['id'], access_token)
        
        if not items:
            print(f"  No items found in {library['title']}")
            continue
        
        files = [item for item in items if item.get('FileSystemObjectType') == 0]
        
        if not files:
            print(f"  No files found in {library['title']}")
            continue
        
        print(f"  Found {len(files)} files in {library['title']}")
        total_files += len(files)
        
        for file_item in files:
            processed += 1
            item_id = file_item.get('Id')
            file_obj = file_item.get('File', {})
            file_name = file_obj.get('Name', f'Item_{item_id}')
            
            print(f"\n  [{processed}/{total_files}] Processing: {file_name} (ID: {item_id})", end="")
            
            file_data = process_file_item(
                site_url, 
                library['id'], 
                file_item, 
                access_token, 
                library['title'],
                library_summary
            )
            
            if file_data:
                all_file_data.append(file_data)
                print(f" ✓ ({file_data['version_count']} versions, {file_data['total_versions_size_mb']:.2f} MB total)")
                print(f"  → Reports updated dynamically")
            else:
                print(" ✗ (Failed to process)")
            
            time.sleep(0.3)
    
    print(f"\n{'='*60}")
    print(f"Processed {len(all_file_data)} files with version history.")
    print("Reports have been updated in real-time.")
    
    close_reports()
    
    return all_file_data

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    """Main function"""
    print("="*80)
    print("FILE VERSION HISTORY REPORT GENERATOR")
    print("(Real-time updates with token refresh & pagination)")
    print("="*80)
    print(f"SharePoint Site: {CONFIG['site_url']}")
    print(f"Output File: {CONFIG['output_csv']}")
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
    print("Reports will be updated after each file is processed.")
    print("Token will be automatically refreshed if it expires.")
    start_time = time.time()
    
    file_data = process_files(site_url, access_token, output_file)
    
    elapsed_time = time.time() - start_time
    print(f"\nProcessing completed in {elapsed_time:.2f} seconds.")
    
    if not file_data:
        print("\nNo files with version history were found.")
        return
    
    print("\n" + "="*80)
    print("PROCESSING COMPLETED SUCCESSFULLY!")
    print("="*80)
    print(f"✓ Main Report: {output_file}")
    print(f"✓ Detailed Report: {output_file.replace('.csv', '_Detailed_Versions.csv')}")
    print(f"✓ Summary Report: {output_file.replace('.csv', '_Summary.csv')}")
    print("="*80)
    print(f"Total files processed: {len(file_data)}")
    print(f"Total versions found: {sum(f['version_count'] for f in file_data)}")
    print("="*80)

if __name__ == "__main__":
    main()
