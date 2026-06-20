import requests
import json
import csv
import uuid
import base64
import time
import os
from datetime import datetime
from urllib.parse import urljoin
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
    "site_url": "https://geekbyteonline.sharepoint.com/sites/Team_Site2",  # Your SharePoint site URL
    
    # Authentication Configuration
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",  # Your tenant ID
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",      # Your app/client ID
    "scope": "https://geekbyteonline.sharepoint.com/.default",  # Your SharePoint scope
    
    # Certificate Paths
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    
    # Output Configuration
    "output_csv": f"File_Version_History_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"
}

# Token cache
TOKEN_CACHE = {
    "token": None,
    "expires": 0
}

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
        
        # Get certificate thumbprint (x5t)
        thumbprint = certificate.fingerprint(hashes.SHA1())
        x5t = base64.urlsafe_b64encode(thumbprint).decode('utf-8').replace('=', '')
        
        # Create JWT header
        jwt_header = {
            "alg": "RS256",
            "typ": "JWT",
            "x5t": x5t
        }
        
        # Create JWT payload
        jwt_payload = {
            "aud": f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            "exp": expiration,
            "iss": CONFIG['app_id'],
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": CONFIG['app_id']
        }
        
        # Encode header and payload
        encoded_header = base64.urlsafe_b64encode(
            json.dumps(jwt_header, separators=(',', ':')).encode('utf-8')
        ).decode('utf-8').replace('=', '')
        
        encoded_payload = base64.urlsafe_b64encode(
            json.dumps(jwt_payload, separators=(',', ':')).encode('utf-8')
        ).decode('utf-8').replace('=', '')
        
        # Combine header and payload
        jwt_unsigned = f"{encoded_header}.{encoded_payload}"
        
        # Sign the JWT
        signature = private_key.sign(
            jwt_unsigned.encode('utf-8'),
            padding.PKCS1v15(),
            hashes.SHA256()
        )
        encoded_signature = base64.urlsafe_b64encode(signature).decode('utf-8').replace('=', '')
        
        # Combine to create final JWT
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

def get_cached_token():
    """Get cached token if it's still valid, otherwise get a new one"""
    cache = TOKEN_CACHE
    
    # If token exists and hasn't expired (with 5 minute buffer)
    if cache["token"] and cache["expires"] > time.time() + 300:
        return cache["token"]
    
    try:
        print("Loading certificate and private key...")
        certificate, private_key = load_certificate_and_key()
        
        print("Generating JWT token...")
        jwt = get_jwt_token(certificate, private_key)
        
        print("Getting access token...")
        token = get_access_token(jwt, CONFIG['scope'])
        
        if token:
            # Cache the token with expiration (assuming 1 hour lifetime)
            cache["token"] = token
            cache["expires"] = time.time() + 3600
            print("✓ Authentication successful")
            return token
        else:
            print("✗ Authentication failed - no token received")
            return None
    except Exception as e:
        print(f"✗ Authentication failed: {str(e)}")
        return None

def make_sharepoint_request(url, access_token):
    """Make a request to SharePoint REST API"""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Request failed for {url}: {str(e)}")
        if hasattr(e, 'response') and e.response:
            try:
                error_data = e.response.json()
                print(f"Error details: {json.dumps(error_data, indent=2)}")
            except:
                pass
        return None

def get_site_url():
    """Get the site URL from CONFIG"""
    return CONFIG['site_url'].rstrip('/')

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def safe_int_conversion(value):
    """Safely convert a value to integer, handling strings and None values"""
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        # Remove any non-numeric characters (like commas, spaces)
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
        # Handle the format from SharePoint
        # Example: "2026-06-20T20:22:33"
        if 'T' in datetime_str:
            dt = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S")
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        return datetime_str
    except Exception as e:
        return datetime_str

# ============================================================
# VERSION RETRIEVAL USING ITEM ID - EXACT ENDPOINT
# ============================================================

def get_file_versions_by_item_id(site_url, list_id, item_id, access_token):
    """
    Get file versions using Item ID
    EXACT ENDPOINT: https://geekbyteonline.sharepoint.com/sites/Team_Site2/_api/Web/Lists(guid'e6a3eb36-59ce-44e2-a7b9-77dd61e2b67b')/items(3)/versions
    """
    try:
        # Build the exact URL as tested
        versions_url = f"{site_url}/_api/Web/Lists(guid'{list_id}')/items({item_id})/versions"
        
        print(f"\n        Fetching versions from: {versions_url}")
        response = make_sharepoint_request(versions_url, access_token)
        
        if not response:
            print(f"        No response from versions API")
            return []
        
        if 'd' not in response:
            print(f"        No 'd' in response")
            return []
        
        if 'results' not in response['d']:
            print(f"        No 'results' in response['d']")
            return []
        
        versions = []
        for version in response['d']['results']:
            # Extract version information - using exact field names from your sample
            version_data = {
                'version_id': version.get('VersionId', 0),
                'version_label': version.get('VersionLabel', ''),
                'ui_version_string': version.get('OData__x005f_UIVersionString', ''),
                'ui_version': version.get('OData__x005f_UIVersion', 0),
                'created': version.get('Created', ''),
                'is_current': version.get('IsCurrentVersion', False),
                'size': safe_int_conversion(version.get('File_x005f_x0020_x005f_Size', '0')),
                'checkin_comment': version.get('OData__x005f_CheckinComment', ''),
                'file_ref': version.get('FileRef', ''),
                'file_leaf_ref': version.get('FileLeafRef', ''),
                'author': version.get('Author', {}).get('LookupValue', '') if version.get('Author') else '',
                'editor': version.get('Editor', {}).get('LookupValue', '') if version.get('Editor') else '',
                'modified': version.get('Modified', '')
            }
            versions.append(version_data)
        
        print(f"        Found {len(versions)} versions")
        return versions
        
    except Exception as e:
        print(f"Error getting versions for item {item_id}: {str(e)}")
        return []

# ============================================================
# SHAREPOINT DATA RETRIEVAL FUNCTIONS
# ============================================================

def get_all_libraries(site_url, access_token):
    """Get all document libraries from SharePoint site"""
    lists_url = f"{site_url}/_api/web/lists"
    response = make_sharepoint_request(lists_url, access_token)
    
    if response and 'd' in response and 'results' in response['d']:
        # Filter for document libraries (BaseTemplate 101)
        libraries = []
        for lst in response['d']['results']:
            if lst['BaseTemplate'] == 101:
                libraries.append({
                    'id': lst['Id'],
                    'title': lst['Title'],
                    'url': lst.get('DefaultViewUrl', '')
                })
        return libraries
    return []

def get_all_files_in_library(site_url, library_id, access_token):
    """Get all files from a document library with pagination"""
    # Use FSObjType filter to get only files
    items_url = f"{site_url}/_api/web/lists(guid'{library_id}')/items?$filter=FSObjType eq 0&$select=Id,Title,FSObjType,Created,Modified,FileLeafRef,FileRef,File_x005f_x0020_x005f_Size"
    
    print(f"\n    Fetching files from library...")
    response = make_sharepoint_request(items_url, access_token)
    
    if not response or 'd' not in response:
        return []
    
    if 'results' not in response['d']:
        return []
    
    return response['d']['results']

def get_file_details_by_item_id(site_url, list_id, item_id, access_token):
    """Get detailed information about a file using its Item ID"""
    try:
        # Get the item details
        item_url = f"{site_url}/_api/Web/Lists(guid'{list_id}')/items({item_id})"
        item_response = make_sharepoint_request(item_url, access_token)
        
        if not item_response or 'd' not in item_response:
            return None
        
        item_data = item_response['d']
        
        # Get basic item info
        title = item_data.get('Title', '')
        created = item_data.get('Created', 'N/A')
        modified = item_data.get('Modified', 'N/A')
        file_leaf_ref = item_data.get('FileLeafRef', '')
        file_ref = item_data.get('FileRef', '')
        file_size = item_data.get('File_x005f_x0020_x005f_Size', '0')
        
        file_name = title if title else file_leaf_ref
        
        # Get versions using Item ID
        versions = get_file_versions_by_item_id(site_url, list_id, item_id, access_token)
        
        # Calculate version history summary
        version_count = len(versions)
        first_version_date = 'N/A'
        last_version_date = 'N/A'
        total_versions_size = 0
        current_file_size = safe_int_conversion(file_size)
        
        if versions:
            # Sort versions by creation date
            sorted_versions = sorted(versions, key=lambda x: x.get('created', ''))
            first_version = sorted_versions[0]
            last_version = sorted_versions[-1]
            
            first_version_date = first_version.get('created', 'N/A')
            last_version_date = last_version.get('created', 'N/A')
            
            # Calculate total size of all versions
            for version in versions:
                total_versions_size += version.get('size', 0)
        
        # If no versions, total size is just the current file size
        if version_count == 0:
            total_versions_size = current_file_size
        
        return {
            'item_id': item_id,
            'list_id': list_id,
            'file_name': file_name,
            'file_ref': file_ref,
            'file_leaf_ref': file_leaf_ref,
            'created': created,
            'modified': modified,
            'current_file_size': current_file_size,
            'current_file_size_mb': bytes_to_mb(current_file_size),
            'version_count': version_count,
            'first_version_date': first_version_date,
            'last_version_date': last_version_date,
            'total_versions_size': total_versions_size,
            'total_versions_size_mb': bytes_to_mb(total_versions_size),
            'versions': versions,
            'created_formatted': format_datetime(created),
            'modified_formatted': format_datetime(modified),
            'first_version_formatted': format_datetime(first_version_date),
            'last_version_formatted': format_datetime(last_version_date)
        }
    except Exception as e:
        print(f"Error getting file details for item {item_id}: {str(e)}")
        return None

# ============================================================
# MAIN PROCESSING FUNCTIONS
# ============================================================

def process_files(site_url, access_token):
    """Process all files in all libraries and gather version history information"""
    # Get all document libraries
    print("\nGetting document libraries...")
    libraries = get_all_libraries(site_url, access_token)
    
    if not libraries:
        print("No document libraries found.")
        return []
    
    print(f"Found {len(libraries)} document libraries:")
    for lib in libraries:
        print(f"  - {lib['title']}")
    
    # Process each library
    all_file_data = []
    total_files = 0
    processed = 0
    
    for library in libraries:
        print(f"\n{'='*60}")
        print(f"Processing library: {library['title']}")
        print(f"{'='*60}")
        
        # Get all files in the library
        files = get_all_files_in_library(site_url, library['id'], access_token)
        
        if not files:
            print(f"  No files found in {library['title']}")
            continue
        
        print(f"  Found {len(files)} files in {library['title']}")
        total_files += len(files)
        
        # Process each file
        for file_item in files:
            processed += 1
            item_id = file_item.get('Id')
            file_name = file_item.get('Title', file_item.get('FileLeafRef', 'Unknown'))
            
            print(f"\n  [{processed}/{total_files}] Processing: {file_name} (ID: {item_id})", end="")
            
            # Get file details using Item ID
            file_data = get_file_details_by_item_id(site_url, library['id'], item_id, access_token)
            
            if file_data:
                file_data['library'] = library['title']
                file_data['library_id'] = library['id']
                all_file_data.append(file_data)
                print(f" ✓ ({file_data['version_count']} versions, {file_data['total_versions_size_mb']:.2f} MB total)")
            else:
                print(" ✗ (Failed to get details)")
            
            # Add a small delay to avoid rate limiting
            time.sleep(0.5)
    
    print(f"\n{'='*60}")
    print(f"Processed {len(all_file_data)} files with version history.")
    return all_file_data

def generate_csv_report(file_data, output_file):
    """Generate CSV report from file version data"""
    try:
        fieldnames = [
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
            'File Modified'
        ]
        
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            # Sort by version count descending
            sorted_data = sorted(file_data, key=lambda x: x['version_count'], reverse=True)
            
            for data in sorted_data:
                writer.writerow({
                    'Library': data.get('library', ''),
                    'Item ID': data.get('item_id', 0),
                    'File Name': data.get('file_name', ''),
                    'File Path': data.get('file_ref', ''),
                    'Current File Size (MB)': f"{data.get('current_file_size_mb', 0.00):.2f}",
                    'Version Count': data.get('version_count', 0),
                    'First Version Date': data.get('first_version_formatted', 'N/A'),
                    'Last Version Date': data.get('last_version_formatted', 'N/A'),
                    'Total Versions Size (MB)': f"{data.get('total_versions_size_mb', 0.00):.2f}",
                    'File Created': data.get('created_formatted', 'N/A'),
                    'File Modified': data.get('modified_formatted', 'N/A')
                })
        
        print(f"\n✓ Main report generated: {output_file}")
        print(f"  Total files with version history: {len(file_data)}")
        
    except Exception as e:
        print(f"Error generating CSV report: {str(e)}")

def generate_detailed_version_report(file_data, output_file):
    """Generate a detailed CSV report with individual version information"""
    try:
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
            'Editor'
        ]
        
        # Create detailed output filename
        detailed_output = output_file.replace('.csv', '_Detailed_Versions.csv')
        
        with open(detailed_output, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=detailed_fieldnames)
            writer.writeheader()
            
            for data in file_data:
                if data.get('versions'):
                    for version in data['versions']:
                        writer.writerow({
                            'Library': data.get('library', ''),
                            'Item ID': data.get('item_id', 0),
                            'File Name': data.get('file_name', ''),
                            'File Path': data.get('file_ref', ''),
                            'Version ID': version.get('version_id', 0),
                            'Version Label': version.get('version_label', ''),
                            'UI Version': version.get('ui_version_string', ''),
                            'Version Created': format_datetime(version.get('created', 'N/A')),
                            'Is Current Version': 'Yes' if version.get('is_current', False) else 'No',
                            'Version Size (MB)': f"{bytes_to_mb(version.get('size', 0)):.2f}",
                            'Check-in Comment': version.get('checkin_comment', ''),
                            'Author': version.get('author', ''),
                            'Editor': version.get('editor', '')
                        })
        
        print(f"✓ Detailed version report generated: {detailed_output}")
        
    except Exception as e:
        print(f"Error generating detailed report: {str(e)}")

def generate_summary_report(file_data, output_file):
    """Generate a summary report with total sizes by library"""
    try:
        summary_fieldnames = [
            'Library',
            'Total Files',
            'Total Versions',
            'Total Current Size (MB)',
            'Total Versions Size (MB)',
            'Average Versions per File'
        ]
        
        # Create summary output filename
        summary_output = output_file.replace('.csv', '_Summary.csv')
        
        # Group by library
        library_summary = {}
        for data in file_data:
            lib = data.get('library', 'Unknown')
            if lib not in library_summary:
                library_summary[lib] = {
                    'files': 0,
                    'versions': 0,
                    'current_size': 0,
                    'versions_size': 0
                }
            
            library_summary[lib]['files'] += 1
            library_summary[lib]['versions'] += data.get('version_count', 0)
            library_summary[lib]['current_size'] += data.get('current_file_size', 0)
            library_summary[lib]['versions_size'] += data.get('total_versions_size', 0)
        
        with open(summary_output, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=summary_fieldnames)
            writer.writeheader()
            
            for lib, stats in library_summary.items():
                avg_versions = stats['versions'] / stats['files'] if stats['files'] > 0 else 0
                writer.writerow({
                    'Library': lib,
                    'Total Files': stats['files'],
                    'Total Versions': stats['versions'],
                    'Total Current Size (MB)': f"{bytes_to_mb(stats['current_size']):.2f}",
                    'Total Versions Size (MB)': f"{bytes_to_mb(stats['versions_size']):.2f}",
                    'Average Versions per File': round(avg_versions, 2)
                })
        
        print(f"✓ Summary report generated: {summary_output}")
        
    except Exception as e:
        print(f"Error generating summary report: {str(e)}")

def print_summary_stats(file_data):
    """Print summary statistics about the file version history"""
    if not file_data:
        print("\nNo data to summarize.")
        return
    
    total_files = len(file_data)
    total_versions = sum(f['version_count'] for f in file_data)
    avg_versions = total_versions / total_files if total_files > 0 else 0
    total_current_size = sum(f['current_file_size'] for f in file_data)
    total_versions_size = sum(f['total_versions_size'] for f in file_data)
    
    # Files with most versions
    top_files = sorted(file_data, key=lambda x: x['version_count'], reverse=True)[:10]
    
    # Files with no versions
    no_versions = [f for f in file_data if f['version_count'] == 0]
    
    print("\n" + "="*80)
    print("VERSION HISTORY SUMMARY")
    print("="*80)
    print(f"Total files analyzed: {total_files}")
    print(f"Total version records: {total_versions}")
    print(f"Average versions per file: {avg_versions:.2f}")
    print(f"Files with no versions: {len(no_versions)}")
    print(f"Total current file size: {bytes_to_mb(total_current_size):.2f} MB")
    print(f"Total versions storage size: {bytes_to_mb(total_versions_size):.2f} MB")
    print()
    
    if top_files:
        print("TOP 10 FILES WITH MOST VERSIONS:")
        print("-"*80)
        for i, file in enumerate(top_files, 1):
            print(f"{i}. {file['file_name']}")
            print(f"   Item ID: {file.get('item_id', 0)}")
            print(f"   Versions: {file['version_count']}")
            print(f"   Library: {file['library']}")
            print(f"   Current Size: {file['current_file_size_mb']:.2f} MB")
            print(f"   Total Versions Size: {file['total_versions_size_mb']:.2f} MB")
            print(f"   First Version: {file['first_version_formatted']}")
            print(f"   Last Version: {file['last_version_formatted']}")
            print()
    
    print("="*80)

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    """Main function to generate file version history report using Item ID"""
    print("="*80)
    print("FILE VERSION HISTORY REPORT GENERATOR")
    print("(Using Item ID for version retrieval)")
    print("="*80)
    print(f"SharePoint Site: {CONFIG['site_url']}")
    print(f"Output File: {CONFIG['output_csv']}")
    print("="*80)
    
    # Authenticate to SharePoint
    print("\nAuthenticating to SharePoint...")
    access_token = get_cached_token()
    
    if not access_token:
        print("✗ Authentication failed. Please check your configuration and certificate files.")
        return
    
    print("✓ Authentication successful\n")
    
    # Get the site URL
    site_url = get_site_url()
    
    # Process all files in the site
    print("Starting file processing...")
    start_time = time.time()
    
    file_data = process_files(site_url, access_token)
    
    elapsed_time = time.time() - start_time
    print(f"\nProcessing completed in {elapsed_time:.2f} seconds.")
    
    if not file_data:
        print("\nNo files with version history were found.")
        return
    
    # Generate reports
    output_file = CONFIG['output_csv']
    
    # Main report
    generate_csv_report(file_data, output_file)
    
    # Detailed version report
    generate_detailed_version_report(file_data, output_file)
    
    # Summary report
    generate_summary_report(file_data, output_file)
    
    # Print summary statistics
    print_summary_stats(file_data)
    
    print("\n" + "="*80)
    print("REPORT GENERATION COMPLETED SUCCESSFULLY!")
    print("="*80)
    print(f"✓ Main Report: {output_file}")
    print(f"✓ Detailed Version Report: {output_file.replace('.csv', '_Detailed_Versions.csv')}")
    print(f"✓ Summary Report: {output_file.replace('.csv', '_Summary.csv')}")

if __name__ == "__main__":
    main()
