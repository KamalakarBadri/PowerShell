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
    "site_url": "https://test.sharepoint.com/sites/New365",  # Your SharePoint site URL
    
    # Authentication Configuration
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",  # Your tenant ID
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",      # Your app/client ID
    "scope": "https://test.sharepoint.com/.default",        # Your SharePoint scope
    
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

def make_sharepoint_request(url, access_token, method='GET', data=None, headers=None):
    """Make a request to SharePoint REST API"""
    default_headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json"
    }
    
    if headers:
        default_headers.update(headers)
    
    try:
        response = requests.request(method, url, headers=default_headers, json=data)
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

def bytes_to_mb(bytes_value):
    """Convert bytes to MB with 2 decimal places"""
    if bytes_value is None or bytes_value == 0:
        return 0.00
    return round(bytes_value / (1024 * 1024), 2)

def format_datetime(datetime_str):
    """Format datetime string to readable format"""
    if not datetime_str or datetime_str == "N/A":
        return "N/A"
    
    try:
        # Remove timezone offset if present
        datetime_str = re.sub(r'[+-]\d{2}:\d{2}$', '', datetime_str)
        
        # Add Z if missing at end
        if not datetime_str.endswith('Z'):
            datetime_str += 'Z'
        
        # Handle both formats: with and without milliseconds
        if '.' in datetime_str:
            dt = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S.%fZ")
        else:
            dt = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%SZ")
        
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception as e:
        return datetime_str  # Return original if parsing fails

# ============================================================
# SHAREPOINT DATA RETRIEVAL FUNCTIONS
# ============================================================

def get_all_libraries(site_url, access_token):
    """Get all document libraries from SharePoint site"""
    lists_url = urljoin(site_url + "/", "_api/web/lists")
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
    # Use the correct API endpoint to get files and folders
    items_url = f"{site_url}/_api/web/lists(guid'{library_id}')/items?$expand=File,Folder&$select=Id,Title,File/Name,File/ServerRelativeUrl,File/Length,File/TimeCreated,File/TimeLastModified,Created,Modified,FileSystemObjectType"
    all_items = []
    next_url = items_url
    
    while next_url:
        response = make_sharepoint_request(next_url, access_token)
        if not response:
            break
        
        if 'd' in response and 'results' in response['d']:
            all_items.extend(response['d']['results'])
        
        # Check for next page
        next_url = None
        if '__next' in response.get('d', {}):
            next_url = response['d']['__next']
    
    # Filter to only files (FileSystemObjectType 0 = File)
    return [item for item in all_items if item.get('FileSystemObjectType') == 0]

def get_file_versions(site_url, file_url, access_token):
    """Get all versions of a file"""
    # Build the versions API URL
    versions_url = f"{site_url}/_api/web/GetFileByServerRelativeUrl('{file_url}')/versions"
    
    response = make_sharepoint_request(versions_url, access_token)
    
    if response and 'd' in response and 'results' in response['d']:
        versions = []
        for version in response['d']['results']:
            versions.append({
                'id': version.get('ID', 0),
                'version_label': version.get('VersionLabel', ''),
                'created': version.get('Created', ''),
                'is_current': version.get('IsCurrentVersion', False),
                'size': version.get('Size', 0),
                'checkin_comment': version.get('CheckInComment', '')
            })
        return versions
    
    # If no versions found, return empty list
    return []

def get_file_details(site_url, file_item, access_token):
    """Get detailed information about a file including its versions"""
    try:
        # Extract file information
        file_obj = file_item.get('File', {})
        
        # Get file URL
        file_url = file_obj.get('ServerRelativeUrl', '')
        
        if not file_url:
            # Try alternative methods to get file URL
            if 'FileDirRef' in file_item and 'FileLeafRef' in file_item:
                file_url = f"{file_item.get('FileDirRef', '')}/{file_item.get('FileLeafRef', '')}"
            else:
                # Skip files without URL
                return None
        
        # Get file name
        file_name = file_obj.get('Name', '')
        if not file_name:
            file_name = file_item.get('Title', '')
        
        # Get file creation and modification times
        created = file_obj.get('TimeCreated', 'N/A')
        modified = file_obj.get('TimeLastModified', 'N/A')
        
        # Get current file size
        current_file_size = file_obj.get('Length', 0)
        
        # Get versions
        versions = get_file_versions(site_url, file_url, access_token)
        
        # Calculate version history summary
        version_count = len(versions)
        first_version_date = 'N/A'
        last_version_date = 'N/A'
        total_versions_size = 0
        
        # Calculate total size of all versions including current
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
            'file_name': file_name,
            'file_url': file_url,
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
        print(f"Error getting file details: {str(e)}")
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
        print(f"\nProcessing library: {library['title']}")
        
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
            file_name = file_item.get('Title', 'Unknown')
            print(f"  Processing file {processed}/{total_files}: {file_name}...", end="")
            
            file_data = get_file_details(site_url, file_item, access_token)
            
            if file_data:
                file_data['library'] = library['title']
                file_data['library_id'] = library['id']
                all_file_data.append(file_data)
                print(f" ✓ ({file_data['version_count']} versions, {file_data['total_versions_size_mb']} MB total)")
            else:
                print(" ✗ (Failed to get details)")
            
            # Add a small delay to avoid rate limiting
            time.sleep(0.3)
    
    print(f"\nProcessed {len(all_file_data)} files with version history.")
    return all_file_data

def generate_csv_report(file_data, output_file):
    """Generate CSV report from file version data"""
    try:
        fieldnames = [
            'Library',
            'File Name',
            'File URL',
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
                    'File Name': data.get('file_name', ''),
                    'File URL': data.get('file_url', ''),
                    'Current File Size (MB)': data.get('current_file_size_mb', 0.00),
                    'Version Count': data.get('version_count', 0),
                    'First Version Date': data.get('first_version_formatted', 'N/A'),
                    'Last Version Date': data.get('last_version_formatted', 'N/A'),
                    'Total Versions Size (MB)': data.get('total_versions_size_mb', 0.00),
                    'File Created': data.get('created_formatted', 'N/A'),
                    'File Modified': data.get('modified_formatted', 'N/A')
                })
        
        print(f"\n✓ Report generated successfully: {output_file}")
        print(f"  Total files with version history: {len(file_data)}")
        
    except Exception as e:
        print(f"Error generating CSV report: {str(e)}")

def generate_detailed_version_report(file_data, output_file):
    """Generate a detailed CSV report with individual version information"""
    try:
        detailed_fieldnames = [
            'Library',
            'File Name',
            'File URL',
            'Version ID',
            'Version Label',
            'Version Created',
            'Is Current Version',
            'Version Size (MB)',
            'Check-in Comment'
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
                            'File Name': data.get('file_name', ''),
                            'File URL': data.get('file_url', ''),
                            'Version ID': version.get('id', 0),
                            'Version Label': version.get('version_label', ''),
                            'Version Created': format_datetime(version.get('created', 'N/A')),
                            'Is Current Version': 'Yes' if version.get('is_current', False) else 'No',
                            'Version Size (MB)': bytes_to_mb(version.get('size', 0)),
                            'Check-in Comment': version.get('checkin_comment', '')
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
                    'Total Current Size (MB)': bytes_to_mb(stats['current_size']),
                    'Total Versions Size (MB)': bytes_to_mb(stats['versions_size']),
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
    print(f"Total current file size: {bytes_to_mb(total_current_size)} MB")
    print(f"Total versions storage size: {bytes_to_mb(total_versions_size)} MB")
    print()
    
    print("TOP 10 FILES WITH MOST VERSIONS:")
    print("-"*80)
    for i, file in enumerate(top_files, 1):
        print(f"{i}. {file['file_name']}")
        print(f"   Versions: {file['version_count']}")
        print(f"   Library: {file['library']}")
        print(f"   Current Size: {file['current_file_size_mb']} MB")
        print(f"   Total Versions Size: {file['total_versions_size_mb']} MB")
        print(f"   First Version: {file['first_version_formatted']}")
        print(f"   Last Version: {file['last_version_formatted']}")
        print()
    
    print("="*80)

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    """Main function to generate file version history report"""
    print("="*80)
    print("FILE VERSION HISTORY REPORT GENERATOR")
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
    
    print("\nReport generation completed successfully!")
    print(f"✓ Main Report: {output_file}")
    print(f"✓ Detailed Version Report: {output_file.replace('.csv', '_Detailed_Versions.csv')}")
    print(f"✓ Summary Report: {output_file.replace('.csv', '_Summary.csv')}")

if __name__ == "__main__":
    main()
