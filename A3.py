import requests
import json
import csv
import uuid
import base64
import time
import os
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
    "output_csv": f"SharePoint_Lists_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"
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
# LIST/LIBRARY RETRIEVAL FUNCTIONS
# ============================================================

def get_all_lists(site_url, access_token):
    """Get all lists and document libraries from SharePoint site"""
    lists_url = f"{site_url}/_api/web/lists"
    all_lists = []
    next_url = lists_url
    
    while next_url:
        print(f"  Fetching lists from: {next_url}")
        response = make_sharepoint_request(next_url, access_token)
        
        if not response:
            break
        
        if 'd' in response and 'results' in response['d']:
            all_lists.extend(response['d']['results'])
        
        # Check for next page
        next_url = None
        if '__next' in response.get('d', {}):
            next_url = response['d']['__next']
    
    return all_lists

def get_list_type_name(base_template):
    """Convert BaseTemplate number to readable name"""
    list_types = {
        100: "Generic List",
        101: "Document Library",
        102: "Survey",
        103: "Links",
        104: "Announcements",
        105: "Contacts",
        106: "Events/Calendar",
        107: "Tasks",
        108: "Discussion Board",
        109: "Picture Library",
        110: "Data Sources",
        111: "Site Template Gallery",
        112: "User Information",
        113: "Web Part Gallery",
        114: "List Template Gallery",
        115: "XML Form Library",
        116: "Master Page Gallery",
        117: "No Code Workflows",
        118: "Workflow History",
        119: "Gantt Tasks",
        120: "Meetings",
        121: "Agenda",
        122: "Meeting Workspace Pages",
        123: "Discussion Board",
        124: "Administrator Tasks",
        125: "Issue Tracking",
        126: "Blog Posts",
        127: "Blog Comments",
        128: "Blog Categories",
        130: "Data Connection Library",
        140: "Workflow Tasks",
        150: "Site Assets",
        151: "Site Pages",
        160: "App Data",
        170: "App List",
        171: "App Attachments",
        172: "App File",
        173: "App Folder",
        174: "App Page",
        175: "App Launch Points",
        200: "Form Library",
        300: "Wiki Page Library",
        400: "Custom List",
        401: "Custom List in Datasheet View",
        402: "External List",
        403: "Custom List with Approval Workflow",
        404: "Custom List in Datasheet View with Approval Workflow",
        405: "Custom List with Content Approval",
        406: "Custom List in Datasheet View with Content Approval",
        407: "Custom List with Workflow",
        408: "Custom List in Datasheet View with Workflow",
        409: "Custom List with Content Approval and Workflow",
        410: "Custom List in Datasheet View with Content Approval and Workflow",
        500: "Report Library",
        600: "Project Tasks",
        700: "Project Issues",
        701: "Project Risks",
        702: "Project Deliverables"
    }
    return list_types.get(base_template, f"Unknown ({base_template})")

def get_list_category(base_template):
    """Categorize lists by BaseTemplate"""
    if base_template == 101:
        return "Document Library"
    elif base_template in [100, 400, 401, 402, 403, 404, 405, 406, 407, 408, 409, 410]:
        return "List"
    elif base_template in [106, 107, 119, 120, 121, 122, 125, 600, 700, 701, 702]:
        return "Task/Calendar"
    elif base_template in [102, 103, 104, 105, 108, 123]:
        return "Communication"
    elif base_template in [109, 200, 500]:
        return "Library"
    elif base_template in [110, 111, 112, 113, 114, 115, 116, 130, 140, 150, 151, 160, 170, 171, 172, 173, 174, 175]:
        return "System/Special"
    else:
        return "Other"

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

def bytes_to_mb(bytes_value):
    """Convert bytes to MB with 2 decimal places"""
    try:
        if isinstance(bytes_value, str):
            bytes_value = float(bytes_value) if bytes_value else 0
        if bytes_value == 0:
            return 0.00
        return round(float(bytes_value) / (1024 * 1024), 2)
    except:
        return 0.00

# ============================================================
# VERSION HISTORY FUNCTIONS USING ITEM ID
# ============================================================

def get_file_versions_by_item_id(site_url, list_id, item_id, access_token):
    """
    Get file versions using Item ID
    API Endpoint: /_api/Web/Lists(guid'{list_id}')/items({item_id})/versions
    """
    try:
        # Build the correct URL for getting versions by item ID
        versions_url = f"{site_url}/_api/Web/Lists(guid'{list_id}')/items({item_id})/versions"
        
        print(f"\n        Fetching versions from: {versions_url}")
        response = make_sharepoint_request(versions_url, access_token)
        
        if not response or 'd' not in response:
            return []
        
        versions = []
        if 'results' in response['d']:
            for version in response['d']['results']:
                # Extract version information from the response
                version_data = {
                    'id': version.get('VersionId', 0),
                    'version_label': version.get('VersionLabel', ''),
                    'version_id': version.get('VersionId', 0),
                    'ui_version': version.get('OData__x005f_UIVersion', 0),
                    'ui_version_string': version.get('OData__x005f_UIVersionString', ''),
                    'created': version.get('Created', ''),
                    'is_current': version.get('IsCurrentVersion', False),
                    'size': version.get('File_x005f_x0020_x005f_Size', '0'),
                    'checkin_comment': version.get('OData__x005f_CheckinComment', ''),
                    'file_ref': version.get('FileRef', ''),
                    'file_leaf_ref': version.get('FileLeafRef', ''),
                    'author': version.get('Author', {}).get('LookupValue', ''),
                    'editor': version.get('Editor', {}).get('LookupValue', ''),
                    'modified': version.get('Modified', ''),
                    'file_size': version.get('File_x005f_x0020_x005f_Size', '0')
                }
                versions.append(version_data)
        
        return versions
        
    except Exception as e:
        print(f"Error getting versions for item {item_id}: {str(e)}")
        return []

def get_file_details_by_item_id(site_url, list_id, item_id, access_token):
    """Get file details using Item ID"""
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
        fsob_type = item_data.get('FSObjType', '0')  # '0' = File, '1' = Folder
        
        # If it's a folder, skip
        if fsob_type == '1':
            return None
        
        file_name = title if title else file_leaf_ref
        
        # Get versions using Item ID
        versions = get_file_versions_by_item_id(site_url, list_id, item_id, access_token)
        
        # Calculate version history summary
        version_count = len(versions)
        first_version_date = 'N/A'
        last_version_date = 'N/A'
        total_versions_size = 0
        current_file_size = 0
        
        # Get current file size from the response
        try:
            current_file_size = float(file_size) if file_size else 0
        except:
            current_file_size = 0
        
        if versions:
            # Sort versions by creation date
            sorted_versions = sorted(versions, key=lambda x: x.get('created', ''))
            first_version = sorted_versions[0]
            last_version = sorted_versions[-1]
            
            first_version_date = first_version.get('created', 'N/A')
            last_version_date = last_version.get('created', 'N/A')
            
            # Calculate total size of all versions
            for version in versions:
                try:
                    version_size = float(version.get('size', '0')) if version.get('size', '0') else 0
                    total_versions_size += version_size
                except:
                    pass
        
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

def get_all_items_in_library(site_url, library_id, access_token):
    """Get all items from a document library with pagination"""
    # Only get items that are files (FSObjType = 0)
    items_url = f"{site_url}/_api/Web/Lists(guid'{library_id}')/items?$filter=FSObjType eq 0&$select=Id,Title,FSObjType,Created,Modified,FileLeafRef,FileRef,File_x005f_x0020_x005f_Size"
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
    
    return all_items

def process_library_items(site_url, library, access_token):
    """Process all items in a library and get their version details using Item ID"""
    library_id = library['id']
    library_title = library['title']
    
    print(f"\n  Processing library: {library_title}")
    
    # Get all items in the library
    items = get_all_items_in_library(site_url, library_id, access_token)
    
    if not items:
        print(f"    No files found in {library_title}")
        return []
    
    print(f"    Found {len(items)} files")
    
    processed_items = []
    processed_count = 0
    
    for item in items:
        processed_count += 1
        item_id = item.get('Id')
        title = item.get('Title', '')
        file_leaf_ref = item.get('FileLeafRef', f'Item_{item_id}')
        file_name = title if title else file_leaf_ref
        
        print(f"\n    [{processed_count}/{len(items)}] Processing: {file_name} (ID: {item_id})", end="")
        
        # Get file details using Item ID
        file_details = get_file_details_by_item_id(site_url, library_id, item_id, access_token)
        
        if file_details:
            file_details['library_title'] = library_title
            file_details['library_id'] = library_id
            processed_items.append(file_details)
            print(f" ✓ ({file_details['version_count']} versions, {file_details['total_versions_size_mb']:.2f} MB total)")
        else:
            print(" ✗ (Failed or not a file)")
        
        # Add a small delay to avoid rate limiting
        time.sleep(0.5)
    
    return processed_items

# ============================================================
# MAIN PROCESSING FUNCTIONS
# ============================================================

def process_lists(site_url, access_token):
    """Process all lists and libraries and generate report data"""
    print("\nFetching all lists and libraries...")
    
    # Get all lists
    lists_data = get_all_lists(site_url, access_token)
    
    if not lists_data:
        print("No lists or libraries found.")
        return [], []
    
    print(f"Found {len(lists_data)} lists/libraries.")
    
    # Process each list for summary
    processed_lists = []
    all_file_data = []
    
    for list_item in lists_data:
        # Extract required fields
        title = list_item.get('Title', '')
        item_count = list_item.get('ItemCount', 0)
        base_template = list_item.get('BaseTemplate', 0)
        hidden = list_item.get('Hidden', False)
        list_id = list_item.get('Id', '')
        
        # Get list type name and category
        list_type_name = get_list_type_name(base_template)
        list_category = get_list_category(base_template)
        
        # Determine if it's a document library
        is_document_library = (base_template == 101)
        
        processed_list = {
            'Title': title,
            'ID': list_id,
            'ItemCount': item_count,
            'BaseTemplate': base_template,
            'TemplateName': list_type_name,
            'Category': list_category,
            'Hidden': 'Yes' if hidden else 'No',
            'IsDocumentLibrary': 'Yes' if is_document_library else 'No'
        }
        
        processed_lists.append(processed_list)
        
        # If it's a document library, process its items for version history
        if is_document_library and not hidden:
            print(f"\n{'='*60}")
            print(f"PROCESSING DOCUMENT LIBRARY: {title}")
            print(f"{'='*60}")
            
            library_info = {
                'id': list_id,
                'title': title
            }
            
            library_items = process_library_items(site_url, library_info, access_token)
            all_file_data.extend(library_items)
            
            if library_items:
                print(f"\n  Completed processing {len(library_items)} files from {title}")
            else:
                print(f"\n  No files processed from {title}")
    
    return processed_lists, all_file_data

def generate_list_summary_report(lists_data, output_file):
    """Generate CSV report for lists summary"""
    try:
        fieldnames = [
            'Title',
            'ID',
            'ItemCount',
            'BaseTemplate',
            'TemplateName',
            'Category',
            'Hidden',
            'IsDocumentLibrary'
        ]
        
        # Create list summary output filename
        list_output = output_file.replace('.csv', '_List_Summary.csv')
        
        with open(list_output, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            sorted_data = sorted(lists_data, key=lambda x: (x['Category'], x['Title']))
            
            for data in sorted_data:
                writer.writerow(data)
        
        print(f"\n✓ List summary report generated: {list_output}")
        
    except Exception as e:
        print(f"Error generating list summary report: {str(e)}")

def generate_version_history_report(file_data, output_file):
    """Generate CSV report for version history"""
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
        
        # Create version history output filename
        version_output = output_file.replace('.csv', '_Version_History.csv')
        
        with open(version_output, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            # Sort by version count descending
            sorted_data = sorted(file_data, key=lambda x: x.get('version_count', 0), reverse=True)
            
            for data in sorted_data:
                writer.writerow({
                    'Library': data.get('library_title', ''),
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
        
        print(f"✓ Version history report generated: {version_output}")
        
    except Exception as e:
        print(f"Error generating version history report: {str(e)}")

def generate_detailed_version_report(file_data, output_file):
    """Generate detailed version report with individual versions"""
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
        
        detailed_output = output_file.replace('.csv', '_Detailed_Versions.csv')
        
        with open(detailed_output, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=detailed_fieldnames)
            writer.writeheader()
            
            for data in file_data:
                if data.get('versions'):
                    for version in data['versions']:
                        writer.writerow({
                            'Library': data.get('library_title', ''),
                            'Item ID': data.get('item_id', 0),
                            'File Name': data.get('file_name', ''),
                            'File Path': data.get('file_ref', ''),
                            'Version ID': version.get('id', 0),
                            'Version Label': version.get('version_label', ''),
                            'UI Version': version.get('ui_version_string', ''),
                            'Version Created': format_datetime(version.get('created', 'N/A')),
                            'Is Current Version': 'Yes' if version.get('is_current', False) else 'No',
                            'Version Size (MB)': f"{bytes_to_mb(version.get('size', '0')):.2f}",
                            'Check-in Comment': version.get('checkin_comment', ''),
                            'Author': version.get('author', ''),
                            'Editor': version.get('editor', '')
                        })
        
        print(f"✓ Detailed version report generated: {detailed_output}")
        
    except Exception as e:
        print(f"Error generating detailed version report: {str(e)}")

def print_summary_stats(lists_data, file_data):
    """Print summary statistics"""
    if not lists_data:
        print("\nNo data to summarize.")
        return
    
    total_lists = len(lists_data)
    doc_libraries = [l for l in lists_data if l.get('IsDocumentLibrary') == 'Yes']
    hidden_lists = [l for l in lists_data if l.get('Hidden') == 'Yes']
    
    total_items = sum(l.get('ItemCount', 0) for l in lists_data)
    
    print("\n" + "="*80)
    print("SHAREPOINT LISTS AND VERSION HISTORY SUMMARY")
    print("="*80)
    print(f"Total lists/libraries: {total_lists}")
    print(f"Document libraries: {len(doc_libraries)}")
    print(f"Hidden lists: {len(hidden_lists)}")
    print(f"Total items across all lists: {total_items:,}")
    
    if file_data:
        total_files = len(file_data)
        total_versions = sum(f.get('version_count', 0) for f in file_data)
        files_with_versions = len([f for f in file_data if f.get('version_count', 0) > 0])
        total_current_size = sum(f.get('current_file_size', 0) for f in file_data)
        total_versions_size = sum(f.get('total_versions_size', 0) for f in file_data)
        
        print(f"\nFiles processed with version history: {total_files}")
        print(f"Files with versions: {files_with_versions}")
        print(f"Total version records: {total_versions}")
        print(f"Total current file size: {bytes_to_mb(total_current_size):.2f} MB")
        print(f"Total versions storage size: {bytes_to_mb(total_versions_size):.2f} MB")
        
        # Find files with most versions
        if file_data:
            top_files = sorted(file_data, key=lambda x: x.get('version_count', 0), reverse=True)[:10]
            print("\nTOP 10 FILES WITH MOST VERSIONS:")
            print("-"*80)
            for i, file in enumerate(top_files, 1):
                if file.get('version_count', 0) > 0:
                    print(f"{i:2d}. {file.get('file_name', 'Unknown')}")
                    print(f"    Versions: {file.get('version_count', 0)}")
                    print(f"    Library: {file.get('library_title', '')}")
                    print(f"    Item ID: {file.get('item_id', 0)}")
                    print(f"    Total Size: {file.get('total_versions_size_mb', 0.00):.2f} MB")
                    print()
    
    print("="*80)

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    """Main function to generate SharePoint lists and version history report"""
    print("="*80)
    print("SHAREPOINT LISTS AND VERSION HISTORY REPORT GENERATOR")
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
    
    # Process all lists
    print("Starting list discovery and version history analysis...")
    start_time = time.time()
    
    lists_data, file_data = process_lists(site_url, access_token)
    
    elapsed_time = time.time() - start_time
    print(f"\nProcessing completed in {elapsed_time:.2f} seconds.")
    
    if not lists_data:
        print("\nNo lists or libraries were found.")
        return
    
    # Generate reports
    output_file = CONFIG['output_csv']
    
    # List summary report
    generate_list_summary_report(lists_data, output_file)
    
    # Version history report
    if file_data:
        generate_version_history_report(file_data, output_file)
        generate_detailed_version_report(file_data, output_file)
    else:
        print("\nNo files with version history were found.")
    
    # Print summary statistics
    print_summary_stats(lists_data, file_data)
    
    print("\nReport generation completed successfully!")
    print(f"✓ List Summary Report: {output_file.replace('.csv', '_List_Summary.csv')}")
    if file_data:
        print(f"✓ Version History Report: {output_file.replace('.csv', '_Version_History.csv')}")
        print(f"✓ Detailed Version Report: {output_file.replace('.csv', '_Detailed_Versions.csv')}")
    
    print("\n" + "="*80)
    print("SAMPLE DATA (First 5 visible document libraries):")
    print("="*80)
    
    visible_doc_libs = [l for l in lists_data if l.get('IsDocumentLibrary') == 'Yes' and l.get('Hidden') == 'No'][:5]
    for i, lib in enumerate(visible_doc_libs, 1):
        print(f"\n{i}. Library: {lib['Title']}")
        print(f"   Items: {lib['ItemCount']:,}")
        print(f"   Template: {lib['TemplateName']}")
        print(f"   Category: {lib['Category']}")

if __name__ == "__main__":
    main()
