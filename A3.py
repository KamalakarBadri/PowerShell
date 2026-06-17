import requests
import json
import csv
import uuid
import base64
import time
import os
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
    # Get all lists with pagination
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

def get_list_details(site_url, list_id, access_token):
    """Get additional details for a specific list"""
    # This is optional - can get more details if needed
    list_url = f"{site_url}/_api/web/lists(guid'{list_id}')"
    response = make_sharepoint_request(list_url, access_token)
    
    if response and 'd' in response:
        return response['d']
    return None

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
        return []
    
    print(f"Found {len(lists_data)} lists/libraries.")
    
    # Process each list
    processed_lists = []
    
    for list_item in lists_data:
        # Extract required fields
        title = list_item.get('Title', '')
        item_count = list_item.get('ItemCount', 0)
        base_template = list_item.get('BaseTemplate', 0)
        hidden = list_item.get('Hidden', False)
        list_id = list_item.get('Id', '')
        
        # Get additional details
        created = list_item.get('Created', 'N/A')
        description = list_item.get('Description', '')
        enable_versioning = list_item.get('EnableVersioning', False)
        enable_minor_versions = list_item.get('EnableMinorVersions', False)
        is_application_list = list_item.get('IsApplicationList', False)
        is_catalog = list_item.get('IsCatalog', False)
        
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
            'IsDocumentLibrary': 'Yes' if is_document_library else 'No',
            'Created': format_datetime(created),
            'Description': description,
            'EnableVersioning': 'Yes' if enable_versioning else 'No',
            'EnableMinorVersions': 'Yes' if enable_minor_versions else 'No',
            'IsApplicationList': 'Yes' if is_application_list else 'No',
            'IsCatalog': 'Yes' if is_catalog else 'No'
        }
        
        processed_lists.append(processed_list)
    
    return processed_lists

def generate_csv_report(lists_data, output_file):
    """Generate CSV report from lists data"""
    try:
        # Define field order for CSV
        fieldnames = [
            'Title',
            'ID',
            'ItemCount',
            'BaseTemplate',
            'TemplateName',
            'Category',
            'Hidden',
            'IsDocumentLibrary',
            'Created',
            'Description',
            'EnableVersioning',
            'EnableMinorVersions',
            'IsApplicationList',
            'IsCatalog'
        ]
        
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            # Sort by Category, then by Title
            sorted_data = sorted(lists_data, key=lambda x: (x['Category'], x['Title']))
            
            for data in sorted_data:
                writer.writerow(data)
        
        print(f"\n✓ Report generated successfully: {output_file}")
        print(f"  Total lists/libraries: {len(lists_data)}")
        
    except Exception as e:
        print(f"Error generating CSV report: {str(e)}")

def print_summary_stats(lists_data):
    """Print summary statistics about the lists"""
    if not lists_data:
        print("\nNo data to summarize.")
        return
    
    total_lists = len(lists_data)
    
    # Count by category
    category_counts = {}
    hidden_count = 0
    doc_library_count = 0
    total_items = 0
    
    for list_item in lists_data:
        category = list_item.get('Category', 'Unknown')
        category_counts[category] = category_counts.get(category, 0) + 1
        
        if list_item.get('Hidden') == 'Yes':
            hidden_count += 1
        
        if list_item.get('IsDocumentLibrary') == 'Yes':
            doc_library_count += 1
        
        total_items += list_item.get('ItemCount', 0)
    
    print("\n" + "="*80)
    print("SHAREPOINT LISTS AND LIBRARIES SUMMARY")
    print("="*80)
    print(f"Total lists/libraries: {total_lists}")
    print(f"Total items across all lists: {total_items:,}")
    print(f"Document libraries: {doc_library_count}")
    print(f"Hidden lists: {hidden_count}")
    print("\nLists by Category:")
    
    # Sort categories by count
    sorted_categories = sorted(category_counts.items(), key=lambda x: x[1], reverse=True)
    for category, count in sorted_categories:
        print(f"  {category}: {count}")
    
    print("="*80)

def print_top_lists(lists_data):
    """Print top lists by item count"""
    if not lists_data:
        return
    
    # Filter out hidden lists and sort by item count
    visible_lists = [l for l in lists_data if l.get('Hidden') == 'No']
    sorted_lists = sorted(visible_lists, key=lambda x: x['ItemCount'], reverse=True)
    
    # Get top 10
    top_lists = sorted_lists[:10]
    
    if top_lists:
        print("\nTOP 10 LISTS/LIBRARIES BY ITEM COUNT:")
        print("-"*80)
        for i, list_item in enumerate(top_lists, 1):
            print(f"{i:2d}. {list_item['Title']}")
            print(f"    Items: {list_item['ItemCount']:,}")
            print(f"    Category: {list_item['Category']}")
            print(f"    Template: {list_item['TemplateName']}")
            print()

def generate_category_report(lists_data, output_file):
    """Generate a summary report by category"""
    try:
        # Create category summary
        category_summary = {}
        for list_item in lists_data:
            category = list_item.get('Category', 'Unknown')
            if category not in category_summary:
                category_summary[category] = {
                    'count': 0,
                    'items': 0,
                    'hidden': 0,
                    'doc_libraries': 0
                }
            
            category_summary[category]['count'] += 1
            category_summary[category]['items'] += list_item.get('ItemCount', 0)
            
            if list_item.get('Hidden') == 'Yes':
                category_summary[category]['hidden'] += 1
            
            if list_item.get('IsDocumentLibrary') == 'Yes':
                category_summary[category]['doc_libraries'] += 1
        
        # Create summary output filename
        summary_output = output_file.replace('.csv', '_Category_Summary.csv')
        
        with open(summary_output, 'w', newline='', encoding='utf-8-sig') as csvfile:
            fieldnames = ['Category', 'Total Lists', 'Total Items', 'Hidden Lists', 'Document Libraries']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for category, stats in sorted(category_summary.items()):
                writer.writerow({
                    'Category': category,
                    'Total Lists': stats['count'],
                    'Total Items': stats['items'],
                    'Hidden Lists': stats['hidden'],
                    'Document Libraries': stats['doc_libraries']
                })
        
        print(f"✓ Category summary report generated: {summary_output}")
        
    except Exception as e:
        print(f"Error generating category summary: {str(e)}")

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    """Main function to generate SharePoint lists report"""
    print("="*80)
    print("SHAREPOINT LISTS AND LIBRARIES REPORT GENERATOR")
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
    print("Starting list discovery...")
    start_time = time.time()
    
    lists_data = process_lists(site_url, access_token)
    
    elapsed_time = time.time() - start_time
    print(f"\nProcessing completed in {elapsed_time:.2f} seconds.")
    
    if not lists_data:
        print("\nNo lists or libraries were found.")
        return
    
    # Generate reports
    output_file = CONFIG['output_csv']
    
    # Main report
    generate_csv_report(lists_data, output_file)
    
    # Category summary report
    generate_category_report(lists_data, output_file)
    
    # Print summary statistics
    print_summary_stats(lists_data)
    
    # Print top lists
    print_top_lists(lists_data)
    
    print("\nReport generation completed successfully!")
    print(f"✓ Main Report: {output_file}")
    print(f"✓ Category Summary: {output_file.replace('.csv', '_Category_Summary.csv')}")
    
    # Print sample of the data
    print("\n" + "="*80)
    print("SAMPLE DATA (First 5 visible lists):")
    print("="*80)
    
    visible_lists = [l for l in lists_data if l.get('Hidden') == 'No'][:5]
    for i, list_item in enumerate(visible_lists, 1):
        print(f"\n{i}. Title: {list_item['Title']}")
        print(f"   Items: {list_item['ItemCount']:,}")
        print(f"   Template: {list_item['TemplateName']} (BaseTemplate: {list_item['BaseTemplate']})")
        print(f"   Category: {list_item['Category']}")
        print(f"   Created: {list_item['Created']}")

if __name__ == "__main__":
    main()
