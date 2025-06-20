import requests
import json
from urllib.parse import urljoin
import csv
from datetime import datetime, timezone
import pytz
import re

# Configuration
SHAREPOINT_SITE = "https://geekbyteonline.sharepoint.com/sites/New365"
TOKEN_FILE = "tokens.json"
OUTPUT_CSV = f"SharePointContentWithCompliance_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"

# Timezone setup
UTC_TZ = timezone.utc
IST_TZ = pytz.timezone('Asia/Kolkata')

def load_access_token(file_path):
    """Load access token from JSON file"""
    try:
        with open(file_path, 'r') as f:
            token_data = json.load(f)
            return token_data.get('access_token')
    except Exception as e:
        print(f"Error loading access token: {str(e)}")
        return None

def make_sharepoint_request(url, access_token, method='GET', headers=None):
    """Make a request to SharePoint REST API"""
    default_headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json"
    }
    
    if headers:
        default_headers.update(headers)
    
    try:
        response = requests.request(method, url, headers=default_headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Request failed for {url}: {str(e)}")
        return None

def normalize_datetime(dt_str):
    """Normalize SharePoint datetime string to consistent format"""
    if not dt_str or dt_str == "N/A":
        return None
    
    # Remove timezone offset if present (e.g., "+05:30")
    dt_str = re.sub(r'[+-]\d{2}:\d{2}$', '', dt_str)
    
    # Add Z if missing at end
    if not dt_str.endswith('Z'):
        dt_str += 'Z'
    
    return dt_str

def convert_utc_to_ist(utc_datetime_str):
    """Convert UTC datetime string to IST"""
    utc_datetime_str = normalize_datetime(utc_datetime_str)
    if not utc_datetime_str:
        return "N/A"
    
    try:
        # Handle both formats: with and without milliseconds
        if '.' in utc_datetime_str:
            utc_dt = datetime.strptime(utc_datetime_str, "%Y-%m-%dT%H:%M:%S.%fZ").replace(tzinfo=UTC_TZ)
        else:
            utc_dt = datetime.strptime(utc_datetime_str, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=UTC_TZ)
        
        ist_dt = utc_dt.astimezone(IST_TZ)
        return ist_dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception as e:
        print(f"Error converting datetime '{utc_datetime_str}': {str(e)}")
        return "N/A"

def get_all_lists(site_url, access_token):
    """Get all document libraries from SharePoint site"""
    lists_url = urljoin(site_url + "/", "_api/web/lists")
    response = make_sharepoint_request(lists_url, access_token)
    
    if response and 'd' in response and 'results' in response['d']:
        return [lst for lst in response['d']['results'] if lst['BaseTemplate'] == 101]
    return []

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
            
        # Check for next page
        next_url = None
        if '__next' in response.get('d', {}):
            next_url = response['d']['__next']
            
    return all_items

def get_file_properties(site_url, list_id, item_id, access_token):
    """Get file properties including compliance information"""
    props_url = f"{site_url}/_api/web/lists(guid'{list_id}')/items({item_id})/file/properties"
    response = make_sharepoint_request(props_url, access_token)
    
    if response and 'd' in response:
        return response['d']
    return {}

def process_item(site_url, list_id, item, access_token):
    """Extract relevant details from an item"""
    item_type = item['FileSystemObjectType']
    details = {
        'Type': 'File' if item_type == 0 else 'Folder',
        'ID': item['Id'],
        'Name': '',
        'Path': '',
        'Size': 0 if item_type == 0 else 'N/A',
        'Created': item.get('Created', 'N/A'),
        'Modified': item.get('Modified', 'N/A'),
        'Author': item.get('Author', {}).get('Title', 'N/A'),
        'Editor': item.get('Editor', {}).get('Title', 'N/A'),
        'ComplianceTag': 'N/A',
        'ComplianceTagWrittenTime_UTC': 'N/A',
        'ComplianceTagWrittenTime_IST': 'N/A',
        'LastModified_UTC': 'N/A',
        'LastModified_IST': 'N/A'
    }
    
    if item_type == 0:  # File
        if 'File' in item:
            file = item['File']
            details['Name'] = file.get('Name', '')
            details['Path'] = file.get('ServerRelativeUrl', '')
            details['Size'] = file.get('Length', 0)
            
            # Get additional file properties
            file_props = get_file_properties(site_url, list_id, item['Id'], access_token)
            if file_props:
                last_modified = file_props.get('vti_x005f_timelastmodified')
                details['LastModified_UTC'] = last_modified if last_modified else 'N/A'
                details['LastModified_IST'] = convert_utc_to_ist(last_modified)
                
                compliance_tag = file_props.get('vti_x005f_complianceTag')
                details['ComplianceTag'] = compliance_tag if compliance_tag else 'N/A'
                
                compliance_time = file_props.get('vti_x005f_complianceTagWrittenTime')
                details['ComplianceTagWrittenTime_UTC'] = compliance_time if compliance_time else 'N/A'
                details['ComplianceTagWrittenTime_IST'] = convert_utc_to_ist(compliance_time)
    else:  # Folder
        if 'Folder' in item:
            folder = item['Folder']
            details['Name'] = folder.get('Name', '')
            details['Path'] = folder.get('ServerRelativeUrl', '')
    
    return details

def main():
    # Load access token
    access_token = load_access_token(TOKEN_FILE)
    if not access_token:
        print("Failed to load access token. Exiting.")
        return
    
    print("Successfully connected to SharePoint")
    
    # Get all document libraries
    print("Retrieving document libraries...")
    libraries = get_all_lists(SHAREPOINT_SITE, access_token)
    
    if not libraries:
        print("No document libraries found.")
        return
    
    # Prepare CSV output
    with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = [
            'Library', 'Type', 'ID', 'Name', 'Path', 'Size', 
            'Created', 'Modified', 'Author', 'Editor',
            'LastModified_UTC', 'LastModified_IST',
            'ComplianceTag', 'ComplianceTagWrittenTime_UTC', 'ComplianceTagWrittenTime_IST'
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        
        # Process each library
        for library in libraries:
            print(f"\nProcessing library: {library['Title']}")
            
            # Get all items in the library
            items = get_list_items(SHAREPOINT_SITE, library['Id'], access_token)
            
            if not items:
                print("No items found in this library")
                continue
            
            # Process each item
            for item in items:
                try:
                    details = process_item(SHAREPOINT_SITE, library['Id'], item, access_token)
                    details['Library'] = library['Title']
                    writer.writerow(details)
                    
                    # Print to console (optional)
                    print(f"{details['Type']}: {details['Name']} (Compliance: {details['ComplianceTag']})")
                except Exception as e:
                    print(f"Error processing item {item.get('Id', 'unknown')}: {str(e)}")
    
    print(f"\nReport generated successfully: {OUTPUT_CSV}")

if __name__ == "__main__":
    main()
