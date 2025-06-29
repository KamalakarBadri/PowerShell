import requests
import json
from urllib.parse import urljoin
import csv
from datetime import datetime, timezone
import pytz
import re
import uuid
import base64
import time
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend

# Configuration
SHAREPOINT_SITE = "https://test.sharepoint.com/sites/New365"
OUTPUT_CSV = f"SharePointContent_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"

# Authentication Configuration
TENANT_NAME = "test.onmicrosoft.com"
APP_ID = "73efa35d-6188-42d4-b258-838a977eb149"
SCOPE_SHAREPOINT = "https://test.sharepoint.com/.default"

# Certificate paths (update these with your actual file paths)
CERTIFICATE_PATH = "certificate.pem"
PRIVATE_KEY_PATH = "private_key.pem"

# Timezone setup
UTC_TZ = timezone.utc
IST_TZ = pytz.timezone('Asia/Kolkata')

def load_certificate_and_key():
    """Load certificate and private key from PEM files"""
    try:
        # Load certificate
        with open(CERTIFICATE_PATH, "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())

        # Load private key
        with open(PRIVATE_KEY_PATH, "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())

        return certificate, private_key

    except Exception as e:
        print(f"Error loading certificate or private key: {str(e)}")
        raise

def get_jwt_token(certificate, private_key, scope):
    """Generate JWT token using certificate and private key"""
    try:
        # Create JWT timestamp for expiration (5 minutes from now)
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
            "aud": f"https://login.microsoftonline.com/{TENANT_NAME}/oauth2/v2.0/token",
            "exp": expiration,
            "iss": APP_ID,
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": APP_ID
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
    url = f"https://login.microsoftonline.com/{TENANT_NAME}/oauth2/v2.0/token"
    
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    
    data = {
        "client_id": APP_ID,
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

def authenticate_sharepoint():
    """Authenticate and get SharePoint access token using certificate"""
    try:
        print("Loading certificate and private key...")
        certificate, private_key = load_certificate_and_key()
        
        print("Generating JWT token...")
        jwt = get_jwt_token(certificate, private_key, SCOPE_SHAREPOINT)
        
        print("Getting access token...")
        access_token = get_access_token(jwt, SCOPE_SHAREPOINT)
        
        print("Successfully authenticated to SharePoint")
        return access_token
        
    except Exception as e:
        print(f"Authentication failed: {str(e)}")
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

def convert_utc_to_ist(utc_datetime_str):
    """Convert UTC datetime string to IST"""
    if not utc_datetime_str or utc_datetime_str == "N/A":
        return "N/A"
    
    try:
        # Remove timezone offset if present (e.g., "+05:30")
        utc_datetime_str = re.sub(r'[+-]\d{2}:\d{2}$', '', utc_datetime_str)
        
        # Add Z if missing at end
        if not utc_datetime_str.endswith('Z'):
            utc_datetime_str += 'Z'
        
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

def process_item(item):
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
        'Created_IST': convert_utc_to_ist(item.get('Created', 'N/A')),
        'Modified_IST': convert_utc_to_ist(item.get('Modified', 'N/A')),
        'Author': item.get('Author', {}).get('Title', 'N/A') if item.get('Author') else 'N/A',
        'Editor': item.get('Editor', {}).get('Title', 'N/A') if item.get('Editor') else 'N/A'
    }
    
    if item_type == 0:  # File
        if 'File' in item:
            file = item['File']
            details['Name'] = file.get('Name', '')
            details['Path'] = file.get('ServerRelativeUrl', '')
            details['Size'] = file.get('Length', 0)
    else:  # Folder
        if 'Folder' in item:
            folder = item['Folder']
            details['Name'] = folder.get('Name', '')
            details['Path'] = folder.get('ServerRelativeUrl', '')
    
    return details

def main():
    # Authenticate to SharePoint
    access_token = authenticate_sharepoint()
    if not access_token:
        print("Failed to authenticate to SharePoint. Exiting.")
        return
    
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
            'Created', 'Modified', 'Created_IST', 'Modified_IST',
            'Author', 'Editor'
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
                    details = process_item(item)
                    details['Library'] = library['Title']
                    writer.writerow(details)
                    
                    # Print to console (optional)
                    print(f"{details['Type']}: {details['Name']}")
                except Exception as e:
                    print(f"Error processing item {item.get('Id', 'unknown')}: {str(e)}")
    
    print(f"\nReport generated successfully: {OUTPUT_CSV}")

if __name__ == "__main__":
    main()
