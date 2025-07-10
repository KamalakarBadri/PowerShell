import requests
import csv
import json
import uuid
import base64
import time
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend

# Configuration
TENANT_NAME = ""
CLIENT_ID = ""
SCOPE = "https://graph.microsoft.com/.default"

# Certificate paths (update these with your actual file paths)
CERTIFICATE_PATH = "certificate.pem"
PRIVATE_KEY_PATH = "private_key.pem"

# Sites to search for
SITE_NAMES = ["New365", "AnotherSite", "ThirdSite"]

# Threading configuration
MAX_WORKERS = 5  # Number of parallel threads for permission checking
REQUEST_DELAY = 0.100  # 100 milliseconds delay between permission requests

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

def get_jwt_token(certificate, private_key):
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
            "iss": CLIENT_ID,
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": CLIENT_ID
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

def get_access_token():
    """Get access token using certificate-based authentication"""
    try:
        # Load certificate and private key
        certificate, private_key = load_certificate_and_key()
        print("Certificate and private key loaded successfully")
        
        # Generate JWT
        jwt = get_jwt_token(certificate, private_key)
        print("Generated JWT token")
        
        # Get access token
        url = f"https://login.microsoftonline.com/{TENANT_NAME}/oauth2/v2.0/token"
        
        headers = {
            "Content-Type": "application/x-www-form-urlencoded"
        }
        
        data = {
            "client_id": CLIENT_ID,
            "client_assertion": jwt,
            "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            "scope": SCOPE,
            "grant_type": "client_credentials"
        }
        
        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()
        access_token = response.json()["access_token"]
        print("Access token retrieved successfully")
        return access_token
        
    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error: {err}")
        print(f"Response: {response.text}")
        return None
    except Exception as e:
        print(f"Failed to get access token: {e}")
        return None

def get_site_guid(site_name, token):
    """Get the site GUID from Microsoft Graph API"""
    endpoint = f"https://graph.microsoft.com/v1.0/sites?search={site_name}"
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(endpoint, headers=headers)
        response.raise_for_status()
        data = response.json()
        
        for site in data.get('value', []):
            if site['name'].lower() == site_name.lower():
                full_id = site['id']
                site_guid = full_id.split(',')[1]
                return {
                    'site_name': site_name,
                    'full_id': full_id,
                    'site_guid': site_guid,
                    'web_url': site['webUrl'],
                    'display_name': site['displayName']
                }
        
        print(f"No exact match found for site: {site_name}")
        return None
        
    except Exception as e:
        print(f"Error searching for site {site_name}: {e}")
        return None

def get_document_libraries(site_guid, token):
    """Get all document libraries (drives) for a site"""
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_guid}/drives"
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(endpoint, headers=headers)
        response.raise_for_status()
        data = response.json()
        
        libraries = []
        for drive in data.get('value', []):
            if drive.get('driveType') == 'documentLibrary':
                libraries.append({
                    'id': drive['id'],
                    'name': drive['name'],
                    'webUrl': drive['webUrl']
                })
        
        return libraries
        
    except Exception as e:
        print(f"Error getting document libraries: {e}")
        return None

def get_folder_contents(site_guid, drive_id, item_id, token, path=""):
    """Recursively get all contents of a folder"""
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_guid}/drives/{drive_id}/items/{item_id}/children"
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        all_items = []
        while endpoint:
            response = requests.get(endpoint, headers=headers)
            response.raise_for_status()
            data = response.json()
            
            for item in data.get('value', []):
                item_info = {
                    'name': item['name'],
                    'id': item['id'],
                    'type': 'folder' if 'folder' in item else 'file',
                    'path': f"{path}/{item['name']}",
                    'created': item.get('createdDateTime'),
                    'lastModified': item.get('lastModifiedDateTime'),
                    'createdBy': item.get('createdBy', {}).get('user', {}).get('displayName'),
                    'size': item.get('size', 0),
                    'webUrl': item.get('webUrl')
                }
                
                all_items.append(item_info)
                
                if item_info['type'] == 'folder':
                    child_items = get_folder_contents(
                        site_guid, 
                        drive_id, 
                        item['id'], 
                        token, 
                        item_info['path']
                    )
                    all_items.extend(child_items)
            
            endpoint = data.get('@odata.nextLink')
        
        return all_items
        
    except Exception as e:
        print(f"Error getting folder contents: {e}")
        return []

def get_item_permissions(site_guid, drive_id, item_id, token):
    """Get permissions for a specific item"""
    endpoint = f"https://graph.microsoft.com/v1.0/sites/{site_guid}/drives/{drive_id}/items/{item_id}/permissions"
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        # Add small delay to prevent rate limiting
        time.sleep(REQUEST_DELAY)
        
        response = requests.get(endpoint, headers=headers)
        
        # Check for rate limiting
        if response.status_code == 429:
            retry_after = int(response.headers.get('Retry-After', 5))
            print(f"Rate limited. Waiting for {retry_after} seconds...")
            time.sleep(retry_after)
            return get_item_permissions(site_guid, drive_id, item_id, token)
            
        response.raise_for_status()
        data = response.json()
        
        permissions = []
        for perm in data.get('value', []):
            permission_info = {
                'id': perm.get('id'),
                'roles': ', '.join(perm.get('roles', [])),
                'inherited': 'inheritedFrom' in perm,
                'granted_to_type': None,
                'email': None,
                'display_name': None
            }
            
            # Only use grantedToV2
            granted_to = perm.get('grantedToV2')
            
            if granted_to and 'user' in granted_to:
                permission_info.update({
                    'granted_to_type': 'user',
                    'email': granted_to['user'].get('email'),
                    'display_name': granted_to['user'].get('displayName')
                })
            elif granted_to and 'siteUser' in granted_to:
                permission_info.update({
                    'granted_to_type': 'user',
                    'email': granted_to['siteUser'].get('email'),
                    'display_name': granted_to['siteUser'].get('displayName')
                })
            elif granted_to and 'group' in granted_to:
                permission_info.update({
                    'granted_to_type': 'group',
                    'email': granted_to['group'].get('email'),
                    'display_name': granted_to['group'].get('displayName')
                })
            elif granted_to and 'siteGroup' in granted_to:
                permission_info.update({
                    'granted_to_type': 'group',
                    'email': granted_to['siteGroup'].get('email'),
                    'display_name': granted_to['siteGroup'].get('displayName')
                })
            
            permissions.append(permission_info)
        
        return permissions
        
    except Exception as e:
        print(f"Error getting permissions for item {item_id}: {e}")
        return []

def process_item_permissions(site_guid, drive_id, item, token):
    """Process permissions for a single item (file or folder)"""
    item['permissions'] = get_item_permissions(site_guid, drive_id, item['id'], token)
    return item

def process_site(site_name, token):
    """Process a single SharePoint site"""
    print(f"\nProcessing site: {site_name}")
    
    # Get site GUID
    site_info = get_site_guid(site_name, token)
    if not site_info:
        return None
    
    print(f"Found site: {site_info['display_name']}")
    print(f"Site GUID: {site_info['site_guid']}")
    
    # Get document libraries
    libraries = get_document_libraries(site_info['site_guid'], token)
    if not libraries:
        print("No document libraries found")
        return None
    
    print(f"\nFound {len(libraries)} document libraries:")
    
    # Process each library
    all_items = []
    for lib in libraries:
        print(f"\nProcessing library: {lib['name']}")
        print(f"Drive ID: {lib['id']}")
        
        # Get all items recursively
        items = get_folder_contents(
            site_info['site_guid'], 
            lib['id'], 
            'root', 
            token,
            f"/{lib['name']}"
        )
        
        print(f"Found {len(items)} items (files and folders) in this library")
        
        # Process permissions in parallel with rate limiting
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = []
            for item in items:
                futures.append(executor.submit(
                    process_item_permissions,
                    site_info['site_guid'],
                    lib['id'],
                    item,
                    token
                ))
            
            # Wait for all futures to complete
            for future in as_completed(futures):
                try:
                    processed_item = future.result()
                    all_items.append(processed_item)
                except Exception as e:
                    print(f"Error processing item permissions: {e}")
    
    return {
        'site_info': site_info,
        'libraries': libraries,
        'items': all_items
    }

def generate_report(site_data, output_format='csv'):
    """Generate comprehensive report of all files and folders with permissions"""
    if not site_data or not site_data.get('items'):
        print("No data to generate report")
        return
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    site_name = site_data['site_info']['site_name']
    filename = f"{site_name}_full_permissions_report_{timestamp}.{output_format}"
    
    if output_format == 'csv':
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'site_name', 'library_name', 'item_type', 'path', 'name', 
                'size', 'created', 'last_modified', 'created_by', 'web_url',
                'item_id', 'api_endpoint',
                'unique_permissions', 'permission_owners', 'permission_writers', 
                'permission_readers', 'all_permissions'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for item in site_data['items']:
                # Extract library info from the item's path
                path_parts = item['path'].split('/')
                library_name = path_parts[1] if len(path_parts) > 1 else 'Unknown'
                
                # Find the matching library to get the drive ID
                library = next(
                    (lib for lib in site_data['libraries'] 
                     if lib['name'] == library_name),
                    {'id': 'Unknown', 'name': library_name}
                )
                
                # Build the exact API endpoint format you need
                api_endpoint = (
                    f"https://graph.microsoft.com/v1.0/sites/"
                    f"{site_data['site_info']['site_guid']}/"
                    f"drives/{library['id']}/"
                    f"items/{item['id']}"
                )
                
                # Organize permissions by role
                owners = []
                writers = []
                readers = []
                all_perms = []
                unique_perms = False
                
                for perm in item.get('permissions', []):
                    perm_str = f"{perm['display_name']} ({perm['email'] or perm['granted_to_type']}) - {perm['roles']}"
                    all_perms.append(perm_str)
                    
                    if not perm['inherited']:
                        unique_perms = True
                    
                    if 'owner' in perm['roles'].lower():
                        owners.append(f"{perm['display_name']} ({perm['email']})")
                    elif 'write' in perm['roles'].lower() or 'edit' in perm['roles'].lower():
                        writers.append(f"{perm['display_name']} ({perm['email']})")
                    elif 'read' in perm['roles'].lower():
                        readers.append(f"{perm['display_name']} ({perm['email']})")
                
                writer.writerow({
                    'site_name': site_data['site_info']['display_name'],
                    'library_name': library['name'],
                    'item_type': item['type'],
                    'path': item['path'],
                    'name': item['name'],
                    'size': item['size'],
                    'created': item['created'],
                    'last_modified': item['lastModified'],
                    'created_by': item['createdBy'],
                    'web_url': item['webUrl'],
                    'item_id': item['id'],
                    'api_endpoint': api_endpoint,
                    'unique_permissions': 'Yes' if unique_perms else 'No',
                    'permission_owners': ', '.join(owners),
                    'permission_writers': ', '.join(writers),
                    'permission_readers': ', '.join(readers),
                    'all_permissions': ' | '.join(all_perms)
                })
        
        print(f"\nComprehensive report generated: {filename}")
    else:
        print("Only CSV output format is currently supported")

def main():
    # Step 1: Get access token using certificate authentication
    print("Authenticating with Microsoft Graph using certificate...")
    token = get_access_token()
    if not token:
        print("Failed to authenticate. Please check your certificate and configuration.")
        return
    
    # Step 2: Process each site
    for site_name in SITE_NAMES:
        site_data = process_site(site_name, token)
        if site_data:
            generate_report(site_data)

if __name__ == "__main__":
    main()
