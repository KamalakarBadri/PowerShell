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

# Updated API endpoints for nested sites structure
# Format: sites/mainsiteguid/sites/subsiteguid/drives/drive_id
DRIVE_ENDPOINTS = [
    {
        "name": "Main_Site_Subsite_Documents",
        "main_site_guid": "your-main-site-guid",
        "sub_site_guid": "your-sub-site-guid", 
        "drive_id": "your-drive-id"
    },
    {
        "name": "Another_Main_Site_Subsite_Library",
        "main_site_guid": "another-main-site-guid",
        "sub_site_guid": "another-sub-site-guid",
        "drive_id": "another-drive-id"
    }
]

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

def get_drive_info(drive_config, token):
    """Get drive information from the nested site structure"""
    # Build the endpoint for nested site structure
    drive_endpoint = (
        f"https://graph.microsoft.com/v1.0/sites/{drive_config['main_site_guid']}/"
        f"sites/{drive_config['sub_site_guid']}/drives/{drive_config['drive_id']}"
    )
    
    headers = {"Authorization": f"Bearer {token}"}
    
    try:
        response = requests.get(drive_endpoint, headers=headers)
        response.raise_for_status()
        data = response.json()
        
        return {
            'main_site_guid': drive_config['main_site_guid'],
            'sub_site_guid': drive_config['sub_site_guid'],
            'drive_id': drive_config['drive_id'],
            'drive_name': data.get('name', 'Unknown'),
            'drive_type': data.get('driveType', 'Unknown'),
            'web_url': data.get('webUrl', ''),
            'endpoint': drive_endpoint
        }
        
    except Exception as e:
        print(f"Error getting drive info from {drive_endpoint}: {e}")
        return None

def get_folder_contents(main_site_guid, sub_site_guid, drive_id, item_id, token, path=""):
    """Recursively get all contents of a folder from nested site structure"""
    endpoint = (
        f"https://graph.microsoft.com/v1.0/sites/{main_site_guid}/"
        f"sites/{sub_site_guid}/drives/{drive_id}/items/{item_id}/children"
    )
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
                
                # Recursively get folder contents
                if item_info['type'] == 'folder':
                    child_items = get_folder_contents(
                        main_site_guid,
                        sub_site_guid,
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

def get_item_permissions(main_site_guid, sub_site_guid, drive_id, item_id, token):
    """Get permissions for a specific item from nested site structure"""
    endpoint = (
        f"https://graph.microsoft.com/v1.0/sites/{main_site_guid}/"
        f"sites/{sub_site_guid}/drives/{drive_id}/items/{item_id}/permissions"
    )
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
            return get_item_permissions(main_site_guid, sub_site_guid, drive_id, item_id, token)
            
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
            
            # Check different permission formats
            granted_to = perm.get('grantedToV2') or perm.get('grantedTo')
            
            if granted_to:
                if 'user' in granted_to:
                    permission_info.update({
                        'granted_to_type': 'user',
                        'email': granted_to['user'].get('email'),
                        'display_name': granted_to['user'].get('displayName')
                    })
                elif 'siteUser' in granted_to:
                    permission_info.update({
                        'granted_to_type': 'user',
                        'email': granted_to['siteUser'].get('email'),
                        'display_name': granted_to['siteUser'].get('displayName')
                    })
                elif 'group' in granted_to:
                    permission_info.update({
                        'granted_to_type': 'group',
                        'email': granted_to['group'].get('email'),
                        'display_name': granted_to['group'].get('displayName')
                    })
                elif 'siteGroup' in granted_to:
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

def process_item_permissions(main_site_guid, sub_site_guid, drive_id, item, token):
    """Process permissions for a single item (file or folder)"""
    item['permissions'] = get_item_permissions(main_site_guid, sub_site_guid, drive_id, item['id'], token)
    return item

def process_drive(drive_config, token):
    """Process a single drive using nested site structure"""
    print(f"\nProcessing drive: {drive_config['name']}")
    print(f"Main Site GUID: {drive_config['main_site_guid']}")
    print(f"Sub Site GUID: {drive_config['sub_site_guid']}")
    print(f"Drive ID: {drive_config['drive_id']}")
    
    # Get drive information
    drive_info = get_drive_info(drive_config, token)
    if not drive_info:
        print(f"Failed to get drive information for {drive_config['name']}")
        return None
    
    print(f"Drive Name: {drive_info['drive_name']}")
    print(f"Drive Type: {drive_info['drive_type']}")
    print(f"Endpoint: {drive_info['endpoint']}")
    
    # Get all items in the drive
    print("Getting all items in drive...")
    items = get_folder_contents(
        drive_info['main_site_guid'],
        drive_info['sub_site_guid'],
        drive_info['drive_id'],
        'root',
        token,
        f"/{drive_info['drive_name']}"
    )
    
    print(f"Found {len(items)} items (files and folders)")
    
    if not items:
        print("No items found in this drive")
        return None
    
    # Process permissions in parallel with rate limiting
    print("Processing permissions...")
    all_items = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = []
        for item in items:
            futures.append(executor.submit(
                process_item_permissions,
                drive_info['main_site_guid'],
                drive_info['sub_site_guid'],
                drive_info['drive_id'],
                item,
                token
            ))
        
        # Wait for all futures to complete
        processed_count = 0
        for future in as_completed(futures):
            try:
                processed_item = future.result()
                all_items.append(processed_item)
                processed_count += 1
                if processed_count % 10 == 0:
                    print(f"Processed {processed_count}/{len(items)} items...")
            except Exception as e:
                print(f"Error processing item permissions: {e}")
    
    return {
        'drive_config': drive_config,
        'drive_info': drive_info,
        'items': all_items
    }

def generate_report(drive_data, output_format='csv'):
    """Generate comprehensive report of all files and folders with permissions"""
    if not drive_data or not drive_data.get('items'):
        print("No data to generate report")
        return
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    drive_name = drive_data['drive_config']['name']
    filename = f"{drive_name}_permissions_report_{timestamp}.{output_format}"
    
    if output_format == 'csv':
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'drive_name', 'drive_endpoint', 'main_site_guid', 'sub_site_guid', 'drive_id',
                'item_type', 'path', 'name', 'size', 'created', 'last_modified', 
                'created_by', 'web_url', 'item_id', 'permissions_api_endpoint',
                'unique_permissions', 'permission_owners', 'permission_writers', 
                'permission_readers', 'all_permissions'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for item in drive_data['items']:
                # Build the permissions API endpoint for nested structure
                permissions_endpoint = (
                    f"https://graph.microsoft.com/v1.0/sites/"
                    f"{drive_data['drive_info']['main_site_guid']}/"
                    f"sites/{drive_data['drive_info']['sub_site_guid']}/"
                    f"drives/{drive_data['drive_info']['drive_id']}/"
                    f"items/{item['id']}/permissions"
                )
                
                # Organize permissions by role
                owners = []
                writers = []
                readers = []
                all_perms = []
                unique_perms = False
                
                for perm in item.get('permissions', []):
                    display_name = perm['display_name'] or 'Unknown'
                    email = perm['email'] or perm['granted_to_type'] or 'Unknown'
                    roles = perm['roles'] or 'Unknown'
                    
                    perm_str = f"{display_name} ({email}) - {roles}"
                    all_perms.append(perm_str)
                    
                    if not perm['inherited']:
                        unique_perms = True
                    
                    if roles and 'owner' in roles.lower():
                        owners.append(f"{display_name} ({email})")
                    elif roles and ('write' in roles.lower() or 'edit' in roles.lower()):
                        writers.append(f"{display_name} ({email})")
                    elif roles and 'read' in roles.lower():
                        readers.append(f"{display_name} ({email})")
                
                writer.writerow({
                    'drive_name': drive_data['drive_info']['drive_name'],
                    'drive_endpoint': drive_data['drive_info']['endpoint'],
                    'main_site_guid': drive_data['drive_info']['main_site_guid'],
                    'sub_site_guid': drive_data['drive_info']['sub_site_guid'],
                    'drive_id': drive_data['drive_info']['drive_id'],
                    'item_type': item['type'],
                    'path': item['path'],
                    'name': item['name'],
                    'size': item['size'],
                    'created': item['created'],
                    'last_modified': item['lastModified'],
                    'created_by': item['createdBy'],
                    'web_url': item['webUrl'],
                    'item_id': item['id'],
                    'permissions_api_endpoint': permissions_endpoint,
                    'unique_permissions': 'Yes' if unique_perms else 'No',
                    'permission_owners': ', '.join(owners),
                    'permission_writers': ', '.join(writers),
                    'permission_readers': ', '.join(readers),
                    'all_permissions': ' | '.join(all_perms)
                })
        
        print(f"\nReport generated: {filename}")
        print(f"Total items processed: {len(drive_data['items'])}")
    else:
        print("Only CSV output format is currently supported")

def main():
    # Step 1: Get access token using certificate authentication
    print("Authenticating with Microsoft Graph using certificate...")
    token = get_access_token()
    if not token:
        print("Failed to authenticate. Please check your certificate and configuration.")
        return
    
    # Step 2: Process each drive endpoint
    for drive_config in DRIVE_ENDPOINTS:
        drive_data = process_drive(drive_config, token)
        if drive_data:
            generate_report(drive_data)
        print("-" * 50)

if __name__ == "__main__":
    main()
