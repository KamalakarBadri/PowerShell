import csv
import json
import uuid
import base64
import time
import requests
from datetime import datetime
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
import sys

class SharePointTokenManager:
    """Manages SharePoint token with automatic renewal"""
    
    def __init__(self, certificate, private_key, tenant_name, app_id, sharepoint_admin_url):
        self.certificate = certificate
        self.private_key = private_key
        self.tenant_name = tenant_name
        self.app_id = app_id
        self.sharepoint_admin_url = sharepoint_admin_url
        self.token = None
        self.token_expiry_time = 0
        self.refresh_buffer = 300
        self.token_lock = Lock()
    
    def get_token(self):
        """Get valid token, renew if expired or about to expire"""
        with self.token_lock:
            current_time = time.time()
            
            if not self.token or current_time >= (self.token_expiry_time - self.refresh_buffer):
                self._renew_token()
            
            return self.token
    
    def _renew_token(self):
        """Renew the access token"""
        print(f"  [Token] Renewing access token...")
        
        scope = f"{self.sharepoint_admin_url}/.default"
        jwt = get_jwt_token(self.certificate, self.private_key, self.tenant_name, self.app_id, scope)
        self.token = get_access_token(jwt, self.tenant_name, self.app_id, scope)
        
        self.token_expiry_time = time.time() + 2700
        
        print(f"  [Token] Token renewed, expires at {datetime.fromtimestamp(self.token_expiry_time).strftime('%H:%M:%S')}")

class GraphTokenManager:
    """Manages Microsoft Graph token with automatic renewal"""
    
    def __init__(self, certificate, private_key, tenant_name, app_id):
        self.certificate = certificate
        self.private_key = private_key
        self.tenant_name = tenant_name
        self.app_id = app_id
        self.token = None
        self.token_expiry_time = 0
        self.refresh_buffer = 300
        self.token_lock = Lock()
    
    def get_token(self):
        """Get valid Graph token, renew if expired or about to expire"""
        with self.token_lock:
            current_time = time.time()
            
            if not self.token or current_time >= (self.token_expiry_time - self.refresh_buffer):
                self._renew_token()
            
            return self.token
    
    def _renew_token(self):
        """Renew the Graph access token"""
        print(f"  [Graph Token] Renewing access token...")
        
        scope = "https://graph.microsoft.com/.default"
        jwt = get_jwt_token(self.certificate, self.private_key, self.tenant_name, self.app_id, scope)
        self.token = get_graph_access_token(jwt, self.tenant_name, self.app_id)
        
        self.token_expiry_time = time.time() + 2700
        
        print(f"  [Graph Token] Token renewed, expires at {datetime.fromtimestamp(self.token_expiry_time).strftime('%H:%M:%S')}")

def load_config(config_file="config.json"):
    """Load configuration from JSON file"""
    try:
        with open(config_file, 'r') as f:
            config = json.load(f)
        
        config.setdefault('page_size', 100)
        config.setdefault('max_retries', 3)
        config.setdefault('max_workers', 20)
        config.setdefault('fetch_metadata', True)
        config.setdefault('skip_deleted_metadata', True)
        config.setdefault('check_owner_exists', True)
        config.setdefault('fetch_manager', True)
        
        return config
    except FileNotFoundError:
        print(f"Error: Configuration file '{config_file}' not found.")
        print("Please create a config.json file with the following structure:")
        print("""
{
    "tenant": "yourtenant.onmicrosoft.com",
    "app_id": "your-app-id",
    "cert_path": "cert.pem",
    "key_path": "key.pem",
    "sharepoint_admin_url": "https://yourtenant-admin.sharepoint.com",
    "list_id": "317f59e4-b925-4d1c-884c-c758bf067a6c",
    "page_size": 100,
    "max_retries": 3,
    "max_workers": 20,
    "fetch_metadata": true,
    "skip_deleted_metadata": true,
    "check_owner_exists": true,
    "fetch_manager": true
}
        """)
        raise
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON format in '{config_file}'.")
        raise

def load_certificate_and_key(certificate_path, private_key_path):
    """Load certificate and private key from PEM files"""
    try:
        with open(certificate_path, "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())

        with open(private_key_path, "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())

        return certificate, private_key
    except Exception as e:
        print(f"Error loading certificate or private key: {str(e)}")
        raise

def get_jwt_token(certificate, private_key, tenant_name, app_id, scope):
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
            "aud": f"https://login.microsoftonline.com/{tenant_name}/oauth2/v2.0/token",
            "exp": expiration,
            "iss": app_id,
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": app_id
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

def get_access_token(jwt, tenant_name, app_id, scope):
    """Get SharePoint access token from Microsoft Identity Platform"""
    url = f"https://login.microsoftonline.com/{tenant_name}/oauth2/v2.0/token"
    
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    
    data = {
        "client_id": app_id,
        "client_assertion": jwt,
        "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        "scope": scope,
        "grant_type": "client_credentials"
    }
    
    try:
        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()
        return response.json()["access_token"]
    except Exception as err:
        print(f"Error getting SharePoint access token: {err}")
        raise

def get_graph_access_token(jwt, tenant_name, app_id):
    """Get Microsoft Graph access token"""
    url = f"https://login.microsoftonline.com/{tenant_name}/oauth2/v2.0/token"
    
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    
    data = {
        "client_id": app_id,
        "client_assertion": jwt,
        "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    
    try:
        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()
        return response.json()["access_token"]
    except Exception as err:
        print(f"Error getting Graph access token: {err}")
        raise

def make_sharepoint_request(token_manager, endpoint, max_retries=3):
    """Make SharePoint request with automatic token renewal on 401 error"""
    for attempt in range(max_retries):
        try:
            headers = {
                "Authorization": f"Bearer {token_manager.get_token()}",
                "Accept": "application/json",
                "Content-Type": "application/json"
            }
            
            response = requests.get(endpoint, headers=headers, timeout=30)
            
            if response.status_code == 401:
                print(f"  [Auth] Token expired, renewing... (Attempt {attempt + 1}/{max_retries})")
                token_manager._renew_token()
                continue
            
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.HTTPError as err:
            if response.status_code == 401 and attempt < max_retries - 1:
                continue
            raise
        except requests.exceptions.Timeout:
            if attempt < max_retries - 1:
                time.sleep(1)
                continue
            raise
        except Exception as err:
            if attempt < max_retries - 1:
                time.sleep(1)
                continue
            raise
    
    raise Exception(f"Failed after {max_retries} attempts")

def check_user_in_deleted(graph_token_manager, user_email, max_retries=3):
    """
    Check if a user is in the deleted users list
    """
    if not user_email or user_email == '':
        return None, "No email provided"
    
    try:
        # URL encode the email for the filter
        encoded_email = requests.utils.quote(user_email)
        endpoint = f"https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?$filter=mail eq '{encoded_email}'&$select=id,userPrincipalName,displayName,mail"
        
        for attempt in range(max_retries):
            try:
                headers = {
                    "Authorization": f"Bearer {graph_token_manager.get_token()}",
                    "Accept": "application/json",
                    "Content-Type": "application/json"
                }
                
                response = requests.get(endpoint, headers=headers, timeout=30)
                
                if response.status_code == 401:
                    print(f"  [Graph Auth] Token expired, renewing... (Attempt {attempt + 1}/{max_retries})")
                    graph_token_manager._renew_token()
                    continue
                
                response.raise_for_status()
                data = response.json()
                
                # Check if we got any results
                value = data.get('value', [])
                if value and len(value) > 0:
                    user = value[0]
                    return {
                        'id': user.get('id', ''),
                        'userPrincipalName': user.get('userPrincipalName', ''),
                        'displayName': user.get('displayName', ''),
                        'mail': user.get('mail', '')
                    }, "User found in deleted items"
                else:
                    return None, "User not in deleted items"
                    
            except requests.exceptions.Timeout:
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None, "Timeout"
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None, f"Error: {str(e)[:100]}"
        
        return None, "Max retries exceeded"
        
    except Exception as e:
        return None, f"Error: {str(e)[:100]}"

def get_user_by_email_filter(graph_token_manager, user_email, max_retries=3):
    """
    Get user by email using Graph API filter: /users?$filter=mail eq 'email'
    Returns user details including UPN and account status
    """
    if not user_email or user_email == '':
        return None, "No email provided"
    
    try:
        # URL encode the email for the filter
        encoded_email = requests.utils.quote(user_email)
        endpoint = f"https://graph.microsoft.com/v1.0/users?$filter=mail eq '{encoded_email}'&$select=id,userPrincipalName,displayName,mail,accountEnabled,userType"
        
        for attempt in range(max_retries):
            try:
                headers = {
                    "Authorization": f"Bearer {graph_token_manager.get_token()}",
                    "Accept": "application/json",
                    "Content-Type": "application/json"
                }
                
                response = requests.get(endpoint, headers=headers, timeout=30)
                
                if response.status_code == 401:
                    print(f"  [Graph Auth] Token expired, renewing... (Attempt {attempt + 1}/{max_retries})")
                    graph_token_manager._renew_token()
                    continue
                
                response.raise_for_status()
                data = response.json()
                
                # Check if we got any results
                value = data.get('value', [])
                if value and len(value) > 0:
                    user = value[0]  # Take the first match
                    upn = user.get('userPrincipalName', '')
                    mail = user.get('mail', user_email)
                    display_name = user.get('displayName', '')
                    account_enabled = user.get('accountEnabled', False)
                    user_type = user.get('userType', '')
                    
                    return {
                        'upn': upn,
                        'mail': mail,
                        'display_name': display_name,
                        'id': user.get('id', ''),
                        'account_enabled': account_enabled,
                        'user_type': user_type,
                        'status': 'Enabled' if account_enabled else 'Disabled'
                    }, f"Found via filter: {display_name} ({upn}) - Status: {'Enabled' if account_enabled else 'Disabled'}"
                else:
                    return None, "User not found with filter"
                    
            except requests.exceptions.Timeout:
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None, "Timeout"
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None, f"Error: {str(e)[:100]}"
        
        return None, "Max retries exceeded"
        
    except Exception as e:
        return None, f"Error: {str(e)[:100]}"

def get_user_manager(graph_token_manager, user_id, max_retries=3):
    """
    Get the manager of a user using Graph API
    """
    if not user_id:
        return None, "No user ID provided"
    
    try:
        endpoint = f"https://graph.microsoft.com/v1.0/users/{user_id}/manager"
        
        for attempt in range(max_retries):
            try:
                headers = {
                    "Authorization": f"Bearer {graph_token_manager.get_token()}",
                    "Accept": "application/json",
                    "Content-Type": "application/json"
                }
                
                response = requests.get(endpoint, headers=headers, timeout=30)
                
                if response.status_code == 401:
                    print(f"  [Graph Auth] Token expired, renewing... (Attempt {attempt + 1}/{max_retries})")
                    graph_token_manager._renew_token()
                    continue
                
                if response.status_code == 200:
                    manager_data = response.json()
                    manager_upn = manager_data.get('userPrincipalName', '')
                    manager_mail = manager_data.get('mail', '')
                    manager_display = manager_data.get('displayName', '')
                    
                    return {
                        'upn': manager_upn,
                        'mail': manager_mail,
                        'display_name': manager_display
                    }, f"Found manager: {manager_display} ({manager_upn})"
                elif response.status_code == 404:
                    return None, "No manager assigned"
                else:
                    if attempt < max_retries - 1:
                        time.sleep(1)
                        continue
                    return None, f"HTTP {response.status_code}"
                    
            except requests.exceptions.Timeout:
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None, "Timeout"
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None, f"Error: {str(e)[:100]}"
        
        return None, "Max retries exceeded"
        
    except Exception as e:
        return None, f"Error: {str(e)[:100]}"

def check_user_and_manager(graph_token_manager, user_email, max_retries=3, fetch_manager=True):
    """
    Check if a user exists, their status (enabled/disabled), and if they're in deleted users
    Uses filter method to find user by email
    """
    if not user_email or user_email == '':
        return {
            'owner_exists': False,
            'owner_status': 'No email provided',
            'owner_upn': '',
            'owner_mail': '',
            'owner_display_name': '',
            'account_status': 'Not Found',
            'user_type': '',
            'is_deleted_user': False,
            'manager_upn': '',
            'manager_mail': '',
            'manager_display_name': '',
            'manager_status': 'No email provided'
        }
    
    # First, try to find user using filter
    user_data, status = get_user_by_email_filter(graph_token_manager, user_email, max_retries)
    
    if not user_data:
        # Try alternative approach: search by userPrincipalName if mail doesn't work
        if '@' in user_email:
            # Try direct lookup by UPN
            try:
                encoded_upn = requests.utils.quote(user_email)
                endpoint = f"https://graph.microsoft.com/v1.0/users/{encoded_upn}"
                
                headers = {
                    "Authorization": f"Bearer {graph_token_manager.get_token()}",
                    "Accept": "application/json",
                    "Content-Type": "application/json"
                }
                
                response = requests.get(endpoint, headers=headers, timeout=30)
                
                if response.status_code == 200:
                    user_data_direct = response.json()
                    user_data = {
                        'upn': user_data_direct.get('userPrincipalName', user_email),
                        'mail': user_data_direct.get('mail', user_email),
                        'display_name': user_data_direct.get('displayName', ''),
                        'id': user_data_direct.get('id', ''),
                        'account_enabled': user_data_direct.get('accountEnabled', False),
                        'user_type': user_data_direct.get('userType', ''),
                        'status': 'Enabled' if user_data_direct.get('accountEnabled', False) else 'Disabled'
                    }
                    status = f"Found via direct UPN lookup: {user_data['display_name']} ({user_data['upn']}) - Status: {user_data['status']}"
                else:
                    # Check if user is in deleted items
                    deleted_user, deleted_status = check_user_in_deleted(graph_token_manager, user_email, max_retries)
                    if deleted_user:
                        return {
                            'owner_exists': False,
                            'owner_status': 'User is in deleted items',
                            'owner_upn': deleted_user.get('userPrincipalName', ''),
                            'owner_mail': deleted_user.get('mail', user_email),
                            'owner_display_name': deleted_user.get('displayName', ''),
                            'account_status': 'Deleted',
                            'user_type': 'Deleted',
                            'is_deleted_user': True,
                            'manager_upn': '',
                            'manager_mail': '',
                            'manager_display_name': '',
                            'manager_status': 'User is deleted'
                        }
                    else:
                        return {
                            'owner_exists': False,
                            'owner_status': f"User not found",
                            'owner_upn': '',
                            'owner_mail': user_email,
                            'owner_display_name': '',
                            'account_status': 'Not Found',
                            'user_type': 'Not Found',
                            'is_deleted_user': False,
                            'manager_upn': '',
                            'manager_mail': '',
                            'manager_display_name': '',
                            'manager_status': 'User not found'
                        }
            except Exception as e:
                # Check if user is in deleted items
                deleted_user, deleted_status = check_user_in_deleted(graph_token_manager, user_email, max_retries)
                if deleted_user:
                    return {
                        'owner_exists': False,
                        'owner_status': 'User is in deleted items',
                        'owner_upn': deleted_user.get('userPrincipalName', ''),
                        'owner_mail': deleted_user.get('mail', user_email),
                        'owner_display_name': deleted_user.get('displayName', ''),
                        'account_status': 'Deleted',
                        'user_type': 'Deleted',
                        'is_deleted_user': True,
                        'manager_upn': '',
                        'manager_mail': '',
                        'manager_display_name': '',
                        'manager_status': 'User is deleted'
                    }
                else:
                    return {
                        'owner_exists': False,
                        'owner_status': f"User not found",
                        'owner_upn': '',
                        'owner_mail': user_email,
                        'owner_display_name': '',
                        'account_status': 'Not Found',
                        'user_type': 'Not Found',
                        'is_deleted_user': False,
                        'manager_upn': '',
                        'manager_mail': '',
                        'manager_display_name': '',
                        'manager_status': 'User not found'
                    }
        else:
            # Check if user is in deleted items
            deleted_user, deleted_status = check_user_in_deleted(graph_token_manager, user_email, max_retries)
            if deleted_user:
                return {
                    'owner_exists': False,
                    'owner_status': 'User is in deleted items',
                    'owner_upn': deleted_user.get('userPrincipalName', ''),
                    'owner_mail': deleted_user.get('mail', user_email),
                    'owner_display_name': deleted_user.get('displayName', ''),
                    'account_status': 'Deleted',
                    'user_type': 'Deleted',
                    'is_deleted_user': True,
                    'manager_upn': '',
                    'manager_mail': '',
                    'manager_display_name': '',
                    'manager_status': 'User is deleted'
                }
            else:
                return {
                    'owner_exists': False,
                    'owner_status': f"User not found",
                    'owner_upn': '',
                    'owner_mail': user_email,
                    'owner_display_name': '',
                    'account_status': 'Not Found',
                    'user_type': 'Not Found',
                    'is_deleted_user': False,
                    'manager_upn': '',
                    'manager_mail': '',
                    'manager_display_name': '',
                    'manager_status': 'User not found'
                }
    
    # User found, now get manager if requested
    manager_data = None
    manager_status = 'Not fetched'
    
    if fetch_manager and user_data.get('id'):
        manager_data, manager_status = get_user_manager(graph_token_manager, user_data['id'], max_retries)
    
    return {
        'owner_exists': True,
        'owner_status': status,
        'owner_upn': user_data.get('upn', ''),
        'owner_mail': user_data.get('mail', user_email),
        'owner_display_name': user_data.get('display_name', ''),
        'account_status': user_data.get('status', 'Unknown'),
        'user_type': user_data.get('user_type', ''),
        'is_deleted_user': False,
        'manager_upn': manager_data.get('upn', '') if manager_data else '',
        'manager_mail': manager_data.get('mail', '') if manager_data else '',
        'manager_display_name': manager_data.get('display_name', '') if manager_data else '',
        'manager_status': manager_status
    }

def is_onedrive_site(site_url):
    """
    Identify OneDrive sites specifically by the URL pattern:
    my.sharepoint.com/personal
    """
    if not site_url:
        return False
    
    site_url_lower = site_url.lower()
    
    # Check for the specific OneDrive URL pattern
    if 'my.sharepoint.com/personal' in site_url_lower:
        return True
    
    # Also check for alternative patterns just in case
    if 'my.sharepoint.com' in site_url_lower and '/personal/' in site_url_lower:
        return True
    
    return False

def should_include_site(site_url, config):
    """Determine if a site should be included based on URL pattern"""
    # Only include sites that match the OneDrive URL pattern
    return is_onedrive_site(site_url)

def get_site_metadata_parallel(token_manager, site_url, max_retries=3):
    """Get site metadata with timeout and retry"""
    try:
        endpoint = f"{site_url.rstrip('/')}/_api/web?$select=LastItemModifiedDate,LastItemUserModifiedDate"
        data = make_sharepoint_request(token_manager, endpoint, max_retries=2)
        
        return {
            'last_item_modified_date': data.get('LastItemModifiedDate', ''),
            'last_item_user_modified_date': data.get('LastItemUserModifiedDate', ''),
            'error': None
        }
    except Exception as e:
        return {
            'last_item_modified_date': 'Error',
            'last_item_user_modified_date': 'Error',
            'error': str(e)[:100]
        }

def check_owner_and_manager(graph_token_manager, user_email, max_retries=3, fetch_manager=True):
    """Wrapper function for parallel owner and manager checking"""
    result = check_user_and_manager(graph_token_manager, user_email, max_retries, fetch_manager)
    return result

def get_all_sites_from_list_optimized(token_manager, graph_token_manager, sharepoint_admin_url, list_id, page_size=100, max_workers=20, fetch_metadata=True, config=None):
    """Get OneDrive sites with metadata and owner validation"""
    print(f"\n=== Retrieving OneDrive Sites from Admin List ===")
    print(f"Filtering for OneDrive sites with URL pattern: my.sharepoint.com/personal")
    
    all_sites = []
    active_sites = []
    deleted_sites = []
    skipped_sites = 0
    
    skip_deleted_metadata = config.get('skip_deleted_metadata', True)
    check_owner = config.get('check_owner_exists', True)
    fetch_manager = config.get('fetch_manager', True)
    
    base_endpoint = f"{sharepoint_admin_url}/_api/Web/Lists(guid'{list_id}')/items"
    endpoint = f"{base_endpoint}?$top={page_size}"
    batch_count = 0
    total_sites = 0
    
    print("Fetching site list...")
    while endpoint:
        batch_count += 1
        try:
            print(f"  Processing batch {batch_count}...")
            data = make_sharepoint_request(token_manager, endpoint)
            
            current_batch = data.get('value', [])
            
            if not current_batch:
                break
            
            print(f"  Found {len(current_batch)} sites in this batch")
            
            for item in current_batch:
                total_sites += 1
                
                site_url = item.get('SiteUrl', '')
                template_name = item.get('TemplateName', '')
                title = item.get('Title', '')
                time_deleted = item.get('TimeDeleted', '')
                created_by_email = item.get('CreatedByEmail', '')
                
                # Check if this is a OneDrive site by URL pattern
                if should_include_site(site_url, config):
                    is_deleted = bool(time_deleted)
                    
                    site_info = {
                        'id': item.get('Id', ''),
                        'time_deleted': time_deleted,
                        'title': title,
                        'site_url': site_url,
                        'site_id': item.get('SiteId', ''),
                        'template_name': template_name,
                        'site_type': 'OneDrive',
                        'is_deleted': is_deleted,
                        'owner_email': created_by_email,
                        'owner_exists': 'Unknown',
                        'owner_status': 'Not checked',
                        'owner_upn': '',
                        'owner_mail': '',
                        'owner_display_name': '',
                        'account_status': 'Not Checked',
                        'user_type': '',
                        'is_deleted_user': False,
                        'manager_upn': '',
                        'manager_mail': '',
                        'manager_display_name': '',
                        'manager_status': 'Not fetched',
                        'storage_quota_bytes': item.get('StorageQuota', 0),
                        'storage_quota_gb': round(item.get('StorageQuota', 0) / (1024**3), 2) if item.get('StorageQuota') else 0,
                        'storage_used_bytes': item.get('StorageUsed', 0),
                        'storage_used_gb': round(item.get('StorageUsed', 0) / (1024**3), 2) if item.get('StorageUsed') else 0,
                        'storage_used_percentage': float(item.get('StorageUsedPercentage', '0')) * 100 if item.get('StorageUsedPercentage') else 0,
                        'created': item.get('Created', ''),
                        'created_by': item.get('CreatedBy', ''),
                        'created_by_email': created_by_email,
                        'modified': item.get('Modified', ''),
                        'last_activity': item.get('LastActivityOn', ''),
                        'num_of_files': item.get('NumOfFiles', 0),
                        'state': item.get('State', 0),
                        'time_created': item.get('TimeCreated', ''),
                        'archive_status': item.get('ArchiveStatus', ''),
                        'last_item_modified_date': 'Skipped (Deleted)' if is_deleted and skip_deleted_metadata else '',
                        'last_item_user_modified_date': 'Skipped (Deleted)' if is_deleted and skip_deleted_metadata else ''
                    }
                    
                    all_sites.append(site_info)
                    
                    if is_deleted:
                        deleted_sites.append(site_info)
                        if skip_deleted_metadata:
                            print(f"  Info: Skipping metadata for deleted site: {title}")
                    else:
                        active_sites.append(site_info)
                else:
                    skipped_sites += 1
                    if skipped_sites % 100 == 0:
                        print(f"  Skipped {skipped_sites} non-OneDrive sites...")
            
            endpoint = data.get('odata.nextLink')
            if endpoint:
                print(f"  Next page available")
            else:
                print("  No more pages")
                
        except Exception as e:
            print(f"Error processing batch {batch_count}: {str(e)}")
            break
    
    print(f"\nTotal sites processed: {total_sites}")
    print(f"  - OneDrive sites found: {len(all_sites)}")
    print(f"    * Active OneDrive sites: {len(active_sites)}")
    print(f"    * Soft-deleted OneDrive sites: {len(deleted_sites)}")
    print(f"  - Non-OneDrive sites skipped: {skipped_sites}")
    
    # Fetch metadata for active sites
    if fetch_metadata and active_sites:
        print(f"\n=== Fetching Metadata for Active OneDrive Sites ===")
        print(f"Fetching metadata for {len(active_sites)} active OneDrive sites using {max_workers} parallel workers...")
        if skip_deleted_metadata:
            print(f"  (Metadata skipped for {len(deleted_sites)} soft-deleted sites)")
        
        processed = 0
        errors = 0
        
        start_time = time.time()
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_site = {
                executor.submit(get_site_metadata_parallel, token_manager, site['site_url']): site
                for site in active_sites if site['site_url']
            }
            
            for future in as_completed(future_to_site):
                site = future_to_site[future]
                try:
                    metadata = future.result(timeout=30)
                    
                    site['last_item_modified_date'] = metadata.get('last_item_modified_date', '')
                    site['last_item_user_modified_date'] = metadata.get('last_item_user_modified_date', '')
                    
                    processed += 1
                    if metadata.get('error'):
                        errors += 1
                    
                    if processed % 10 == 0 or processed == 1:
                        print(f"  Progress: {processed}/{len(active_sites)} active sites processed")
                        
                except Exception as e:
                    errors += 1
                    site['last_item_modified_date'] = 'Error'
                    site['last_item_user_modified_date'] = 'Error'
                    print(f"  Warning: Error processing {site.get('title', 'Unknown')}: {str(e)[:50]}")
        
        elapsed = time.time() - start_time
        print(f"\nMetadata fetching completed in {elapsed:.2f} seconds")
        print(f"  Successfully processed: {processed - errors} active sites")
        if errors > 0:
            print(f"  Errors: {errors} sites")
    
    # Check owner exists for active sites
    if check_owner and active_sites:
        print(f"\n=== Checking Owner Status for Active OneDrive Sites ===")
        print(f"Checking owner existence for {len(active_sites)} active OneDrive sites using Graph API...")
        print(f"  - Using filter: /users?$filter=mail eq 'email'")
        print(f"  - Checking if user is Enabled, Disabled, or in Deleted Users")
        if fetch_manager:
            print(f"  - Fetching manager for each user")
        
        processed = 0
        owner_errors = 0
        orphaned_count = 0
        found_count = 0
        enabled_count = 0
        disabled_count = 0
        deleted_user_count = 0
        manager_found_count = 0
        
        start_time = time.time()
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_site = {
                executor.submit(check_owner_and_manager, graph_token_manager, site.get('created_by_email', ''), 3, fetch_manager): site
                for site in active_sites if site.get('created_by_email')
            }
            
            for site in active_sites:
                if not site.get('created_by_email'):
                    site['owner_exists'] = False
                    site['owner_status'] = 'No email provided'
                    site['owner_upn'] = ''
                    site['owner_mail'] = ''
                    site['owner_display_name'] = ''
                    site['account_status'] = 'Not Found'
                    site['user_type'] = ''
                    site['is_deleted_user'] = False
                    site['manager_upn'] = ''
                    site['manager_mail'] = ''
                    site['manager_display_name'] = ''
                    site['manager_status'] = 'No email provided'
                    orphaned_count += 1
            
            for future in as_completed(future_to_site):
                site = future_to_site[future]
                try:
                    result = future.result(timeout=30)
                    site['owner_exists'] = result.get('owner_exists', False)
                    site['owner_status'] = result.get('owner_status', 'Unknown')
                    site['owner_upn'] = result.get('owner_upn', '')
                    site['owner_mail'] = result.get('owner_mail', '')
                    site['owner_display_name'] = result.get('owner_display_name', '')
                    site['account_status'] = result.get('account_status', 'Not Found')
                    site['user_type'] = result.get('user_type', '')
                    site['is_deleted_user'] = result.get('is_deleted_user', False)
                    site['manager_upn'] = result.get('manager_upn', '')
                    site['manager_mail'] = result.get('manager_mail', '')
                    site['manager_display_name'] = result.get('manager_display_name', '')
                    site['manager_status'] = result.get('manager_status', 'Not fetched')
                    
                    processed += 1
                    if result.get('owner_exists', False):
                        found_count += 1
                        if result.get('account_status') == 'Enabled':
                            enabled_count += 1
                        elif result.get('account_status') == 'Disabled':
                            disabled_count += 1
                        elif result.get('is_deleted_user', False):
                            deleted_user_count += 1
                        
                        if fetch_manager and result.get('manager_upn'):
                            manager_found_count += 1
                    else:
                        orphaned_count += 1
                        if result.get('is_deleted_user', False):
                            deleted_user_count += 1
                        owner_errors += 1 if 'Error' in result.get('owner_status', '') else 0
                    
                    if processed % 10 == 0 or processed == 1:
                        print(f"  Progress: {processed}/{len(active_sites)} owner checks completed")
                        print(f"    Users found: {found_count}, Orphaned: {orphaned_count}")
                        print(f"    Enabled: {enabled_count}, Disabled: {disabled_count}, Deleted: {deleted_user_count}")
                        if fetch_manager:
                            print(f"    Managers found: {manager_found_count}")
                        
                except Exception as e:
                    owner_errors += 1
                    site['owner_exists'] = False
                    site['owner_status'] = f'Error: {str(e)[:50]}'
                    site['owner_upn'] = ''
                    site['owner_mail'] = ''
                    site['owner_display_name'] = ''
                    site['account_status'] = 'Error'
                    site['user_type'] = ''
                    site['is_deleted_user'] = False
                    site['manager_upn'] = ''
                    site['manager_mail'] = ''
                    site['manager_display_name'] = ''
                    site['manager_status'] = f'Error: {str(e)[:50]}'
                    orphaned_count += 1
                    print(f"  Warning: Error checking owner for {site.get('title', 'Unknown')}: {str(e)[:50]}")
        
        elapsed = time.time() - start_time
        print(f"\nOwner checking completed in {elapsed:.2f} seconds")
        print(f"  Total active sites checked: {processed + len([s for s in active_sites if not s.get('created_by_email')])}")
        print(f"  Users found: {found_count}")
        print(f"    - Enabled: {enabled_count}")
        print(f"    - Disabled: {disabled_count}")
        print(f"    - Deleted users: {deleted_user_count}")
        print(f"  Orphaned sites found: {orphaned_count}")
        if fetch_manager:
            print(f"  Managers found: {manager_found_count}")
        if owner_errors > 0:
            print(f"  Errors: {owner_errors}")
    
    if deleted_sites:
        print(f"\nFound {len(deleted_sites)} soft-deleted OneDrive sites")
        print(f"   (Owner status and metadata were {'skipped' if skip_deleted_metadata else 'fetched'} for deleted sites)")
    
    return all_sites

def save_to_csv(sites, filename):
    """Save sites data to CSV file"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'Site URL',
                'User UPN',
                'User Email',
                'User Display Name',
                'User Account Status',
                'User Type',
                'Is Deleted User',
                'Manager UPN',
                'Manager Email',
                'Manager Display Name',
                'Created On',
                'Deleted On',
                'Archive Status',
                'Title',
                'Site ID',
                'Template Name',
                'Site Type',
                'Is Deleted',
                'Owner Exists',
                'Owner Status',
                'Manager Status',
                'Storage Used (GB)',
                'Storage Quota (GB)',
                'Storage Used (%)',
                'Created By',
                'Created By Email',
                'Modified',
                'Last Activity',
                'Number of Files',
                'State',
                'Time Created',
                'Last Item Modified Date',
                'Last Item User Modified Date'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            
            for site in sites:
                writer.writerow({
                    'Site URL': site.get('site_url', ''),
                    'User UPN': site.get('owner_upn', ''),
                    'User Email': site.get('owner_mail', site.get('owner_email', '')),
                    'User Display Name': site.get('owner_display_name', ''),
                    'User Account Status': site.get('account_status', 'Not Checked'),
                    'User Type': site.get('user_type', ''),
                    'Is Deleted User': 'Yes' if site.get('is_deleted_user', False) else 'No',
                    'Manager UPN': site.get('manager_upn', ''),
                    'Manager Email': site.get('manager_mail', ''),
                    'Manager Display Name': site.get('manager_display_name', ''),
                    'Created On': site.get('time_created', ''),
                    'Deleted On': site.get('time_deleted', ''),
                    'Archive Status': site.get('archive_status', ''),
                    'Title': site.get('title', ''),
                    'Site ID': site.get('site_id', ''),
                    'Template Name': site.get('template_name', ''),
                    'Site Type': site.get('site_type', 'OneDrive'),
                    'Is Deleted': 'Yes' if site.get('is_deleted', False) else 'No',
                    'Owner Exists': 'Yes' if site.get('owner_exists', False) else 'No',
                    'Owner Status': site.get('owner_status', 'Not checked'),
                    'Manager Status': site.get('manager_status', 'Not fetched'),
                    'Storage Used (GB)': site.get('storage_used_gb', 0),
                    'Storage Quota (GB)': site.get('storage_quota_gb', 0),
                    'Storage Used (%)': round(site.get('storage_used_percentage', 0), 4),
                    'Created By': site.get('created_by', ''),
                    'Created By Email': site.get('created_by_email', ''),
                    'Modified': site.get('modified', ''),
                    'Last Activity': site.get('last_activity', ''),
                    'Number of Files': site.get('num_of_files', 0),
                    'State': site.get('state', 0),
                    'Time Created': site.get('time_created', ''),
                    'Last Item Modified Date': site.get('last_item_modified_date', ''),
                    'Last Item User Modified Date': site.get('last_item_user_modified_date', '')
                })
        
        print(f"\nCSV report saved to: {filename}")
        
    except Exception as e:
        print(f"Error saving CSV: {str(e)}")

def generate_filename(tenant_name, site_type="onedrive"):
    """Generate filename with tenant name, site type, and current timestamp"""
    tenant_clean = tenant_name.split('.')[0] if '.' in tenant_name else tenant_name
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{tenant_clean}_{site_type}_sites_report_{timestamp}.csv"

def main():
    # Load configuration
    config = load_config("config.json")
    
    tenant_name = config.get('tenant')
    app_id = config.get('app_id')
    certificate_path = config.get('cert_path')
    private_key_path = config.get('key_path')
    sharepoint_admin_url = config.get('sharepoint_admin_url')
    list_id = config.get('list_id')
    page_size = config.get('page_size', 100)
    max_retries = config.get('max_retries', 3)
    max_workers = config.get('max_workers', 20)
    fetch_metadata = config.get('fetch_metadata', True)
    skip_deleted_metadata = config.get('skip_deleted_metadata', True)
    check_owner = config.get('check_owner_exists', True)
    fetch_manager = config.get('fetch_manager', True)
    
    print(f"Configuration loaded:")
    print(f"  Tenant: {tenant_name}")
    print(f"  SharePoint Admin URL: {sharepoint_admin_url}")
    print(f"  List ID: {list_id}")
    print(f"  Page Size: {page_size}")
    print(f"  Max Retries: {max_retries}")
    print(f"  Max Workers: {max_workers}")
    print(f"  Fetch Metadata: {fetch_metadata}")
    print(f"  Skip Deleted Site Metadata: {skip_deleted_metadata}")
    print(f"  Check Owner Exists: {check_owner}")
    print(f"  Fetch Manager: {fetch_manager}")
    
    if not sharepoint_admin_url:
        print("Error: sharepoint_admin_url is required in config.json")
        return
    if not list_id:
        print("Error: list_id is required in config.json")
        return
    
    try:
        certificate, private_key = load_certificate_and_key(certificate_path, private_key_path)
        print("Certificate and private key loaded successfully")
        
        sharepoint_token_manager = SharePointTokenManager(certificate, private_key, tenant_name, app_id, sharepoint_admin_url)
        graph_token_manager = GraphTokenManager(certificate, private_key, tenant_name, app_id)
        
        initial_token = sharepoint_token_manager.get_token()
        print("SharePoint access token retrieved successfully")
        print(f"  Token expires at: {datetime.fromtimestamp(sharepoint_token_manager.token_expiry_time).strftime('%H:%M:%S')}")
        
        if check_owner:
            initial_graph_token = graph_token_manager.get_token()
            print("Graph access token retrieved successfully")
            print(f"  Token expires at: {datetime.fromtimestamp(graph_token_manager.token_expiry_time).strftime('%H:%M:%S')}")
        
        onedrive_sites = get_all_sites_from_list_optimized(
            sharepoint_token_manager,
            graph_token_manager,
            sharepoint_admin_url,
            list_id,
            page_size,
            max_workers,
            fetch_metadata,
            config
        )
        
        if not onedrive_sites:
            print("\nNo OneDrive sites found!")
            return
        
        filename = generate_filename(tenant_name, "onedrive")
        save_to_csv(onedrive_sites, filename)
        
        # Print summary
        active_sites = [s for s in onedrive_sites if not s.get('is_deleted', False)]
        deleted_sites = [s for s in onedrive_sites if s.get('is_deleted', False)]
        
        total_storage = sum(s['storage_used_gb'] for s in onedrive_sites)
        total_quota = sum(s['storage_quota_gb'] for s in onedrive_sites)
        total_files = sum(s['num_of_files'] for s in onedrive_sites)
        
        orphaned_sites = [s for s in active_sites if not s.get('owner_exists', True)]
        found_users = [s for s in active_sites if s.get('owner_exists', False)]
        enabled_users = [s for s in active_sites if s.get('account_status') == 'Enabled']
        disabled_users = [s for s in active_sites if s.get('account_status') == 'Disabled']
        deleted_users = [s for s in active_sites if s.get('is_deleted_user', False)]
        manager_found = [s for s in active_sites if s.get('manager_upn', '')]
        
        print(f"\n{'='*50}")
        print(f"ONEDRIVE SITES SUMMARY")
        print(f"{'='*50}")
        print(f"Total OneDrive Sites: {len(onedrive_sites)}")
        print(f"  Active sites: {len(active_sites)}")
        print(f"  Soft-deleted sites: {len(deleted_sites)}")
        
        if check_owner:
            print(f"\nOwner Status (Active Sites Only):")
            print(f"  Total active sites: {len(active_sites)}")
            print(f"  Users found: {len(found_users)}")
            print(f"    - Enabled: {len(enabled_users)}")
            print(f"    - Disabled: {len(disabled_users)}")
            print(f"    - Deleted users: {len(deleted_users)}")
            print(f"  Orphaned sites (owner not found): {len(orphaned_sites)}")
            if len(active_sites) > 0:
                print(f"  Percentage orphaned: {(len(orphaned_sites)/len(active_sites)*100):.2f}%")
            
            if fetch_manager:
                print(f"\nManager Status (Active Sites Only):")
                print(f"  Managers found: {len(manager_found)}")
                if len(found_users) > 0:
                    print(f"  Percentage with managers: {(len(manager_found)/len(found_users)*100):.2f}%")
        
        print(f"\nStorage Usage:")
        print(f"  Total Storage Used: {total_storage:.2f} GB")
        print(f"  Total Storage Quota: {total_quota:.2f} GB")
        if total_quota > 0:
            print(f"  Overall Usage: {(total_storage / total_quota) * 100:.2f}%")
        print(f"  Total Files: {total_files:,}")
        
        # Top 5 largest sites
        largest_sites = sorted(onedrive_sites, key=lambda x: x['storage_used_gb'], reverse=True)[:5]
        if largest_sites:
            print(f"\nTop 5 Largest OneDrive Sites by Storage:")
            for i, site in enumerate(largest_sites, 1):
                status = "[DELETED]" if site.get('is_deleted') else "[ACTIVE]"
                owner_status = "[ORPHANED]" if not site.get('owner_exists', True) and not site.get('is_deleted') else ""
                user_status = f"[{site.get('account_status', 'Unknown')}]" if site.get('account_status') and site.get('account_status') != 'Not Checked' else ""
                manager_info = f"Manager: {site.get('manager_upn', 'N/A')}" if site.get('manager_upn') else ""
                print(f"  {i}. {status} {site['title']}: {site['storage_used_gb']:.2f} GB {owner_status} {user_status} {manager_info}")
        
        # Orphaned sites
        if check_owner and orphaned_sites:
            print(f"\nOrphaned OneDrive Sites (Owner Not Found):")
            for site in orphaned_sites[:5]:
                print(f"  - {site['title']} ({site['owner_email']}) - {site['owner_status']}")
            if len(orphaned_sites) > 5:
                print(f"  ... and {len(orphaned_sites) - 5} more")
            
            orphaned_storage = sum(s['storage_used_gb'] for s in orphaned_sites)
            if orphaned_storage > 0:
                print(f"  Storage used by orphaned sites: {orphaned_storage:.2f} GB")
        
        # Disabled users
        if check_owner and disabled_users:
            print(f"\nDisabled User Accounts ({len(disabled_users)} total):")
            for site in disabled_users[:5]:
                print(f"  - {site['title']} - Owner: {site.get('owner_upn', site.get('owner_email', 'Unknown'))}")
            if len(disabled_users) > 5:
                print(f"  ... and {len(disabled_users) - 5} more")
            
            disabled_storage = sum(s['storage_used_gb'] for s in disabled_users)
            if disabled_storage > 0:
                print(f"  Storage used by disabled accounts: {disabled_storage:.2f} GB")
        
        # Deleted users
        if check_owner and deleted_users:
            print(f"\nDeleted User Accounts ({len(deleted_users)} total):")
            for site in deleted_users[:5]:
                print(f"  - {site['title']} - Owner: {site.get('owner_upn', site.get('owner_email', 'Unknown'))}")
            if len(deleted_users) > 5:
                print(f"  ... and {len(deleted_users) - 5} more")
            
            deleted_user_storage = sum(s['storage_used_gb'] for s in deleted_users)
            if deleted_user_storage > 0:
                print(f"  Storage used by deleted accounts: {deleted_user_storage:.2f} GB")
        
        # Deleted sites
        if deleted_sites:
            print(f"\nSoft-Deleted OneDrive Sites ({len(deleted_sites)} total):")
            for site in deleted_sites[:5]:
                print(f"  - {site['title']} - Deleted: {site.get('time_deleted', 'Unknown')}")
            if len(deleted_sites) > 5:
                print(f"  ... and {len(deleted_sites) - 5} more")
            
            deleted_storage = sum(s['storage_used_gb'] for s in deleted_sites)
            if deleted_storage > 0:
                print(f"  Storage used by deleted sites: {deleted_storage:.2f} GB")
        
        print(f"\n{'='*50}")
        print(f"Script completed successfully!")
        print(f"Report saved as: {filename}")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main()
