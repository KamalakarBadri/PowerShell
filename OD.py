import csv
import json
import uuid
import base64
import time
import requests
from datetime import datetime, timedelta
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
import sys
import os
import re

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
        with self.token_lock:
            current_time = time.time()
            if not self.token or current_time >= (self.token_expiry_time - self.refresh_buffer):
                self._renew_token()
            return self.token
    
    def _renew_token(self):
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
        with self.token_lock:
            current_time = time.time()
            if not self.token or current_time >= (self.token_expiry_time - self.refresh_buffer):
                self._renew_token()
            return self.token
    
    def _renew_token(self):
        print(f"  [Graph Token] Renewing access token...")
        scope = "https://graph.microsoft.com/.default"
        jwt = get_jwt_token(self.certificate, self.private_key, self.tenant_name, self.app_id, scope)
        self.token = get_graph_access_token(jwt, self.tenant_name, self.app_id)
        self.token_expiry_time = time.time() + 2700
        print(f"  [Graph Token] Token renewed, expires at {datetime.fromtimestamp(self.token_expiry_time).strftime('%H:%M:%S')}")

def load_config(config_file="config.json"):
    try:
        with open(config_file, 'r') as f:
            config = json.load(f)
        
        config.setdefault('page_size', 100)
        config.setdefault('max_retries', 3)
        config.setdefault('max_workers', 20)
        config.setdefault('check_owner_exists', True)
        config.setdefault('fetch_manager', True)
        config.setdefault('master_report', True)
        config.setdefault('ignore_url_pattern', r'm_[A-Za-z0-9]+_[A-Za-z0-9]+')
        config.setdefault('log_file', 'change_log.txt')
        config.setdefault('enable_logging', True)
        
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
    "check_owner_exists": true,
    "fetch_manager": true,
    "master_report": true,
    "ignore_url_pattern": "m_[A-Za-z0-9]+_[A-Za-z0-9]+",
    "log_file": "change_log.txt",
    "enable_logging": true
}
        """)
        raise
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON format in '{config_file}'.")
        raise

def load_certificate_and_key(certificate_path, private_key_path):
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
    try:
        now = int(time.time())
        expiration = now + 300
        
        thumbprint = certificate.fingerprint(hashes.SHA1())
        x5t = base64.urlsafe_b64encode(thumbprint).decode('utf-8').replace('=', '')
        
        jwt_header = {"alg": "RS256", "typ": "JWT", "x5t": x5t}
        jwt_payload = {
            "aud": f"https://login.microsoftonline.com/{tenant_name}/oauth2/v2.0/token",
            "exp": expiration,
            "iss": app_id,
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": app_id
        }
        
        encoded_header = base64.urlsafe_b64encode(json.dumps(jwt_header, separators=(',', ':')).encode('utf-8')).decode('utf-8').replace('=', '')
        encoded_payload = base64.urlsafe_b64encode(json.dumps(jwt_payload, separators=(',', ':')).encode('utf-8')).decode('utf-8').replace('=', '')
        jwt_unsigned = f"{encoded_header}.{encoded_payload}"
        
        signature = private_key.sign(jwt_unsigned.encode('utf-8'), padding.PKCS1v15(), hashes.SHA256())
        encoded_signature = base64.urlsafe_b64encode(signature).decode('utf-8').replace('=', '')
        
        return f"{jwt_unsigned}.{encoded_signature}"
    except Exception as e:
        print(f"Error generating JWT: {str(e)}")
        raise

def get_access_token(jwt, tenant_name, app_id, scope):
    url = f"https://login.microsoftonline.com/{tenant_name}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
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
    url = f"https://login.microsoftonline.com/{tenant_name}/oauth2/v2.0/token"
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
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

def get_user_by_email_filter(graph_token_manager, user_email, max_retries=3):
    if not user_email:
        return None, "No email provided"
    
    try:
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
                value = data.get('value', [])
                if value and len(value) > 0:
                    user = value[0]
                    return {
                        'upn': user.get('userPrincipalName', ''),
                        'mail': user.get('mail', user_email),
                        'display_name': user.get('displayName', ''),
                        'id': user.get('id', ''),
                        'account_enabled': user.get('accountEnabled', False),
                        'user_type': user.get('userType', ''),
                        'status': 'Enabled' if user.get('accountEnabled', False) else 'Disabled'
                    }, f"Found: {user.get('displayName', '')} ({user.get('userPrincipalName', '')})"
                else:
                    return None, "User not found"
                    
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

def check_user_in_deleted(graph_token_manager, user_email, max_retries=3):
    if not user_email:
        return None, None, "No email provided"
    
    try:
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
                value = data.get('value', [])
                if value and len(value) > 0:
                    user = value[0]
                    user_id = user.get('id', '')
                    
                    # Get deletion time
                    deleted_time_endpoint = f"https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user/{user_id}/deletedDateTime"
                    time_response = requests.get(deleted_time_endpoint, headers=headers, timeout=30)
                    deleted_time = ""
                    if time_response.status_code == 200:
                        deleted_time = time_response.json().get('value', '')
                    
                    return {
                        'id': user_id,
                        'userPrincipalName': user.get('userPrincipalName', ''),
                        'displayName': user.get('displayName', ''),
                        'mail': user.get('mail', '')
                    }, deleted_time, "Deleted user"
                else:
                    return None, None, "Not in deleted"
                    
            except requests.exceptions.Timeout:
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None, None, "Timeout"
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return None, None, f"Error: {str(e)[:100]}"
        
        return None, None, "Max retries exceeded"
    except Exception as e:
        return None, None, f"Error: {str(e)[:100]}"

def get_user_manager(graph_token_manager, user_id, max_retries=3):
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
                    return {
                        'upn': manager_data.get('userPrincipalName', ''),
                        'mail': manager_data.get('mail', ''),
                        'display_name': manager_data.get('displayName', '')
                    }, "Found"
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

def check_user_status(graph_token_manager, user_email, max_retries=3):
    """Check user status - returns current status and if it's a deleted user"""
    if not user_email:
        return {
            'exists': False,
            'status': 'No email',
            'upn': '',
            'mail': '',
            'display_name': '',
            'is_deleted': False,
            'deleted_time': ''
        }
    
    # Try to find user
    user_data, status = get_user_by_email_filter(graph_token_manager, user_email, max_retries)
    
    if user_data:
        return {
            'exists': True,
            'status': user_data.get('status', 'Unknown'),
            'upn': user_data.get('upn', ''),
            'mail': user_data.get('mail', user_email),
            'display_name': user_data.get('display_name', ''),
            'is_deleted': False,
            'deleted_time': ''
        }
    
    # Check deleted users
    deleted_user, deleted_time, _ = check_user_in_deleted(graph_token_manager, user_email, max_retries)
    if deleted_user:
        return {
            'exists': False,
            'status': 'Deleted',
            'upn': deleted_user.get('userPrincipalName', ''),
            'mail': deleted_user.get('mail', user_email),
            'display_name': deleted_user.get('displayName', ''),
            'is_deleted': True,
            'deleted_time': deleted_time
        }
    
    return {
        'exists': False,
        'status': 'Not Found',
        'upn': '',
        'mail': user_email,
        'display_name': '',
        'is_deleted': False,
        'deleted_time': ''
    }

def should_ignore_url(site_url, config):
    """
    Check if URL should be ignored based on pattern
    Default pattern: m_[A-Za-z0-9]+_[A-Za-z0-9]+ (matches m_XXXXXX_XXX)
    """
    if not site_url:
        return False
    
    ignore_pattern = config.get('ignore_url_pattern', r'm_[A-Za-z0-9]+_[A-Za-z0-9]+')
    
    # Check if the URL contains the pattern
    if re.search(ignore_pattern, site_url, re.IGNORECASE):
        return True
    
    return False

def is_onedrive_site(site_url):
    if not site_url:
        return False
    site_url_lower = site_url.lower()
    return 'my.sharepoint.com/personal' in site_url_lower

def should_include_site(site_url, config):
    """Determine if a site should be included"""
    # First check if URL should be ignored
    if should_ignore_url(site_url, config):
        return False
    
    # Then check if it's a OneDrive site
    return is_onedrive_site(site_url)

def safe_str(value, default=''):
    """Safely convert a value to string, handling None"""
    if value is None:
        return default
    return str(value).strip()

def load_existing_master_report(master_file):
    """Load existing master report using Site ID as primary key"""
    if not os.path.exists(master_file):
        return {}, []
    
    try:
        existing_sites = {}
        site_ids = []
        with open(master_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                site_id = row.get('Site ID', '')
                if site_id:
                    existing_sites[site_id] = row
                    site_ids.append(site_id)
        return existing_sites, site_ids
    except Exception as e:
        print(f"Warning: Could not load master report: {str(e)}")
        return {}, []

def write_to_log(log_file, message, enable_logging=True):
    """Write a message to the log file"""
    if not enable_logging:
        return
    
    try:
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception as e:
        print(f"Warning: Could not write to log file: {str(e)}")

def is_deleted_user(user_info):
    """Check if user is deleted based on status"""
    if not user_info:
        return False
    return user_info.get('is_deleted', False) or user_info.get('status') == 'Deleted'

def update_master_report(current_sites, master_file, config):
    """Update master report with change tracking using Site ID as primary key"""
    # Load existing master report
    existing_sites, existing_site_ids = load_existing_master_report(master_file)
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Get logging configuration
    log_file = config.get('log_file', 'change_log.txt')
    enable_logging = config.get('enable_logging', True)
    
    # Check if this is the first run
    is_first_run = len(existing_sites) == 0
    
    if is_first_run:
        print("\n📝 FIRST RUN DETECTED - Creating baseline without history")
        write_to_log(log_file, "=" * 80, enable_logging)
        write_to_log(log_file, f"FIRST RUN - Baseline created with {len(current_sites)} sites", enable_logging)
        write_to_log(log_file, "=" * 80, enable_logging)
    
    # Track changes and new sites
    all_changes = []
    newly_added_sites = []
    
    # Prepare data for master report
    master_data = []
    current_site_ids = set()
    
    # Process current sites
    for site in current_sites:
        site_id = safe_str(site.get('site_id', ''))
        if not site_id:
            continue
            
        current_site_ids.add(site_id)
        
        # Get existing data if available
        existing_row = existing_sites.get(site_id, {})
        
        # Check if this is a newly added site
        is_newly_added = not is_first_run and site_id not in existing_sites
        
        if is_newly_added:
            newly_added_sites.append(site)
        
        # Get user info
        user_info = site.get('user_info', {})
        is_user_deleted = is_deleted_user(user_info)
        
        # Build the updated row
        row = {
            'Site ID': site_id,
            'Title': safe_str(site.get('title', existing_row.get('Title', ''))),
            'Template Name': safe_str(site.get('template_name', existing_row.get('Template Name', ''))),
            'Created On': safe_str(site.get('time_created', existing_row.get('Created On', ''))),
            'Storage Used (GB)': float(site.get('storage_used_gb', existing_row.get('Storage Used (GB)', 0)) or 0),
            'Is Newly Added': 'Yes' if is_newly_added else 'No',
            'Last Updated': current_time
        }
        
        # ===== SITE URL CHANGE =====
        new_site_url = safe_str(site.get('site_url', ''))
        old_site_url = safe_str(existing_row.get('Site URL', ''))
        
        if is_first_run:
            row['Site URL'] = new_site_url
            row['Site URL Change History'] = ''
        else:
            if new_site_url and new_site_url != old_site_url and not is_newly_added:
                row['Site URL'] = new_site_url
                history = safe_str(existing_row.get('Site URL Change History', ''))
                change_entry = f"[{current_time}] {old_site_url} >> {new_site_url}"
                row['Site URL Change History'] = f"{history}\n{change_entry}" if history else change_entry
                all_changes.append({
                    'site_id': site_id,
                    'site_url': new_site_url,
                    'field': 'Site URL',
                    'old_value': old_site_url,
                    'new_value': new_site_url
                })
                write_to_log(log_file, f"SITE URL CHANGE | Site: {site_id} | {old_site_url} >> {new_site_url}", enable_logging)
            else:
                row['Site URL'] = old_site_url if old_site_url else new_site_url
                row['Site URL Change History'] = safe_str(existing_row.get('Site URL Change History', ''))
        
        # ===== USER UPN =====
        new_upn = safe_str(user_info.get('upn', ''))
        old_upn = safe_str(existing_row.get('User UPN', ''))
        
        if is_first_run:
            row['User UPN'] = new_upn if new_upn else safe_str(site.get('owner_email', ''))
            row['UPN Change History'] = ''
        else:
            # Check if user is deleted - if so, preserve old UPN and don't track changes
            if is_user_deleted and old_upn:
                row['User UPN'] = old_upn  # Keep the old UPN
                row['UPN Change History'] = safe_str(existing_row.get('UPN Change History', ''))
            elif new_upn and new_upn != old_upn and not is_newly_added:
                row['User UPN'] = new_upn
                history = safe_str(existing_row.get('UPN Change History', ''))
                change_entry = f"[{current_time}] {old_upn} >> {new_upn}"
                row['UPN Change History'] = f"{history}\n{change_entry}" if history else change_entry
                all_changes.append({
                    'site_id': site_id,
                    'site_url': safe_str(site.get('site_url', '')),
                    'field': 'UPN',
                    'old_value': old_upn,
                    'new_value': new_upn
                })
                write_to_log(log_file, f"UPN CHANGE | Site: {site_id} | {old_upn} >> {new_upn}", enable_logging)
            else:
                row['User UPN'] = old_upn if old_upn else (new_upn if new_upn else safe_str(site.get('owner_email', '')))
                row['UPN Change History'] = safe_str(existing_row.get('UPN Change History', ''))
        
        # ===== USER DETAILS =====
        new_mail = safe_str(user_info.get('mail', site.get('owner_email', '')))
        new_display = safe_str(user_info.get('display_name', ''))
        
        row['User Email'] = new_mail if new_mail else safe_str(existing_row.get('User Email', site.get('owner_email', '')))
        row['User Display Name'] = new_display if new_display else safe_str(existing_row.get('User Display Name', ''))
        row['User Account Status'] = safe_str(user_info.get('status', existing_row.get('User Account Status', 'Not Found')))
        row['User Deleted Time'] = safe_str(user_info.get('deleted_time', existing_row.get('User Deleted Time', '')))
        
        # ===== USER ACCOUNT STATUS CHANGE =====
        new_status = safe_str(user_info.get('status', ''))
        old_status = safe_str(existing_row.get('User Account Status', ''))
        
        if is_first_run:
            row['User Account Status Change History'] = ''
        else:
            if new_status and new_status != old_status and not is_newly_added and not is_user_deleted:
                history = safe_str(existing_row.get('User Account Status Change History', ''))
                change_entry = f"[{current_time}] {old_status} >> {new_status}"
                row['User Account Status Change History'] = f"{history}\n{change_entry}" if history else change_entry
                all_changes.append({
                    'site_id': site_id,
                    'site_url': safe_str(site.get('site_url', '')),
                    'field': 'User Status',
                    'old_value': old_status,
                    'new_value': new_status
                })
                write_to_log(log_file, f"USER STATUS CHANGE | Site: {site_id} | {old_status} >> {new_status}", enable_logging)
            else:
                row['User Account Status Change History'] = safe_str(existing_row.get('User Account Status Change History', ''))
        
        # ===== MANAGER =====
        manager_info = site.get('manager_info', {})
        new_manager_upn = safe_str(manager_info.get('upn', ''))
        old_manager_upn = safe_str(existing_row.get('Manager UPN', ''))
        
        if is_first_run:
            row['Manager UPN'] = new_manager_upn
            row['Manager Change History'] = ''
        else:
            if new_manager_upn and new_manager_upn != old_manager_upn and not is_newly_added:
                row['Manager UPN'] = new_manager_upn
                history = safe_str(existing_row.get('Manager Change History', ''))
                change_entry = f"[{current_time}] {old_manager_upn} >> {new_manager_upn}"
                row['Manager Change History'] = f"{history}\n{change_entry}" if history else change_entry
                all_changes.append({
                    'site_id': site_id,
                    'site_url': safe_str(site.get('site_url', '')),
                    'field': 'Manager',
                    'old_value': old_manager_upn,
                    'new_value': new_manager_upn
                })
                write_to_log(log_file, f"MANAGER CHANGE | Site: {site_id} | {old_manager_upn} >> {new_manager_upn}", enable_logging)
            else:
                row['Manager UPN'] = old_manager_upn if old_manager_upn else new_manager_upn
                row['Manager Change History'] = safe_str(existing_row.get('Manager Change History', ''))
        
        # Manager details
        row['Manager Email'] = safe_str(manager_info.get('mail', existing_row.get('Manager Email', '')))
        row['Manager Display Name'] = safe_str(manager_info.get('display_name', existing_row.get('Manager Display Name', '')))
        row['Manager Status'] = safe_str(manager_info.get('status', existing_row.get('Manager Status', 'Not fetched')))
        
        # ===== ARCHIVE STATUS =====
        new_archive_status = site.get('archive_status')
        if new_archive_status is None:
            new_archive_status = ''
        else:
            new_archive_status = str(new_archive_status).strip()
        
        old_archive_status = existing_row.get('Archive Status', '')
        if old_archive_status is None:
            old_archive_status = ''
        else:
            old_archive_status = str(old_archive_status).strip()
        
        if is_first_run:
            row['Archive Status'] = new_archive_status
            row['Archive Change History'] = ''
        else:
            if new_archive_status != old_archive_status and not is_newly_added:
                if old_archive_status or new_archive_status:
                    row['Archive Status'] = new_archive_status
                    history = safe_str(existing_row.get('Archive Change History', ''))
                    old_display = old_archive_status if old_archive_status else '(empty)'
                    new_display = new_archive_status if new_archive_status else '(empty)'
                    change_entry = f"[{current_time}] {old_display} >> {new_display}"
                    row['Archive Change History'] = f"{history}\n{change_entry}" if history else change_entry
                    all_changes.append({
                        'site_id': site_id,
                        'site_url': safe_str(site.get('site_url', '')),
                        'field': 'Archive Status',
                        'old_value': old_archive_status if old_archive_status else '(empty)',
                        'new_value': new_archive_status if new_archive_status else '(empty)'
                    })
                    write_to_log(log_file, f"ARCHIVE CHANGE | Site: {site_id} | {old_display} >> {new_display}", enable_logging)
                else:
                    row['Archive Status'] = new_archive_status
                    row['Archive Change History'] = safe_str(existing_row.get('Archive Change History', ''))
            else:
                row['Archive Status'] = old_archive_status if old_archive_status else new_archive_status
                row['Archive Change History'] = safe_str(existing_row.get('Archive Change History', ''))
        
        # ===== DELETION STATUS =====
        new_time_deleted = site.get('time_deleted', '')
        if new_time_deleted is None:
            new_time_deleted = ''
        else:
            new_time_deleted = str(new_time_deleted).strip()
        
        old_time_deleted = existing_row.get('Deleted On', '')
        if old_time_deleted is None:
            old_time_deleted = ''
        else:
            old_time_deleted = str(old_time_deleted).strip()
        
        new_is_deleted = bool(new_time_deleted and new_time_deleted.strip())
        old_is_deleted = bool(old_time_deleted and old_time_deleted.strip())
        
        if is_first_run:
            row['Deleted On'] = new_time_deleted
            row['Deletion Change History'] = ''
            row['Deletion Status'] = 'Deleted' if new_is_deleted else ''
        else:
            if new_is_deleted != old_is_deleted and not is_newly_added:
                row['Deleted On'] = new_time_deleted
                history = safe_str(existing_row.get('Deletion Change History', ''))
                
                if new_is_deleted and not old_is_deleted:
                    change_entry = f"[{current_time}] >> Deleted (on {new_time_deleted})"
                    row['Deletion Status'] = 'Deleted'
                    write_to_log(log_file, f"DELETION CHANGE | Site: {site_id} | Site DELETED on {new_time_deleted}", enable_logging)
                elif not new_is_deleted and old_is_deleted:
                    change_entry = f"[{current_time}] >> Restored (Deleted on {old_time_deleted} was removed)"
                    row['Deletion Status'] = ''
                    write_to_log(log_file, f"DELETION CHANGE | Site: {site_id} | Site RESTORED (was deleted on {old_time_deleted})", enable_logging)
                else:
                    change_entry = f"[{current_time}] Status changed"
                    row['Deletion Status'] = 'Deleted' if new_is_deleted else ''
                
                row['Deletion Change History'] = f"{history}\n{change_entry}" if history else change_entry
                all_changes.append({
                    'site_id': site_id,
                    'site_url': safe_str(site.get('site_url', '')),
                    'field': 'Deletion Status',
                    'old_value': 'Deleted' if old_is_deleted else '',
                    'new_value': 'Deleted' if new_is_deleted else ''
                })
            else:
                row['Deleted On'] = old_time_deleted if old_time_deleted else new_time_deleted
                row['Deletion Change History'] = safe_str(existing_row.get('Deletion Change History', ''))
                row['Deletion Status'] = 'Deleted' if new_is_deleted else ''
        
        # ===== COMBINED CHANGE HISTORY (with newlines) =====
        change_histories = []
        
        site_url_history = safe_str(row.get('Site URL Change History', ''))
        if site_url_history:
            change_histories.append(f"Site URL:\n{site_url_history}")
        
        upn_history = safe_str(row.get('UPN Change History', ''))
        if upn_history:
            change_histories.append(f"UPN:\n{upn_history}")
        
        user_status_history = safe_str(row.get('User Account Status Change History', ''))
        if user_status_history:
            change_histories.append(f"User Status:\n{user_status_history}")
        
        manager_history = safe_str(row.get('Manager Change History', ''))
        if manager_history:
            change_histories.append(f"Manager:\n{manager_history}")
        
        archive_history = safe_str(row.get('Archive Change History', ''))
        if archive_history:
            change_histories.append(f"Archive:\n{archive_history}")
        
        deletion_history = safe_str(row.get('Deletion Change History', ''))
        if deletion_history:
            change_histories.append(f"Deletion:\n{deletion_history}")
        
        if change_histories:
            row['Change History'] = "\n\n".join(change_histories)
        else:
            row['Change History'] = ''
        
        master_data.append(row)
    
    # Check for removed sites
    for site_id in existing_site_ids:
        if site_id not in current_site_ids:
            existing_row = existing_sites[site_id]
            row = dict(existing_row)
            row['Last Updated'] = current_time
            if not is_first_run:
                history = safe_str(existing_row.get('Deletion Change History', ''))
                change_entry = f"[{current_time}] Site removed from SharePoint list"
                row['Deletion Change History'] = f"{history}\n{change_entry}" if history else change_entry
                row['Deletion Status'] = 'Deleted' if row.get('Deleted On') else ''
                write_to_log(log_file, f"SITE REMOVED | Site: {site_id} | {existing_row.get('Title', 'Unknown')} - Removed from SharePoint list", enable_logging)
            master_data.append(row)
    
    # Log newly added sites
    if not is_first_run and newly_added_sites:
        for site in newly_added_sites:
            write_to_log(log_file, f"SITE ADDED | Site: {safe_str(site.get('site_id', ''))} | {safe_str(site.get('title', 'Unknown'))} - New site discovered", enable_logging)
    
    # Write master report
    try:
        fieldnames = [
            'Site ID',
            'Site URL',
            'Site URL Change History',
            'Title',
            'Template Name',
            'User UPN',
            'UPN Change History',
            'User Email',
            'User Display Name',
            'User Account Status',
            'User Account Status Change History',
            'User Deleted Time',
            'Manager UPN',
            'Manager Change History',
            'Manager Email',
            'Manager Display Name',
            'Manager Status',
            'Archive Status',
            'Archive Change History',
            'Deleted On',
            'Deletion Status',
            'Deletion Change History',
            'Created On',
            'Storage Used (GB)',
            'Is Newly Added',
            'Change History',
            'Last Updated'
        ]
        
        with open(master_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            master_data_sorted = sorted(master_data, key=lambda x: float(x.get('Storage Used (GB)', 0) or 0), reverse=True)
            writer.writerows(master_data_sorted)
        
        print(f"\n✅ Master report updated: {master_file}")
        
        # Write summary to log
        if not is_first_run:
            write_to_log(log_file, "-" * 80, enable_logging)
            if newly_added_sites:
                write_to_log(log_file, f"SUMMARY: {len(newly_added_sites)} new site(s) added", enable_logging)
            if all_changes:
                write_to_log(log_file, f"SUMMARY: {len(all_changes)} change(s) detected", enable_logging)
            if not newly_added_sites and not all_changes:
                write_to_log(log_file, "SUMMARY: No changes detected", enable_logging)
            write_to_log(log_file, "-" * 80, enable_logging)
            write_to_log(log_file, "", enable_logging)
        
        # Print summary
        if is_first_run:
            print(f"\n📝 FIRST RUN COMPLETE - Baseline created with {len(current_sites)} sites")
            print("   No change history tracked yet. Future runs will track changes.")
            write_to_log(log_file, "=" * 80, enable_logging)
            write_to_log(log_file, f"FIRST RUN COMPLETE - {len(current_sites)} sites added to baseline", enable_logging)
            write_to_log(log_file, "=" * 80, enable_logging)
            write_to_log(log_file, "", enable_logging)
        else:
            if newly_added_sites:
                print(f"\n{'='*60}")
                print("🆕 NEWLY ADDED SITES")
                print(f"{'='*60}")
                print(f"Total newly added sites: {len(newly_added_sites)}")
                write_to_log(log_file, f"NEW SITES ({len(newly_added_sites)}):", enable_logging)
                for site in newly_added_sites[:10]:
                    site_title = safe_str(site.get('title', 'Unknown'))
                    site_id = safe_str(site.get('site_id', 'N/A'))
                    print(f"  • {site_title} (ID: {site_id})")
                    print(f"    URL: {safe_str(site.get('site_url', ''))}")
                    print(f"    Owner: {safe_str(site.get('owner_email', 'Unknown'))}")
                    print(f"    Created: {safe_str(site.get('time_created', 'Unknown'))}")
                    write_to_log(log_file, f"  - {site_title} (ID: {site_id})", enable_logging)
                if len(newly_added_sites) > 10:
                    print(f"  ... and {len(newly_added_sites) - 10} more")
                    write_to_log(log_file, f"  ... and {len(newly_added_sites) - 10} more", enable_logging)
            
            if all_changes:
                print(f"\n{'='*60}")
                print("📊 CHANGES SUMMARY")
                print(f"{'='*60}")
                print(f"Total changes detected: {len(all_changes)}")
                write_to_log(log_file, f"CHANGES ({len(all_changes)}):", enable_logging)
                
                change_types = {}
                for change in all_changes:
                    field = change['field']
                    if field not in change_types:
                        change_types[field] = []
                    change_types[field].append(change)
                
                for field, changes in change_types.items():
                    print(f"\n{field} Changes: {len(changes)}")
                    write_to_log(log_file, f"  {field} ({len(changes)}):", enable_logging)
                    for change in changes[:5]:
                        old_val = change['old_value'] if change['old_value'] else '(empty)'
                        new_val = change['new_value'] if change['new_value'] else '(empty)'
                        print(f"  • Site ID: {change['site_id']} - {change['site_url']}: {old_val} >> {new_val}")
                        write_to_log(log_file, f"    - Site: {change['site_id']} | {old_val} >> {new_val}", enable_logging)
                    if len(changes) > 5:
                        print(f"  ... and {len(changes) - 5} more")
                        write_to_log(log_file, f"    ... and {len(changes) - 5} more", enable_logging)
            else:
                if not newly_added_sites:
                    print(f"\n📊 No changes detected since last run")
        
        return all_changes, newly_added_sites
        
    except Exception as e:
        print(f"Error updating master report: {str(e)}")
        import traceback
        traceback.print_exc()
        return None, None

def get_all_sites_from_list_optimized(token_manager, graph_token_manager, sharepoint_admin_url, list_id, page_size=100, max_workers=20, config=None):
    """Get OneDrive sites with owner and manager information"""
    print(f"\n{'='*60}")
    print("📁 FETCHING ONEDRIVE SITES")
    print(f"{'='*60}")
    
    all_sites = []
    skipped_sites = 0
    ignored_sites = 0
    
    check_owner = config.get('check_owner_exists', True)
    fetch_manager = config.get('fetch_manager', True)
    
    base_endpoint = f"{sharepoint_admin_url}/_api/Web/Lists(guid'{list_id}')/items"
    endpoint = f"{base_endpoint}?$top={page_size}"
    batch_count = 0
    total_sites = 0
    
    while endpoint:
        batch_count += 1
        try:
            print(f"  Processing batch {batch_count}...")
            data = make_sharepoint_request(token_manager, endpoint)
            current_batch = data.get('value', [])
            
            if not current_batch:
                break
            
            print(f"    Found {len(current_batch)} sites in this batch")
            
            for item in current_batch:
                total_sites += 1
                
                site_url = item.get('SiteUrl', '')
                if site_url is None:
                    site_url = ''
                else:
                    site_url = str(site_url)
                
                if should_ignore_url(site_url, config):
                    ignored_sites += 1
                    continue
                
                if is_onedrive_site(site_url):
                    site_info = {
                        'site_id': safe_str(item.get('SiteId', '')),
                        'site_url': site_url,
                        'title': safe_str(item.get('Title', '')),
                        'template_name': safe_str(item.get('TemplateName', '')),
                        'owner_email': safe_str(item.get('CreatedByEmail', '')),
                        'time_created': safe_str(item.get('TimeCreated', '')),
                        'time_deleted': safe_str(item.get('TimeDeleted', '')),
                        'archive_status': safe_str(item.get('ArchiveStatus', '')),
                        'storage_used_gb': round(float(item.get('StorageUsed', 0) or 0) / (1024**3), 2) if item.get('StorageUsed') else 0,
                        'created_by': safe_str(item.get('CreatedBy', '')),
                        'created_by_email': safe_str(item.get('CreatedByEmail', ''))
                    }
                    all_sites.append(site_info)
                else:
                    skipped_sites += 1
            
            endpoint = data.get('odata.nextLink')
        except Exception as e:
            print(f"Error processing batch {batch_count}: {str(e)}")
            break
    
    print(f"\n📊 Total sites processed: {total_sites}")
    print(f"  ✅ OneDrive sites found: {len(all_sites)}")
    print(f"  ⏭️  Ignored sites (pattern match): {ignored_sites}")
    print(f"  ⏭️  Non-OneDrive sites skipped: {skipped_sites}")
    
    if check_owner and all_sites:
        print(f"\n{'='*60}")
        print("👤 CHECKING USER AND MANAGER INFORMATION")
        print(f"{'='*60}")
        print(f"Processing {len(all_sites)} OneDrive sites...")
        print(f"  - Checking user status (Enabled/Disabled/Deleted/Not Found)")
        if fetch_manager:
            print(f"  - Fetching manager for each user")
        
        processed = 0
        start_time = time.time()
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            def process_site(site):
                owner_email = site.get('owner_email', '')
                if owner_email is None:
                    owner_email = ''
                
                user_info = check_user_status(graph_token_manager, owner_email)
                site['user_info'] = user_info
                
                if fetch_manager and user_info.get('exists', False):
                    user_data, _ = get_user_by_email_filter(graph_token_manager, owner_email, 2)
                    if user_data and user_data.get('id'):
                        manager_data, manager_status = get_user_manager(graph_token_manager, user_data.get('id'), 2)
                        if manager_data:
                            site['manager_info'] = {
                                'upn': safe_str(manager_data.get('upn', '')),
                                'mail': safe_str(manager_data.get('mail', '')),
                                'display_name': safe_str(manager_data.get('display_name', '')),
                                'status': 'Found'
                            }
                        else:
                            site['manager_info'] = {
                                'upn': '',
                                'mail': '',
                                'display_name': '',
                                'status': manager_status if manager_status else 'No manager assigned'
                            }
                    else:
                        site['manager_info'] = {
                            'upn': '',
                            'mail': '',
                            'display_name': '',
                            'status': 'User not found'
                        }
                else:
                    site['manager_info'] = {
                        'upn': '',
                        'mail': '',
                        'display_name': '',
                        'status': 'Not fetched'
                    }
                
                return site
            
            futures = {
                executor.submit(process_site, site): site
                for site in all_sites
            }
            
            for future in as_completed(futures):
                try:
                    future.result(timeout=60)
                    processed += 1
                    if processed % 10 == 0:
                        print(f"  Progress: {processed}/{len(all_sites)}")
                except Exception as e:
                    print(f"  ⚠️ Error processing site: {str(e)[:100]}")
        
        elapsed = time.time() - start_time
        print(f"\n✅ Processing completed in {elapsed:.2f} seconds")
        
        enabled = sum(1 for s in all_sites if s.get('user_info', {}).get('status') == 'Enabled')
        disabled = sum(1 for s in all_sites if s.get('user_info', {}).get('status') == 'Disabled')
        deleted = sum(1 for s in all_sites if s.get('user_info', {}).get('is_deleted', False))
        not_found = sum(1 for s in all_sites if s.get('user_info', {}).get('status') == 'Not Found')
        no_email = sum(1 for s in all_sites if not s.get('owner_email'))
        
        managers_found = sum(1 for s in all_sites if s.get('manager_info', {}).get('upn'))
        
        print(f"\n📊 User Status Summary:")
        print(f"  ✅ Enabled: {enabled}")
        print(f"  ⚠️  Disabled: {disabled}")
        print(f"  ❌ Deleted: {deleted}")
        print(f"  🔍 Not Found: {not_found}")
        print(f"  📧 No Email: {no_email}")
        print(f"\n👔 Managers Found: {managers_found}")
    
    return all_sites

def main():
    config = load_config("config.json")
    
    tenant_name = config.get('tenant')
    app_id = config.get('app_id')
    certificate_path = config.get('cert_path')
    private_key_path = config.get('key_path')
    sharepoint_admin_url = config.get('sharepoint_admin_url')
    list_id = config.get('list_id')
    page_size = config.get('page_size', 100)
    max_workers = config.get('max_workers', 20)
    master_report = config.get('master_report', True)
    ignore_pattern = config.get('ignore_url_pattern', r'm_[A-Za-z0-9]+_[A-Za-z0-9]+')
    log_file = config.get('log_file', 'change_log.txt')
    enable_logging = config.get('enable_logging', True)
    
    print(f"\n{'='*60}")
    print("📊 ONEDRIVE MASTER REPORT - WITH CHANGE TRACKING")
    print(f"{'='*60}")
    print(f"📅 Report Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"🏢 Tenant: {tenant_name}")
    print(f"🔍 Ignore URL Pattern: {ignore_pattern}")
    print(f"🔑 Primary Key: Site ID (Unique)")
    print(f"📝 Log File: {log_file}")
    print(f"📝 Logging Enabled: {enable_logging}")
    
    if not sharepoint_admin_url:
        print("Error: sharepoint_admin_url is required in config.json")
        return
    if not list_id:
        print("Error: list_id is required in config.json")
        return
    
    try:
        if enable_logging:
            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(f"\n{'='*80}\n")
                f.write(f"SESSION START: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"{'='*80}\n")
        
        certificate, private_key = load_certificate_and_key(certificate_path, private_key_path)
        print("✅ Certificate and private key loaded successfully")
        
        sharepoint_token_manager = SharePointTokenManager(certificate, private_key, tenant_name, app_id, sharepoint_admin_url)
        graph_token_manager = GraphTokenManager(certificate, private_key, tenant_name, app_id)
        
        sharepoint_token_manager.get_token()
        graph_token_manager.get_token()
        print("✅ Tokens retrieved successfully")
        
        onedrive_sites = get_all_sites_from_list_optimized(
            sharepoint_token_manager,
            graph_token_manager,
            sharepoint_admin_url,
            list_id,
            page_size,
            max_workers,
            config
        )
        
        if not onedrive_sites:
            print("\n⚠️ No OneDrive sites found!")
            if enable_logging:
                with open(log_file, 'a', encoding='utf-8') as f:
                    f.write(f"WARNING: No OneDrive sites found on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                    f.write(f"{'='*80}\n")
            return
        
        tenant_clean = tenant_name.split('.')[0] if '.' in tenant_name else tenant_name
        master_file = f"{tenant_clean}_onedrive_master_report.csv"
        
        is_first_run = not os.path.exists(master_file)
        
        if is_first_run:
            print(f"\n📝 FIRST RUN DETECTED!")
            print("   Creating baseline master report without change history.")
            print("   Change tracking will begin from the next run.")
        
        if master_report:
            changes, newly_added = update_master_report(onedrive_sites, master_file, config)
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            master_file = f"{tenant_clean}_onedrive_report_{timestamp}.csv"
            changes = None
            newly_added = None
        
        print(f"\n{'='*60}")
        print("✅ SCRIPT COMPLETED SUCCESSFULLY!")
        print(f"{'='*60}")
        print(f"📄 Master Report: {master_file}")
        print(f"📝 Log File: {log_file}")
        
        if is_first_run:
            print(f"📝 Status: First run - baseline created with {len(onedrive_sites)} sites")
        else:
            if newly_added:
                print(f"🆕 Newly Added Sites: {len(newly_added)}")
            if changes:
                print(f"📊 Changes detected: {len(changes)}")
            if not newly_added and not changes:
                print(f"📊 No new sites or changes detected")
        
        if enable_logging:
            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(f"SESSION END: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"{'='*80}\n\n")
        
    except Exception as e:
        error_msg = f"❌ An error occurred: {str(e)}"
        print(f"\n{error_msg}")
        if enable_logging:
            with open(log_file, 'a', encoding='utf-8') as f:
                f.write(f"ERROR: {error_msg}\n")
                import traceback
                f.write(f"TRACEBACK:\n{traceback.format_exc()}\n")
                f.write(f"{'='*80}\n\n")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main()
