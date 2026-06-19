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
        config.setdefault('track_archive_status', True)
        config.setdefault('track_deletion_status', True)
        config.setdefault('track_user_status', True)
        
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
    "track_archive_status": true,
    "track_deletion_status": true,
    "track_user_status": true
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
            'user_type': user_data.get('user_type', ''),
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
            'user_type': 'Deleted',
            'is_deleted': True,
            'deleted_time': deleted_time
        }
    
    return {
        'exists': False,
        'status': 'Not Found',
        'upn': '',
        'mail': user_email,
        'display_name': '',
        'user_type': '',
        'is_deleted': False,
        'deleted_time': ''
    }

def is_onedrive_site(site_url):
    if not site_url:
        return False
    site_url_lower = site_url.lower()
    return 'my.sharepoint.com/personal' in site_url_lower

def should_include_site(site_url, config):
    return is_onedrive_site(site_url)

def load_existing_master_report(master_file):
    """Load existing master report with all historical data"""
    if not os.path.exists(master_file):
        return {}, []
    
    try:
        existing_sites = {}
        site_urls = []
        with open(master_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                site_url = row.get('Site URL', '')
                if site_url:
                    existing_sites[site_url] = row
                    site_urls.append(site_url)
        return existing_sites, site_urls
    except Exception as e:
        print(f"Warning: Could not load master report: {str(e)}")
        return {}, []

def update_change_history(existing_row, current_value, field_name, new_value, changes):
    """
    Update change history for a specific field
    Returns: updated change_history string
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Get existing history
    history_field = f"{field_name}_history"
    old_history = existing_row.get(history_field, '') if existing_row else ''
    
    # Get old value
    old_value = existing_row.get(field_name, '') if existing_row else ''
    
    # Only update if value actually changed
    if str(old_value) != str(new_value) and new_value:
        change_entry = f"[{timestamp}] {old_value} → {new_value}"
        if old_history:
            new_history = f"{old_history} | {change_entry}"
        else:
            new_history = change_entry
        
        # Track the change for summary
        changes.append({
            'site_url': existing_row.get('Site URL', ''),
            'field': field_name,
            'old_value': old_value,
            'new_value': new_value,
            'timestamp': timestamp
        })
        
        return new_history
    
    return old_history

def update_master_report(current_sites, master_file, config):
    """Update master report preserving all historical data"""
    # Load existing master report
    existing_sites, existing_urls = load_existing_master_report(master_file)
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    today = datetime.now().strftime("%Y-%m-%d")
    
    # Track changes for summary
    all_changes = []
    
    # Prepare data for master report
    master_data = []
    current_urls = set()
    
    # Process current sites
    for site in current_sites:
        site_url = site['site_url']
        current_urls.add(site_url)
        
        # Get existing data if available
        existing_row = existing_sites.get(site_url, {})
        
        # Build the updated row with history preservation
        row = {
            'Site URL': site_url,
            'Title': site.get('title', existing_row.get('Title', '')),
            'Site ID': site.get('site_id', existing_row.get('Site ID', '')),
            'Template Name': site.get('template_name', existing_row.get('Template Name', '')),
            'Owner Email': site.get('owner_email', existing_row.get('Owner Email', '')),
            'Created On': site.get('time_created', existing_row.get('Created On', '')),
            'Storage Used (GB)': site.get('storage_used_gb', existing_row.get('Storage Used (GB)', 0)),
            'Storage Quota (GB)': site.get('storage_quota_gb', existing_row.get('Storage Quota (GB)', 0)),
            'Last Updated': current_time
        }
        
        # ===== TRACK DELETION STATUS =====
        new_deleted_on = site.get('time_deleted', '')
        old_deleted_on = existing_row.get('Deleted On', '')
        
        # If site was deleted, update Deleted On
        if new_deleted_on and not old_deleted_on:
            row['Deleted On'] = new_deleted_on
            row['deleted_on_history'] = update_change_history(
                existing_row, 'Deleted On', 'deleted_on',
                f"Site deleted on {new_deleted_on}", all_changes
            )
        elif old_deleted_on and not new_deleted_on:
            # Site was restored
            row['Deleted On'] = ''
            row['deleted_on_history'] = update_change_history(
                existing_row, 'Deleted On', 'deleted_on',
                "Site restored", all_changes
            )
        else:
            row['Deleted On'] = old_deleted_on
            row['deleted_on_history'] = existing_row.get('deleted_on_history', '')
        
        # ===== TRACK ARCHIVE STATUS =====
        new_archive_status = site.get('archive_status', '')
        old_archive_status = existing_row.get('Archive Status', '')
        
        if new_archive_status != old_archive_status:
            row['Archive Status'] = new_archive_status
            row['archive_status_history'] = update_change_history(
                existing_row, 'Archive Status', 'archive_status',
                f"Changed from '{old_archive_status}' to '{new_archive_status}'", all_changes
            )
        else:
            row['Archive Status'] = old_archive_status
            row['archive_status_history'] = existing_row.get('archive_status_history', '')
        
        # ===== TRACK USER STATUS =====
        user_info = site.get('user_info', {})
        new_status = user_info.get('status', 'Not Found')
        old_status = existing_row.get('User Account Status', '')
        
        if new_status != old_status and new_status:
            row['User Account Status'] = new_status
            row['user_status_history'] = update_change_history(
                existing_row, 'User Account Status', 'user_status',
                f"{old_status} → {new_status}", all_changes
            )
        else:
            row['User Account Status'] = old_status or new_status
            row['user_status_history'] = existing_row.get('user_status_history', '')
        
        # User UPN - preserve if not found
        new_upn = user_info.get('upn', '')
        old_upn = existing_row.get('User UPN', '')
        if new_upn:
            row['User UPN'] = new_upn
        else:
            row['User UPN'] = old_upn  # Keep historical UPN
        
        # User Email - preserve if not found
        new_mail = user_info.get('mail', site.get('owner_email', ''))
        old_mail = existing_row.get('User Email', '')
        if new_mail:
            row['User Email'] = new_mail
        else:
            row['User Email'] = old_mail
        
        # User Display Name - preserve if not found
        new_display = user_info.get('display_name', '')
        old_display = existing_row.get('User Display Name', '')
        if new_display:
            row['User Display Name'] = new_display
        else:
            row['User Display Name'] = old_display
        
        # User Type
        new_user_type = user_info.get('user_type', '')
        old_user_type = existing_row.get('User Type', '')
        row['User Type'] = new_user_type if new_user_type else old_user_type
        
        # Is Deleted User
        row['Is Deleted User'] = 'Yes' if user_info.get('is_deleted', False) else 'No'
        
        # User Deleted Time
        new_deleted_time = user_info.get('deleted_time', '')
        if new_deleted_time:
            row['User Deleted Time'] = new_deleted_time
        else:
            row['User Deleted Time'] = existing_row.get('User Deleted Time', '')
        
        # ===== TRACK MANAGER DETAILS =====
        manager_info = site.get('manager_info', {})
        new_manager_upn = manager_info.get('upn', '')
        old_manager_upn = existing_row.get('Manager UPN', '')
        
        if new_manager_upn and new_manager_upn != old_manager_upn:
            row['Manager UPN'] = new_manager_upn
            row['manager_upn_history'] = update_change_history(
                existing_row, 'Manager UPN', 'manager_upn',
                f"{old_manager_upn} → {new_manager_upn}", all_changes
            )
        else:
            row['Manager UPN'] = old_manager_upn or new_manager_upn
            row['manager_upn_history'] = existing_row.get('manager_upn_history', '')
        
        # Manager Email - preserve if not found
        new_manager_mail = manager_info.get('mail', '')
        old_manager_mail = existing_row.get('Manager Email', '')
        row['Manager Email'] = new_manager_mail if new_manager_mail else old_manager_mail
        
        # Manager Display Name - preserve if not found
        new_manager_display = manager_info.get('display_name', '')
        old_manager_display = existing_row.get('Manager Display Name', '')
        row['Manager Display Name'] = new_manager_display if new_manager_display else old_manager_display
        
        # Manager Status
        new_manager_status = manager_info.get('status', 'Not fetched')
        old_manager_status = existing_row.get('Manager Status', '')
        if new_manager_status != old_manager_status:
            row['Manager Status'] = new_manager_status
        else:
            row['Manager Status'] = old_manager_status or new_manager_status
        
        # ===== OWNER STATUS =====
        new_owner_exists = user_info.get('exists', False)
        if new_owner_exists:
            row['Owner Exists'] = 'Yes'
            row['Owner Status'] = 'Found'
        else:
            row['Owner Exists'] = existing_row.get('Owner Exists', 'No')
            row['Owner Status'] = user_info.get('status', 'Not Found')
        
        # ===== SUMMARY FIELD =====
        # Create a summary of last change
        if all_changes:
            last_change = all_changes[-1]
            row['Last Change Summary'] = f"{last_change['field']}: {last_change['old_value']} → {last_change['new_value']} ({last_change['timestamp']})"
        else:
            row['Last Change Summary'] = existing_row.get('Last Change Summary', 'No changes')
        
        master_data.append(row)
    
    # Check for removed sites - keep them in report but mark as removed
    for url in existing_urls:
        if url not in current_urls:
            existing_row = existing_sites[url]
            row = dict(existing_row)  # Copy existing data
            row['Last Updated'] = current_time
            row['Last Change Summary'] = f"Site removed from SharePoint list on {current_time}"
            row['deleted_on_history'] = update_change_history(
                existing_row, 'Deleted On', 'deleted_on',
                f"Site removed from list on {current_time}", all_changes
            )
            master_data.append(row)
    
    # Write master report
    try:
        fieldnames = [
            'Site URL',
            'Title',
            'Site ID',
            'Template Name',
            'Owner Email',
            'User UPN',
            'User Email',
            'User Display Name',
            'User Account Status',
            'User Type',
            'Is Deleted User',
            'User Deleted Time',
            'Manager UPN',
            'Manager Email',
            'Manager Display Name',
            'Manager Status',
            'Owner Exists',
            'Owner Status',
            'Created On',
            'Deleted On',
            'Archive Status',
            'Storage Used (GB)',
            'Storage Quota (GB)',
            'Last Updated',
            'Last Change Summary',
            'deleted_on_history',
            'archive_status_history',
            'user_status_history',
            'manager_upn_history'
        ]
        
        with open(master_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(master_data)
        
        print(f"\n✅ Master report updated: {master_file}")
        
        # Print changes summary
        if all_changes:
            print(f"\n{'='*60}")
            print("📊 CHANGES SUMMARY")
            print(f"{'='*60}")
            print(f"Total changes detected: {len(all_changes)}")
            
            # Group changes by type
            change_types = {}
            for change in all_changes:
                field = change['field']
                if field not in change_types:
                    change_types[field] = []
                change_types[field].append(change)
            
            for field, changes in change_types.items():
                print(f"\n{field.replace('_', ' ').title()} Changes: {len(changes)}")
                for change in changes[:5]:  # Show first 5
                    site_title = existing_sites.get(change['site_url'], {}).get('Title', change['site_url'])
                    print(f"  • {site_title}: {change['old_value']} → {change['new_value']}")
                if len(changes) > 5:
                    print(f"  ... and {len(changes) - 5} more")
        
        return all_changes
        
    except Exception as e:
        print(f"Error updating master report: {str(e)}")
        return None

def get_all_sites_from_list_optimized(token_manager, graph_token_manager, sharepoint_admin_url, list_id, page_size=100, max_workers=20, config=None):
    """Get OneDrive sites with owner and manager information"""
    print(f"\n{'='*60}")
    print("📁 FETCHING ONEDRIVE SITES")
    print(f"{'='*60}")
    
    all_sites = []
    skipped_sites = 0
    
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
                
                if should_include_site(site_url, config):
                    site_info = {
                        'site_url': site_url,
                        'title': item.get('Title', ''),
                        'site_id': item.get('SiteId', ''),
                        'template_name': item.get('TemplateName', ''),
                        'owner_email': item.get('CreatedByEmail', ''),
                        'time_created': item.get('TimeCreated', ''),
                        'time_deleted': item.get('TimeDeleted', ''),
                        'archive_status': item.get('ArchiveStatus', ''),
                        'storage_used_gb': round(item.get('StorageUsed', 0) / (1024**3), 2) if item.get('StorageUsed') else 0,
                        'storage_quota_gb': round(item.get('StorageQuota', 0) / (1024**3), 2) if item.get('StorageQuota') else 0,
                        'created_by': item.get('CreatedBy', ''),
                        'created_by_email': item.get('CreatedByEmail', '')
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
    print(f"  ⏭️  Non-OneDrive sites skipped: {skipped_sites}")
    
    # Check owner and manager for all sites
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
                
                # Check user status
                user_info = check_user_status(graph_token_manager, owner_email)
                site['user_info'] = user_info
                
                # Get manager if user exists and manager fetch is enabled
                if fetch_manager and user_info.get('exists', False):
                    # We need user ID to get manager
                    user_data, _ = get_user_by_email_filter(graph_token_manager, owner_email, 2)
                    if user_data and user_data.get('id'):
                        manager_data, manager_status = get_user_manager(graph_token_manager, user_data.get('id'), 2)
                        if manager_data:
                            site['manager_info'] = {
                                'upn': manager_data.get('upn', ''),
                                'mail': manager_data.get('mail', ''),
                                'display_name': manager_data.get('display_name', ''),
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
        
        # Count statuses
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
    # Load configuration
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
    
    print(f"\n{'='*60}")
    print("📊 ONEDRIVE MASTER REPORT - WITH CHANGE HISTORY")
    print(f"{'='*60}")
    print(f"📅 Report Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"🏢 Tenant: {tenant_name}")
    
    if not sharepoint_admin_url:
        print("Error: sharepoint_admin_url is required in config.json")
        return
    if not list_id:
        print("Error: list_id is required in config.json")
        return
    
    try:
        # Load certificates
        certificate, private_key = load_certificate_and_key(certificate_path, private_key_path)
        print("✅ Certificate and private key loaded successfully")
        
        # Initialize token managers
        sharepoint_token_manager = SharePointTokenManager(certificate, private_key, tenant_name, app_id, sharepoint_admin_url)
        graph_token_manager = GraphTokenManager(certificate, private_key, tenant_name, app_id)
        
        # Get initial tokens
        sharepoint_token_manager.get_token()
        graph_token_manager.get_token()
        print("✅ Tokens retrieved successfully")
        
        # Get OneDrive sites
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
            return
        
        # Master report filename
        tenant_clean = tenant_name.split('.')[0] if '.' in tenant_name else tenant_name
        master_file = f"{tenant_clean}_onedrive_master_report.csv"
        
        # Update master report
        if master_report:
            changes = update_master_report(onedrive_sites, master_file, config)
        else:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            master_file = f"{tenant_clean}_onedrive_report_{timestamp}.csv"
            # Simple save without history tracking
            save_simple_report(onedrive_sites, master_file)
            changes = None
        
        print(f"\n{'='*60}")
        print("✅ SCRIPT COMPLETED SUCCESSFULLY!")
        print(f"{'='*60}")
        print(f"📄 Master Report: {master_file}")
        if changes:
            print(f"📊 Changes detected: {len(changes)}")
        
    except Exception as e:
        print(f"\n❌ An error occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main()
