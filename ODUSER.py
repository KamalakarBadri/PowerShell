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
        
        config.setdefault('page_size', 500)
        config.setdefault('max_retries', 3)
        config.setdefault('max_workers', 50)
        config.setdefault('ignore_url_pattern', r'm_[A-Za-z0-9]+_[A-Za-z0-9]+')
        
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
    "page_size": 500,
    "max_retries": 3,
    "max_workers": 50,
    "ignore_url_pattern": "m_[A-Za-z0-9]+_[A-Za-z0-9]+"
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

def get_user_upn_from_graph(graph_token_manager, user_email, max_retries=3):
    """
    Get user UPN from Graph API using email filter
    """
    if not user_email:
        return ''
    
    try:
        encoded_email = requests.utils.quote(user_email)
        endpoint = f"https://graph.microsoft.com/v1.0/users?$filter=mail eq '{encoded_email}'&$select=userPrincipalName"
        
        for attempt in range(max_retries):
            try:
                headers = {
                    "Authorization": f"Bearer {graph_token_manager.get_token()}",
                    "Accept": "application/json",
                    "Content-Type": "application/json"
                }
                response = requests.get(endpoint, headers=headers, timeout=15)
                
                if response.status_code == 401:
                    print(f"  [Graph Auth] Token expired, renewing... (Attempt {attempt + 1}/{max_retries})")
                    graph_token_manager._renew_token()
                    continue
                
                response.raise_for_status()
                data = response.json()
                value = data.get('value', [])
                if value and len(value) > 0:
                    user = value[0]
                    return user.get('userPrincipalName', '')
                else:
                    return ''
                    
            except requests.exceptions.Timeout:
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return ''
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(1)
                    continue
                return ''
        
        return ''
        
    except Exception as e:
        return ''

def should_ignore_url(site_url, config):
    if not site_url:
        return False
    
    ignore_pattern = config.get('ignore_url_pattern', r'm_[A-Za-z0-9]+_[A-Za-z0-9]+')
    
    if re.search(ignore_pattern, site_url, re.IGNORECASE):
        return True
    
    return False

def is_onedrive_site(site_url):
    if not site_url:
        return False
    site_url_lower = site_url.lower()
    return 'my.sharepoint.com/personal' in site_url_lower

def should_include_site(site_url, config):
    if should_ignore_url(site_url, config):
        return False
    return is_onedrive_site(site_url)

def safe_str(value, default=''):
    if value is None:
        return default
    return str(value).strip()

def load_existing_report(report_file):
    """Load existing report using Site ID as primary key"""
    if not os.path.exists(report_file):
        return {}, []
    
    try:
        existing_sites = {}
        site_ids = []
        with open(report_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                site_id = row.get('Site ID', '')
                if site_id:
                    existing_sites[site_id] = row
                    site_ids.append(site_id)
        return existing_sites, site_ids
    except Exception as e:
        print(f"Warning: Could not load report: {str(e)}")
        return {}, []

def get_all_sites_from_list(token_manager, graph_token_manager, sharepoint_admin_url, list_id, page_size=500, max_workers=50, config=None):
    """Get OneDrive sites with all required parameters including UPN"""
    print(f"\n{'='*60}")
    print("📁 FETCHING ONEDRIVE SITES")
    print(f"{'='*60}")
    
    all_sites = []
    skipped_sites = 0
    ignored_sites = 0
    upn_cache = {}  # Cache to avoid duplicate Graph API calls
    
    # Get ALL required fields: SiteId, SiteUrl, SiteOwnerName, SiteOwnerEmail, CreatedBy, CreatedByEmail
    base_endpoint = f"{sharepoint_admin_url}/_api/Web/Lists(guid'{list_id}')/items"
    endpoint = f"{base_endpoint}?$top={page_size}&$select=SiteId,SiteUrl,SiteOwnerName,SiteOwnerEmail,CreatedBy,CreatedByEmail"
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
                    # Get all required parameters
                    site_owner_email = safe_str(item.get('SiteOwnerEmail', ''))
                    created_by_email = safe_str(item.get('CreatedByEmail', ''))
                    
                    site_info = {
                        'site_id': safe_str(item.get('SiteId', '')),
                        'site_url': site_url,
                        'site_owner_name': safe_str(item.get('SiteOwnerName', '')),
                        'site_owner_email': site_owner_email,
                        'created_by': safe_str(item.get('CreatedBy', '')),
                        'created_by_email': created_by_email,
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
    
    # Get UPN for each site using Graph API (with caching)
    if all_sites:
        print(f"\n{'='*60}")
        print("🔍 FETCHING USER UPN FROM GRAPH API")
        print(f"{'='*60}")
        print(f"Processing {len(all_sites)} sites...")
        print(f"  - Graph API permission required: User.Read.All or Directory.Read.All")
        
        processed = 0
        start_time = time.time()
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            def process_site(site):
                site_owner_email = site.get('site_owner_email', '')
                created_by_email = site.get('created_by_email', '')
                
                # Get UPN for Site Owner Email (with caching)
                if site_owner_email:
                    if site_owner_email not in upn_cache:
                        upn_cache[site_owner_email] = get_user_upn_from_graph(graph_token_manager, site_owner_email)
                    site['user_upn'] = upn_cache[site_owner_email]
                else:
                    site['user_upn'] = ''
                
                # Get UPN for Created By Email (with caching)
                if created_by_email:
                    if created_by_email not in upn_cache:
                        upn_cache[created_by_email] = get_user_upn_from_graph(graph_token_manager, created_by_email)
                    site['created_by_upn'] = upn_cache[created_by_email]
                else:
                    site['created_by_upn'] = ''
                
                return site
            
            futures = {
                executor.submit(process_site, site): site
                for site in all_sites
            }
            
            for future in as_completed(futures):
                try:
                    future.result(timeout=30)
                    processed += 1
                    if processed % 50 == 0 or processed == 1:
                        print(f"  Progress: {processed}/{len(all_sites)}")
                except Exception as e:
                    print(f"  ⚠️ Error processing site: {str(e)[:100]}")
        
        elapsed = time.time() - start_time
        print(f"\n✅ UPN fetching completed in {elapsed:.2f} seconds")
        
        # Count UPNs found
        upn_found = sum(1 for s in all_sites if s.get('user_upn', ''))
        created_upn_found = sum(1 for s in all_sites if s.get('created_by_upn', ''))
        
        print(f"\n📊 UPN Summary:")
        print(f"  ✅ Site Owner UPN found: {upn_found}")
        print(f"  ❌ Site Owner UPN not found: {len(all_sites) - upn_found}")
        print(f"  ✅ Created By UPN found: {created_upn_found}")
        print(f"  ❌ Created By UPN not found: {len(all_sites) - created_upn_found}")
    
    return all_sites

def update_report(current_sites, report_file, config):
    """Update report with change tracking - each field has its own history column"""
    existing_sites, existing_site_ids = load_existing_report(report_file)
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    is_first_run = len(existing_sites) == 0
    
    if is_first_run:
        print("\n📝 FIRST RUN DETECTED - Creating baseline report")
    
    all_changes = []
    newly_added_sites = []
    
    master_data = []
    current_site_ids = set()
    
    for site in current_sites:
        site_id = safe_str(site.get('site_id', ''))
        if not site_id:
            continue
        
        current_site_ids.add(site_id)
        existing_row = existing_sites.get(site_id, {})
        
        is_newly_added = not is_first_run and site_id not in existing_sites
        
        if is_newly_added:
            newly_added_sites.append(site)
        
        # Build row with ALL required parameters and their history columns
        row = {
            'Site ID': site_id,
            
            # Site URL and its history
            'Site URL': safe_str(site.get('site_url', existing_row.get('Site URL', ''))),
            'Site URL Change History': existing_row.get('Site URL Change History', ''),
            
            # Site Owner Name and its history
            'Site Owner Name': safe_str(site.get('site_owner_name', existing_row.get('Site Owner Name', ''))),
            'Site Owner Name Change History': existing_row.get('Site Owner Name Change History', ''),
            
            # Site Owner Email and its history
            'Site Owner Email': safe_str(site.get('site_owner_email', existing_row.get('Site Owner Email', ''))),
            'Site Owner Email Change History': existing_row.get('Site Owner Email Change History', ''),
            
            # User UPN and its history (NEW)
            'User UPN': safe_str(site.get('user_upn', existing_row.get('User UPN', ''))),
            'User UPN Change History': existing_row.get('User UPN Change History', ''),
            
            # Created By (display name) and its history
            'Created By': safe_str(site.get('created_by', existing_row.get('Created By', ''))),
            'Created By Change History': existing_row.get('Created By Change History', ''),
            
            # Created By Email and its history
            'Created By Email': safe_str(site.get('created_by_email', existing_row.get('Created By Email', ''))),
            'Created By Email Change History': existing_row.get('Created By Email Change History', ''),
            
            # Created By UPN and its history (NEW)
            'Created By UPN': safe_str(site.get('created_by_upn', existing_row.get('Created By UPN', ''))),
            'Created By UPN Change History': existing_row.get('Created By UPN Change History', ''),
            
            'Is Newly Added': 'Yes' if is_newly_added else 'No',
            'Last Updated': current_time
        }
        
        # Track changes for each field
        if not is_first_run and not is_newly_added:
            changes_detected = False
            
            # Check Site URL change
            old_url = existing_row.get('Site URL', '')
            new_url = row['Site URL']
            if new_url != old_url:
                old_history = existing_row.get('Site URL Change History', '')
                change_entry = f"[{current_time}] {old_url} >> {new_url}"
                row['Site URL Change History'] = f"{old_history}\n{change_entry}" if old_history else change_entry
                changes_detected = True
                all_changes.append({
                    'site_id': site_id,
                    'field': 'Site URL',
                    'old_value': old_url,
                    'new_value': new_url
                })
            
            # Check Site Owner Name change
            old_name = existing_row.get('Site Owner Name', '')
            new_name = row['Site Owner Name']
            if new_name != old_name:
                old_history = existing_row.get('Site Owner Name Change History', '')
                change_entry = f"[{current_time}] {old_name} >> {new_name}"
                row['Site Owner Name Change History'] = f"{old_history}\n{change_entry}" if old_history else change_entry
                changes_detected = True
                all_changes.append({
                    'site_id': site_id,
                    'field': 'Site Owner Name',
                    'old_value': old_name,
                    'new_value': new_name
                })
            
            # Check Site Owner Email change
            old_email = existing_row.get('Site Owner Email', '')
            new_email = row['Site Owner Email']
            if new_email != old_email:
                old_history = existing_row.get('Site Owner Email Change History', '')
                change_entry = f"[{current_time}] {old_email} >> {new_email}"
                row['Site Owner Email Change History'] = f"{old_history}\n{change_entry}" if old_history else change_entry
                changes_detected = True
                all_changes.append({
                    'site_id': site_id,
                    'field': 'Site Owner Email',
                    'old_value': old_email,
                    'new_value': new_email
                })
            
            # Check User UPN change (NEW)
            old_upn = existing_row.get('User UPN', '')
            new_upn = row['User UPN']
            if new_upn != old_upn:
                old_history = existing_row.get('User UPN Change History', '')
                change_entry = f"[{current_time}] {old_upn} >> {new_upn}"
                row['User UPN Change History'] = f"{old_history}\n{change_entry}" if old_history else change_entry
                changes_detected = True
                all_changes.append({
                    'site_id': site_id,
                    'field': 'User UPN',
                    'old_value': old_upn,
                    'new_value': new_upn
                })
            
            # Check Created By (display name) change
            old_created = existing_row.get('Created By', '')
            new_created = row['Created By']
            if new_created != old_created:
                old_history = existing_row.get('Created By Change History', '')
                change_entry = f"[{current_time}] {old_created} >> {new_created}"
                row['Created By Change History'] = f"{old_history}\n{change_entry}" if old_history else change_entry
                changes_detected = True
                all_changes.append({
                    'site_id': site_id,
                    'field': 'Created By',
                    'old_value': old_created,
                    'new_value': new_created
                })
            
            # Check Created By Email change
            old_created_email = existing_row.get('Created By Email', '')
            new_created_email = row['Created By Email']
            if new_created_email != old_created_email:
                old_history = existing_row.get('Created By Email Change History', '')
                change_entry = f"[{current_time}] {old_created_email} >> {new_created_email}"
                row['Created By Email Change History'] = f"{old_history}\n{change_entry}" if old_history else change_entry
                changes_detected = True
                all_changes.append({
                    'site_id': site_id,
                    'field': 'Created By Email',
                    'old_value': old_created_email,
                    'new_value': new_created_email
                })
            
            # Check Created By UPN change (NEW)
            old_created_upn = existing_row.get('Created By UPN', '')
            new_created_upn = row['Created By UPN']
            if new_created_upn != old_created_upn:
                old_history = existing_row.get('Created By UPN Change History', '')
                change_entry = f"[{current_time}] {old_created_upn} >> {new_created_upn}"
                row['Created By UPN Change History'] = f"{old_history}\n{change_entry}" if old_history else change_entry
                changes_detected = True
                all_changes.append({
                    'site_id': site_id,
                    'field': 'Created By UPN',
                    'old_value': old_created_upn,
                    'new_value': new_created_upn
                })
        
        master_data.append(row)
    
    # Check for removed sites
    for site_id in existing_site_ids:
        if site_id not in current_site_ids:
            existing_row = existing_sites[site_id]
            row = dict(existing_row)
            row['Last Updated'] = current_time
            row['Is Newly Added'] = 'No'
            master_data.append(row)
    
    # Write report
    try:
        fieldnames = [
            'Site ID',
            'Site URL',
            'Site URL Change History',
            'Site Owner Name',
            'Site Owner Name Change History',
            'Site Owner Email',
            'Site Owner Email Change History',
            'User UPN',
            'User UPN Change History',
            'Created By',
            'Created By Change History',
            'Created By Email',
            'Created By Email Change History',
            'Created By UPN',
            'Created By UPN Change History',
            'Is Newly Added',
            'Last Updated'
        ]
        
        with open(report_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(master_data)
        
        print(f"\n✅ Report updated: {report_file}")
        
        # Print summary
        if is_first_run:
            print(f"\n📝 FIRST RUN COMPLETE - Baseline created with {len(current_sites)} sites")
        else:
            if newly_added_sites:
                print(f"\n🆕 Newly Added Sites: {len(newly_added_sites)}")
                for site in newly_added_sites[:5]:
                    print(f"  • {safe_str(site.get('site_url', ''))}")
                    print(f"    Owner: {safe_str(site.get('site_owner_name', 'Unknown'))}")
                if len(newly_added_sites) > 5:
                    print(f"  ... and {len(newly_added_sites) - 5} more")
            
            if all_changes:
                print(f"\n📊 Changes detected: {len(all_changes)}")
                for change in all_changes[:5]:
                    print(f"  • Site ID: {change['site_id']}")
                    print(f"    {change['field']}: {change['old_value']} >> {change['new_value']}")
                if len(all_changes) > 5:
                    print(f"  ... and {len(all_changes) - 5} more changes")
            
            if not newly_added_sites and not all_changes:
                print(f"\n📊 No changes detected since last run")
        
        return all_changes, newly_added_sites
        
    except Exception as e:
        print(f"Error updating report: {str(e)}")
        return None, None

def main():
    config = load_config("config.json")
    
    tenant_name = config.get('tenant')
    app_id = config.get('app_id')
    certificate_path = config.get('cert_path')
    private_key_path = config.get('key_path')
    sharepoint_admin_url = config.get('sharepoint_admin_url')
    list_id = config.get('list_id')
    page_size = config.get('page_size', 500)
    max_workers = config.get('max_workers', 50)
    ignore_pattern = config.get('ignore_url_pattern', r'm_[A-Za-z0-9]+_[A-Za-z0-9]+')
    
    print(f"\n{'='*60}")
    print("📊 ONEDRIVE OWNER REPORT")
    print(f"{'='*60}")
    print(f"📅 Report Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"🏢 Tenant: {tenant_name}")
    print(f"🔍 Ignore URL Pattern: {ignore_pattern}")
    print(f"🔑 Primary Key: Site ID (Unique)")
    print(f"📝 Data Sources: SharePoint + Graph API (for UPN)")
    print(f"🔐 Graph Permission Required: User.Read.All or Directory.Read.All")
    
    if not sharepoint_admin_url:
        print("Error: sharepoint_admin_url is required in config.json")
        return
    if not list_id:
        print("Error: list_id is required in config.json")
        return
    
    try:
        certificate, private_key = load_certificate_and_key(certificate_path, private_key_path)
        print("✅ Certificate and private key loaded successfully")
        
        sharepoint_token_manager = SharePointTokenManager(certificate, private_key, tenant_name, app_id, sharepoint_admin_url)
        graph_token_manager = GraphTokenManager(certificate, private_key, tenant_name, app_id)
        
        sharepoint_token_manager.get_token()
        graph_token_manager.get_token()
        print("✅ SharePoint and Graph tokens retrieved successfully")
        
        onedrive_sites = get_all_sites_from_list(
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
        
        tenant_clean = tenant_name.split('.')[0] if '.' in tenant_name else tenant_name
        report_file = f"{tenant_clean}_onedrive_owner_report.csv"
        
        is_first_run = not os.path.exists(report_file)
        
        if is_first_run:
            print(f"\n📝 FIRST RUN DETECTED!")
            print("   Creating baseline report without change history.")
            print("   Change tracking will begin from the next run.")
        
        changes, newly_added = update_report(onedrive_sites, report_file, config)
        
        print(f"\n{'='*60}")
        print("✅ SCRIPT COMPLETED SUCCESSFULLY!")
        print(f"{'='*60}")
        print(f"📄 Report: {report_file}")
        
        if is_first_run:
            print(f"📝 Status: First run - baseline created with {len(onedrive_sites)} sites")
        else:
            if newly_added:
                print(f"🆕 Newly Added Sites: {len(newly_added)}")
            if changes:
                print(f"📊 Changes detected: {len(changes)}")
            if not newly_added and not changes:
                print(f"📊 No new sites or changes detected")
        
    except Exception as e:
        print(f"\n❌ An error occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main()
