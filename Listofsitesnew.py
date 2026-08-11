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
        self.refresh_buffer = 300  # Refresh 5 minutes before expiry
    
    def get_token(self):
        """Get valid token, renew if expired or about to expire"""
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
        
        # Token expires in 3600 seconds (1 hour), set expiry 45 minutes from now
        self.token_expiry_time = time.time() + 2700  # 45 minutes
        
        print(f"  [Token] Token renewed, expires at {datetime.fromtimestamp(self.token_expiry_time).strftime('%H:%M:%S')}")

def load_config(config_file="config.json"):
    """Load configuration from JSON file"""
    try:
        with open(config_file, 'r') as f:
            config = json.load(f)
        
        config.setdefault('page_size', 5000)
        config.setdefault('max_retries', 3)
        config.setdefault('max_concurrent_requests', 10)
        
        return config
    except FileNotFoundError:
        print(f"Error: Configuration file '{config_file}' not found.")
        print("Please create a config.json file with the following structure:")
        print("""
{
    "tenants": {
        "tenant1": {
            "name": "tenant1.onmicrosoft.com",
            "app_id": "your-app-id-1",
            "cert_path": "cert1.pem",
            "key_path": "key1.pem",
            "sharepoint_admin_url": "https://tenant1-admin.sharepoint.com",
            "list_id": "317f59e4-b925-4d1c-884c-c758bf067a6c",
            "ignore_url_contains": ["sites/deleted", "sites/test", "sites/archive"]
        },
        "tenant2": {
            "name": "tenant2.onmicrosoft.com",
            "app_id": "your-app-id-2",
            "cert_path": "cert2.pem",
            "key_path": "key2.pem",
            "sharepoint_admin_url": "https://tenant2-admin.sharepoint.com",
            "list_id": "317f59e4-b925-4d1c-884c-c758bf067a6c",
            "ignore_url_contains": ["sites/temp", "sites/backup"]
        },
        "tenant3": {
            "name": "tenant3.onmicrosoft.com",
            "app_id": "your-app-id-3",
            "cert_path": "cert3.pem",
            "key_path": "key3.pem",
            "sharepoint_admin_url": "https://tenant3-admin.sharepoint.com",
            "list_id": "317f59e4-b925-4d1c-884c-c758bf067a6c",
            "ignore_url_contains": []
        }
    },
    "page_size": 5000,
    "max_retries": 3,
    "max_concurrent_requests": 10
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
    """Get access token from Microsoft Identity Platform"""
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
        print(f"Error getting access token: {err}")
        raise

def make_sharepoint_request_with_retry(token_manager, endpoint, max_retries=3):
    """Make SharePoint request with automatic token renewal on 401 error"""
    headers = {
        "Authorization": f"Bearer {token_manager.get_token()}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    
    for attempt in range(max_retries):
        try:
            response = requests.get(endpoint, headers=headers)
            
            # If unauthorized, renew token and retry
            if response.status_code == 401:
                print(f"  [Auth] Token expired, renewing... (Attempt {attempt + 1}/{max_retries})")
                token_manager._renew_token()
                headers["Authorization"] = f"Bearer {token_manager.get_token()}"
                continue
            
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.HTTPError as err:
            if response.status_code == 401 and attempt < max_retries - 1:
                continue
            print(f"HTTP Error: {err}")
            print(f"Response: {response.text}")
            raise
        except Exception as err:
            print(f"Error making SharePoint request: {err}")
            raise
    
    raise Exception(f"Failed after {max_retries} attempts")

def safe_float_conversion(value, default=0.0):
    """Safely convert a value to float, handling errors"""
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip()
        if 'error' in cleaned.lower() or '#' in cleaned:
            return default
        try:
            return float(cleaned)
        except ValueError:
            return default
    return default

def should_ignore_site(site_url, ignore_list):
    """Check if site should be ignored based on URL containing any keyword"""
    if not ignore_list:
        return False
    url_lower = site_url.lower()
    for keyword in ignore_list:
        if keyword.lower() in url_lower:
            return True
    return False

def get_all_sites_from_list(token_manager, sharepoint_admin_url, list_id, page_size=5000, ignore_url_contains=None):
    """
    Get all sites from the Tenant Admin Aggregated Sites List with proper pagination
    """
    if ignore_url_contains is None:
        ignore_url_contains = []
    
    print(f"\n=== Retrieving SharePoint Sites from Admin List ===")
    print(f"Using page size: {page_size}")
    print(f"Ignoring sites with URL containing: {ignore_url_contains if ignore_url_contains else 'None'}")
    
    all_sites = []
    ignored_count = 0
    processed_count = 0
    batch_count = 0
    total_sites = 0
    has_more_pages = True
    failed_pages = 0
    max_failures = 3
    
    # First, get the total count of items in the list
    try:
        count_endpoint = f"{sharepoint_admin_url}/_api/Web/Lists(guid'{list_id}')/ItemCount"
        count_data = make_sharepoint_request_with_retry(token_manager, count_endpoint, max_retries=2)
        total_items = count_data.get('value', 0)
        print(f"Total items in list: {total_items}")
    except Exception as e:
        print(f"Warning: Could not get total item count: {e}")
        total_items = "Unknown"
    
    # Start with the first page using $skiptoken=0
    endpoint = f"{sharepoint_admin_url}/_api/Web/Lists(guid'{list_id}')/items?$skiptoken=0&$top={page_size}"
    
    while has_more_pages:
        batch_count += 1
        try:
            print(f"\nProcessing batch {batch_count}...")
            
            data = make_sharepoint_request_with_retry(token_manager, endpoint, max_retries=3)
            
            current_batch = data.get('value', [])
            batch_size = len(current_batch)
            
            print(f"  Found {batch_size} sites in this batch")
            
            if batch_size == 0:
                print("  No items in this batch, stopping pagination")
                break
            
            # Process each site in the batch
            for idx, item in enumerate(current_batch, 1):
                total_sites += 1
                
                # Get title and URL for checking
                title = item.get('Title', '')
                site_url = item.get('SiteUrl', '')
                
                # Check if site should be ignored based on URL
                if should_ignore_site(site_url, ignore_url_contains):
                    ignored_count += 1
                    if ignored_count % 10 == 0 or ignored_count == 1:
                        print(f"  [IGNORED] {ignored_count}: {site_url[:80]}...")
                    continue
                
                processed_count += 1
                
                # Show progress every 100 sites
                if processed_count % 100 == 0 or processed_count == 1:
                    print(f"  [{processed_count}] Processing: {title[:50]}...")
                
                try:
                    # Extract only the fields we need
                    site_info = {
                        'id': item.get('Id', ''),
                        'item_id': item.get('ID', ''),
                        'title': title,
                        'site_url': site_url,
                        'site_id': item.get('SiteId', ''),
                        'template_name': item.get('TemplateName', ''),
                        
                        # Storage
                        'storage_used_gb': round(safe_float_conversion(item.get('StorageUsed', 0)) / (1024**3), 2),
                        'storage_quota_gb': round(safe_float_conversion(item.get('StorageQuota', 0)) / (1024**3), 2),
                        
                        # Dates
                        'created': item.get('Created', ''),
                        'modified': item.get('Modified', ''),
                        'time_created': item.get('TimeCreated', ''),
                        'time_deleted': item.get('TimeDeleted', ''),
                        'last_activity': item.get('LastActivityOn', ''),
                        
                        # Users
                        'created_by': item.get('CreatedBy', ''),
                        'created_by_email': item.get('CreatedByEmail', ''),
                        'site_owner_email': item.get('SiteOwnerEmail', ''),
                        
                        # Files
                        'num_of_files': item.get('NumOfFiles', 0),
                        
                        # Archive Status
                        'archive_status': item.get('ArchiveStatus', '')
                    }
                    
                    all_sites.append(site_info)
                    
                except Exception as e:
                    print(f"  ⚠️ Error processing item {total_sites} (ID: {item.get('Id', 'Unknown')}): {str(e)[:100]}")
                    continue
            
            # Check for next link for pagination
            next_link = data.get('odata.nextLink')
            if next_link:
                print(f"  ✓ Batch {batch_count} complete. Next page available.")
                endpoint = next_link
                time.sleep(0.5)
                failed_pages = 0
            else:
                print(f"  ✓ No more pages available. All items retrieved.")
                has_more_pages = False
                
        except Exception as e:
            print(f"❌ Error processing batch {batch_count}: {str(e)}")
            failed_pages += 1
            
            if "view threshold" in str(e).lower() or "5000" in str(e) or "throttle" in str(e).lower():
                print("  ⚠️ List view threshold or throttling issue detected!")
                
                if page_size > 1000:
                    page_size = 1000
                    endpoint = f"{sharepoint_admin_url}/_api/Web/Lists(guid'{list_id}')/items?$skiptoken=0&$top={page_size}"
                    print(f"  Retrying with page size: {page_size}")
                    continue
                elif page_size > 500:
                    page_size = 500
                    endpoint = f"{sharepoint_admin_url}/_api/Web/Lists(guid'{list_id}')/items?$skiptoken=0&$top={page_size}"
                    print(f"  Retrying with page size: {page_size}")
                    continue
                elif page_size > 100:
                    page_size = 100
                    endpoint = f"{sharepoint_admin_url}/_api/Web/Lists(guid'{list_id}')/items?$skiptoken=0&$top={page_size}"
                    print(f"  Retrying with page size: {page_size}")
                    continue
            
            if failed_pages >= max_failures:
                print(f"  Too many failures ({failed_pages}). Stopping pagination.")
                has_more_pages = False
            elif all_sites:
                print(f"  Continuing with next page...")
                if 'next_link' in locals() and next_link:
                    endpoint = next_link
                    continue
                else:
                    has_more_pages = False
    
    print(f"\n{'='*50}")
    print(f"Total sites processed: {processed_count}")
    print(f"Total sites ignored (by URL): {ignored_count}")
    print(f"Total sites retrieved: {len(all_sites)}")
    print(f"Total batches processed: {batch_count}")
    
    return all_sites

def save_to_csv(sites, filename):
    """Save sites data to CSV file"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'ID', 'Item ID', 'Title', 'Site URL', 'Site ID', 'Template Name',
                'Storage Used (GB)', 'Storage Quota (GB)',
                'Created', 'Modified', 'Time Created', 'Time Deleted', 'Last Activity',
                'Created By', 'Created By Email', 'Site Owner Email',
                'Number of Files', 'Archive Status'
            ]
            
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for site in sites:
                writer.writerow({
                    'ID': site.get('id', ''),
                    'Item ID': site.get('item_id', ''),
                    'Title': site.get('title', ''),
                    'Site URL': site.get('site_url', ''),
                    'Site ID': site.get('site_id', ''),
                    'Template Name': site.get('template_name', ''),
                    'Storage Used (GB)': site.get('storage_used_gb', 0),
                    'Storage Quota (GB)': site.get('storage_quota_gb', 0),
                    'Created': site.get('created', ''),
                    'Modified': site.get('modified', ''),
                    'Time Created': site.get('time_created', ''),
                    'Time Deleted': site.get('time_deleted', ''),
                    'Last Activity': site.get('last_activity', ''),
                    'Created By': site.get('created_by', ''),
                    'Created By Email': site.get('created_by_email', ''),
                    'Site Owner Email': site.get('site_owner_email', ''),
                    'Number of Files': site.get('num_of_files', 0),
                    'Archive Status': site.get('archive_status', '')
                })
        
        print(f"\n✅ CSV report saved to: {filename}")
        
    except Exception as e:
        print(f"Error saving CSV: {str(e)}")

def generate_filename(tenant_name):
    """Generate filename with tenant name and current timestamp"""
    tenant_clean = tenant_name.split('.')[0] if '.' in tenant_name else tenant_name
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{tenant_clean}_sharepoint_sites_report_{timestamp}.csv"

def process_tenant(tenant_config, tenant_key, page_size, max_retries):
    """Process a single tenant"""
    print(f"\n{'='*60}")
    print(f"PROCESSING TENANT: {tenant_key.upper()}")
    print(f"{'='*60}")
    
    tenant_name = tenant_config.get('name')
    app_id = tenant_config.get('app_id')
    certificate_path = tenant_config.get('cert_path')
    private_key_path = tenant_config.get('key_path')
    sharepoint_admin_url = tenant_config.get('sharepoint_admin_url')
    list_id = tenant_config.get('list_id')
    ignore_url_contains = tenant_config.get('ignore_url_contains', [])
    
    print(f"Configuration loaded for {tenant_key}:")
    print(f"  Tenant: {tenant_name}")
    print(f"  SharePoint Admin URL: {sharepoint_admin_url}")
    print(f"  List ID: {list_id}")
    print(f"  Ignore URL keywords: {ignore_url_contains if ignore_url_contains else 'None'}")
    print(f"  Page Size: {page_size}")
    print(f"  Max Retries: {max_retries}")
    
    # Validate required fields
    if not sharepoint_admin_url:
        print(f"Error: sharepoint_admin_url is required for {tenant_key}")
        return None
    if not list_id:
        print(f"Error: list_id is required for {tenant_key}")
        return None
    
    try:
        # Load certificate and key
        certificate, private_key = load_certificate_and_key(certificate_path, private_key_path)
        print("Certificate and private key loaded successfully")
        
        # Create token manager
        token_manager = SharePointTokenManager(certificate, private_key, tenant_name, app_id, sharepoint_admin_url)
        
        # Get initial token
        initial_token = token_manager.get_token()
        print("SharePoint access token retrieved successfully")
        print(f"  Token expires at: {datetime.fromtimestamp(token_manager.token_expiry_time).strftime('%H:%M:%S')}")
        
        # Get all sites from the admin list
        all_sites = get_all_sites_from_list(
            token_manager, 
            sharepoint_admin_url, 
            list_id, 
            page_size,
            ignore_url_contains
        )
        
        if not all_sites:
            print("No sites found!")
            return None
        
        # Generate filename and save to CSV
        filename = generate_filename(tenant_name)
        save_to_csv(all_sites, filename)
        
        # Print summary
        total_storage = sum(s['storage_used_gb'] for s in all_sites)
        total_quota = sum(s['storage_quota_gb'] for s in all_sites)
        total_files = sum(s['num_of_files'] for s in all_sites)
        deleted_sites = [s for s in all_sites if s.get('time_deleted')]
        
        print(f"\n{'='*50}")
        print(f"SUMMARY - {tenant_key.upper()}")
        print(f"{'='*50}")
        print(f"Total Sites: {len(all_sites)}")
        print(f"  - Active sites: {len(all_sites) - len(deleted_sites)}")
        print(f"  - Soft-deleted sites: {len(deleted_sites)}")
        print(f"Total Storage Used: {total_storage:.2f} GB")
        print(f"Total Storage Quota: {total_quota:.2f} GB")
        print(f"Total Files: {total_files:,}")
        
        if total_quota > 0:
            print(f"Overall Usage: {(total_storage / total_quota) * 100:.2f}%")
        
        # Show top 5 largest sites
        largest_sites = sorted(all_sites, key=lambda x: x['storage_used_gb'], reverse=True)[:5]
        if largest_sites:
            print(f"\nTop 5 Largest Sites by Storage:")
            for i, site in enumerate(largest_sites, 1):
                print(f"  {i}. {site['title']}: {site['storage_used_gb']:.2f} GB")
        
        print(f"\n{'='*50}")
        print(f"✅ {tenant_key.upper()} completed successfully!")
        print(f"{'='*50}")
        
        return all_sites
        
    except Exception as e:
        print(f"An error occurred processing {tenant_key}: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def select_tenant(tenants):
    """Display menu and get user selection for tenant"""
    print("\n" + "="*60)
    print("SELECT TENANT TO PROCESS")
    print("="*60)
    
    tenant_keys = list(tenants.keys())
    for i, key in enumerate(tenant_keys, 1):
        print(f"{i}. {key.upper()} - {tenants[key].get('name', 'Unknown')}")
    
    print(f"{len(tenant_keys) + 1}. All Tenants")
    print(f"{len(tenant_keys) + 2}. Exit")
    print("-"*60)
    
    while True:
        try:
            choice = input(f"Enter your choice (1-{len(tenant_keys) + 2}): ").strip()
            choice_int = int(choice)
            if 1 <= choice_int <= len(tenant_keys) + 2:
                return choice_int
            else:
                print(f"Invalid choice. Please enter a number between 1 and {len(tenant_keys) + 2}.")
        except ValueError:
            print("Invalid input. Please enter a number.")
        except KeyboardInterrupt:
            print("\nExiting...")
            return len(tenant_keys) + 2

def main():
    # Load configuration
    config = load_config("config.json")
    
    # Get tenant configurations
    tenants = config.get('tenants', {})
    
    if not tenants:
        print("Error: No tenants found in configuration file.")
        print("Please add tenant configurations to config.json")
        return
    
    # Get global settings
    page_size = config.get('page_size', 5000)
    max_retries = config.get('max_retries', 3)
    
    # Get user selection
    choice = select_tenant(tenants)
    tenant_keys = list(tenants.keys())
    
    # Process based on selection
    if choice == len(tenant_keys) + 2:  # Exit
        print("Exiting...")
        return
    
    if choice == len(tenant_keys) + 1:  # All Tenants
        print(f"\n{'#'*60}")
        print(f"PROCESSING ALL TENANTS")
        print(f"{'#'*60}")
        
        for tenant_key in tenant_keys:
            process_tenant(tenants[tenant_key], tenant_key, page_size, max_retries)
            print(f"\n{'#'*60}\n")
        
        print(f"\n{'#'*60}")
        print(f"✅ ALL TENANTS COMPLETED SUCCESSFULLY!")
        print(f"{'#'*60}")
    else:
        # Process single tenant
        tenant_index = choice - 1
        if tenant_index < len(tenant_keys):
            tenant_key = tenant_keys[tenant_index]
            process_tenant(tenants[tenant_key], tenant_key, page_size, max_retries)
        else:
            print(f"Error: Invalid tenant selection")

if __name__ == "__main__":
    main()
