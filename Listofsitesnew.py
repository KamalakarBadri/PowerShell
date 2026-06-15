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
        
        config.setdefault('page_size', 100)
        config.setdefault('max_retries', 3)
        config.setdefault('max_concurrent_requests', 10)  # New setting for concurrent requests
        
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

def get_site_metadata(token_manager, site_url):
    """Get site metadata including LastItemModifiedDate and LastItemUserModifiedDate"""
    try:
        endpoint = f"{site_url.rstrip('/')}/_api/web"
        data = make_sharepoint_request_with_retry(token_manager, endpoint, max_retries=2)
        
        return {
            'last_item_modified_date': data.get('LastItemModifiedDate', ''),
            'last_item_user_modified_date': data.get('LastItemUserModifiedDate', '')
        }
    except Exception as e:
        print(f"  ⚠️ Error getting metadata for {site_url}: {str(e)[:100]}")
        return {
            'last_item_modified_date': 'Error',
            'last_item_user_modified_date': 'Error'
        }

def get_all_sites_from_list(token_manager, sharepoint_admin_url, list_id, page_size=100, max_concurrent=10):
    """Get all sites from the Tenant Admin Aggregated Sites List with additional metadata"""
    print(f"\n=== Retrieving SharePoint Sites from Admin List ===")
    
    all_sites = []
    
    # Base endpoint for the list items
    base_endpoint = f"{sharepoint_admin_url}/_api/Web/Lists(guid'{list_id}')/items"
    
    # Initial request with top parameter
    endpoint = f"{base_endpoint}?$top={page_size}"
    batch_count = 0
    total_sites = 0
    
    while endpoint:
        batch_count += 1
        try:
            print(f"Processing batch {batch_count}...")
            data = make_sharepoint_request_with_retry(token_manager, endpoint)
            
            current_batch = data.get('value', [])
            
            if not current_batch:
                break
            
            print(f"  Found {len(current_batch)} sites in this batch")
            
            # Process each site in the batch
            for idx, item in enumerate(current_batch, 1):
                total_sites += 1
                
                # Extract basic site info
                site_url = item.get('SiteUrl', '')
                
                print(f"  [{total_sites}] Fetching metadata for: {item.get('Title', 'Unknown')}")
                
                # Get additional metadata from the site
                site_metadata = get_site_metadata(token_manager, site_url)
                
                site_info = {
                    # New columns from list item
                    'id': item.get('Id', ''),  # Added Id field
                    'time_deleted': item.get('TimeDeleted', ''),  # Added TimeDeleted field
                    
                    # Existing fields
                    'title': item.get('Title', ''),
                    'site_url': site_url,
                    'site_id': item.get('SiteId', ''),
                    'template_name': item.get('TemplateName', ''),
                    'storage_quota_bytes': item.get('StorageQuota', 0),
                    'storage_quota_gb': round(item.get('StorageQuota', 0) / (1024**3), 2) if item.get('StorageQuota') else 0,
                    'storage_used_bytes': item.get('StorageUsed', 0),
                    'storage_used_gb': round(item.get('StorageUsed', 0) / (1024**3), 2) if item.get('StorageUsed') else 0,
                    'storage_used_percentage': float(item.get('StorageUsedPercentage', '0')) * 100 if item.get('StorageUsedPercentage') else 0,
                    'created': item.get('Created', ''),
                    'created_by': item.get('CreatedBy', ''),
                    'created_by_email': item.get('CreatedByEmail', ''),
                    'modified': item.get('Modified', ''),
                    'last_activity': item.get('LastActivityOn', ''),
                    'num_of_files': item.get('NumOfFiles', 0),
                    'page_views': item.get('PageViews', 0),
                    'pages_visited': item.get('PagesVisited', 0),
                    'external_sharing': item.get('ExternalSharing', ''),
                    'allow_guest_signin': item.get('AllowGuestUserSignIn', False),
                    'group_id': item.get('GroupId', ''),
                    'hub_site_id': item.get('HubSiteId', ''),
                    'state': item.get('State', 0),
                    'time_created': item.get('TimeCreated', ''),
                    'archive_status': item.get('ArchiveStatus', ''),
                    
                    # Fields from site metadata
                    'last_item_modified_date': site_metadata.get('last_item_modified_date', ''),
                    'last_item_user_modified_date': site_metadata.get('last_item_user_modified_date', '')
                }
                
                all_sites.append(site_info)
                
                # Print progress with new fields
                print(f"    ✓ ID: {site_info['id']}")
                if site_info['time_deleted']:
                    print(f"    ⚠️ Time Deleted: {site_info['time_deleted']}")
                print(f"    ✓ Storage: {site_info['storage_used_gb']} GB / {site_info['storage_quota_gb']} GB")
                if site_info['last_item_user_modified_date'] and site_info['last_item_user_modified_date'] != 'Error':
                    print(f"    ✓ Last modified: {site_info['last_item_user_modified_date']}")
            
            # Check for next link for pagination
            endpoint = data.get('odata.nextLink')
            if endpoint:
                print(f"  Next page available")
            else:
                print("  No more pages")
                
        except Exception as e:
            print(f"Error processing batch {batch_count}: {str(e)}")
            break
    
    print(f"\nTotal sites retrieved: {len(all_sites)}")
    
    # Count deleted sites
    deleted_sites = [s for s in all_sites if s.get('time_deleted')]
    if deleted_sites:
        print(f"\n⚠️ Found {len(deleted_sites)} sites with TimeDeleted value (soft-deleted sites)")
    
    return all_sites

def save_to_csv(sites, filename):
    """Save sites data to CSV file"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                # New columns
                'ID', 'Time Deleted',
                
                # Existing columns
                'Title', 'Site URL', 'Site ID', 'Template Name',
                'Storage Used (GB)', 'Storage Quota (GB)', 'Storage Used (%)',
                'Created', 'Created By', 'Created By Email', 'Modified', 'Last Activity',
                'Number of Files', 'Page Views', 'Pages Visited',
                'External Sharing', 'Allow Guest SignIn', 'Group ID', 'Hub Site ID',
                'State', 'Time Created', 'Archive Status',
                'Last Item Modified Date', 'Last Item User Modified Date'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            
            for site in sites:
                writer.writerow({
                    # New columns
                    'ID': site.get('id', ''),
                    'Time Deleted': site.get('time_deleted', ''),
                    
                    # Existing columns
                    'Title': site['title'],
                    'Site URL': site['site_url'],
                    'Site ID': site['site_id'],
                    'Template Name': site['template_name'],
                    'Storage Used (GB)': site['storage_used_gb'],
                    'Storage Quota (GB)': site['storage_quota_gb'],
                    'Storage Used (%)': round(site['storage_used_percentage'], 4),
                    'Created': site['created'],
                    'Created By': site['created_by'],
                    'Created By Email': site['created_by_email'],
                    'Modified': site['modified'],
                    'Last Activity': site['last_activity'],
                    'Number of Files': site['num_of_files'],
                    'Page Views': site['page_views'],
                    'Pages Visited': site['pages_visited'],
                    'External Sharing': site['external_sharing'],
                    'Allow Guest SignIn': 'Yes' if site['allow_guest_signin'] else 'No',
                    'Group ID': site['group_id'],
                    'Hub Site ID': site['hub_site_id'],
                    'State': site['state'],
                    'Time Created': site['time_created'],
                    'Archive Status': site['archive_status'],
                    'Last Item Modified Date': site.get('last_item_modified_date', ''),
                    'Last Item User Modified Date': site.get('last_item_user_modified_date', '')
                })
        
        print(f"\n✅ CSV report saved to: {filename}")
        
    except Exception as e:
        print(f"Error saving CSV: {str(e)}")

def generate_filename(tenant_name):
    """Generate filename with tenant name and current timestamp"""
    tenant_clean = tenant_name.split('.')[0] if '.' in tenant_name else tenant_name
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{tenant_clean}_sharepoint_sites_report_{timestamp}.csv"

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
    
    print(f"Configuration loaded:")
    print(f"  Tenant: {tenant_name}")
    print(f"  SharePoint Admin URL: {sharepoint_admin_url}")
    print(f"  List ID: {list_id}")
    print(f"  Page Size: {page_size}")
    print(f"  Max Retries: {max_retries}")
    
    # Validate required fields
    if not sharepoint_admin_url:
        print("Error: sharepoint_admin_url is required in config.json")
        return
    if not list_id:
        print("Error: list_id is required in config.json")
        print("The list ID is: 317f59e4-b925-4d1c-884c-c758bf067a6c (Tenant Admin Aggregated Sites List)")
        return
    
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
        print(f"  Auto-renewal will happen if token expires during script execution")
        
        # Get all sites from the admin list with additional metadata
        all_sites = get_all_sites_from_list(token_manager, sharepoint_admin_url, list_id, page_size)
        
        if not all_sites:
            print("No sites found!")
            return
        
        # Generate filename and save to CSV
        filename = generate_filename(tenant_name)
        save_to_csv(all_sites, filename)
        
        # Print summary
        total_storage = sum(s['storage_used_gb'] for s in all_sites)
        total_quota = sum(s['storage_quota_gb'] for s in all_sites)
        total_files = sum(s['num_of_files'] for s in all_sites)
        
        # Count sites with recent activity
        sites_with_recent_activity = sum(1 for s in all_sites 
                                        if s.get('last_item_user_modified_date') and 
                                        s.get('last_item_user_modified_date') != 'Error')
        
        # Count deleted sites
        deleted_sites = [s for s in all_sites if s.get('time_deleted')]
        
        print(f"\n=== SUMMARY ===")
        print(f"Total Sites: {len(all_sites)}")
        print(f"  - Active sites: {len(all_sites) - len(deleted_sites)}")
        print(f"  - Soft-deleted sites: {len(deleted_sites)}")
        print(f"Total Storage Used: {total_storage:.2f} GB")
        print(f"Total Storage Quota: {total_quota:.2f} GB")
        print(f"Total Files: {total_files}")
        print(f"Sites with activity data: {sites_with_recent_activity}/{len(all_sites)}")
        
        if total_quota > 0:
            print(f"Overall Usage: {(total_storage / total_quota) * 100:.2f}%")
        
        # Show top 5 largest sites
        largest_sites = sorted(all_sites, key=lambda x: x['storage_used_gb'], reverse=True)[:5]
        if largest_sites:
            print(f"\nTop 5 Largest Sites by Storage:")
            for i, site in enumerate(largest_sites, 1):
                print(f"  {i}. {site['title']}: {site['storage_used_gb']:.2f} GB")
        
        # Show recently modified sites
        recently_modified = [s for s in all_sites 
                            if s.get('last_item_user_modified_date') and 
                            s.get('last_item_user_modified_date') != 'Error'][:5]
        recently_modified.sort(key=lambda x: x.get('last_item_user_modified_date', ''), reverse=True)
        
        if recently_modified:
            print(f"\nTop 5 Recently Modified Sites:")
            for i, site in enumerate(recently_modified[:5], 1):
                modified_date = site.get('last_item_user_modified_date', 'Unknown')
                print(f"  {i}. {site['title']}: {modified_date}")
        
        # Show deleted sites if any
        if deleted_sites:
            print(f"\n⚠️ Soft-Deleted Sites (with TimeDeleted value):")
            for site in deleted_sites[:5]:
                print(f"  • {site['title']} - Deleted: {site.get('time_deleted', 'Unknown')}")
            if len(deleted_sites) > 5:
                print(f"  ... and {len(deleted_sites) - 5} more")
        
        print(f"\n✅ Script completed successfully!")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

if __name__ == "__main__":
    main()
