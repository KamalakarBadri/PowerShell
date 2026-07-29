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
    "tenant": "yourtenant.onmicrosoft.com",
    "app_id": "your-app-id",
    "cert_path": "cert.pem",
    "key_path": "key.pem",
    "sharepoint_admin_url": "https://yourtenant-admin.sharepoint.com",
    "list_id": "317f59e4-b925-4d1c-884c-c758bf067a6c",
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

def get_all_sites_from_list(token_manager, sharepoint_admin_url, list_id, page_size=5000):
    """
    Get all sites from the Tenant Admin Aggregated Sites List with proper pagination
    Using $skiptoken to handle large lists beyond the 5000 item view threshold
    """
    print(f"\n=== Retrieving SharePoint Sites from Admin List ===")
    print(f"Using page size: {page_size}")
    
    all_sites = []
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
                
                # Show progress every 100 sites
                if total_sites % 100 == 0 or total_sites == 1:
                    print(f"  [{total_sites}] Processing: {item.get('Title', 'Unknown')[:50]}...")
                
                # Extract all important fields from the list item
                site_info = {
                    # Core identifiers
                    'id': item.get('Id', ''),
                    'item_id': item.get('ID', ''),  # Sometimes ID is capitalized
                    'guid': item.get('GUID', ''),
                    
                    # Site information
                    'title': item.get('Title', ''),
                    'site_url': item.get('SiteUrl', ''),
                    'site_id': item.get('SiteId', ''),
                    'template_name': item.get('TemplateName', ''),
                    'site_creation_source': item.get('SiteCreationSource', ''),
                    'site_flags': item.get('SiteFlags', 0),
                    
                    # Storage information
                    'storage_quota_bytes': item.get('StorageQuota', 0),
                    'storage_quota_gb': round(item.get('StorageQuota', 0) / (1024**3), 2) if item.get('StorageQuota') else 0,
                    'storage_used_bytes': item.get('StorageUsed', 0),
                    'storage_used_gb': round(item.get('StorageUsed', 0) / (1024**3), 2) if item.get('StorageUsed') else 0,
                    'storage_used_percentage': float(item.get('StorageUsedPercentage', '0')) * 100 if item.get('StorageUsedPercentage') else 0,
                    
                    # Dates
                    'created': item.get('Created', ''),
                    'modified': item.get('Modified', ''),
                    'time_created': item.get('TimeCreated', ''),
                    'time_deleted': item.get('TimeDeleted', ''),
                    'last_activity': item.get('LastActivityOn', ''),
                    'operation_start_time': item.get('OperationStartTime', ''),
                    
                    # Users and ownership
                    'created_by': item.get('CreatedBy', ''),
                    'created_by_email': item.get('CreatedByEmail', ''),
                    'deleted_by': item.get('DeletedBy', ''),
                    'site_owner_name': item.get('SiteOwnerName', ''),
                    'site_owner_email': item.get('SiteOwnerEmail', ''),
                    
                    # Groups and hubs
                    'group_id': item.get('GroupId', ''),
                    'related_group_id': item.get('RelatedGroupId', ''),
                    'hub_site_id': item.get('HubSiteId', ''),
                    
                    # File and usage statistics
                    'num_of_files': item.get('NumOfFiles', 0),
                    'page_views': item.get('PageViews', 0),
                    'pages_visited': item.get('PagesVisited', 0),
                    'file_viewed_or_edited': item.get('FileViewedOrEdited', 0),
                    
                    # Sharing and permissions
                    'external_sharing': item.get('ExternalSharing', ''),
                    'allow_guest_signin': item.get('AllowGuestUserSignIn', False),
                    'anonymous_link_count': item.get('AnonymousLinkCount', 0),
                    'company_link_count': item.get('CompanyLinkCount', 0),
                    'sensitivity_label': item.get('SensitivityLabel', ''),
                    
                    # State and status
                    'state': item.get('State', 0),
                    'archive_status': item.get('ArchiveStatus', ''),
                    'block_download_policy': item.get('BlockDownloadPolicy', False),
                    'conditional_access_policy': item.get('ConditionalAccessPolicy', 0),
                    'ib_mode': item.get('IBMode', ''),
                    
                    # Channel information
                    'channel_type': item.get('ChannelType', 0),
                    'channel_sites_count': item.get('ChannelSitesCount', 0),
                    'was_segment_applied': item.get('WasSegmentApplied', False),
                    
                    # Locale and timezone
                    'locale_id': item.get('LocaleId', 0),
                    'timezone_id': item.get('TimeZoneId', 0),
                    
                    # Other metadata
                    'content_type_id': item.get('ContentTypeId', ''),
                    'compliance_asset_id': item.get('ComplianceAssetId', ''),
                    'is_authoritative': item.get('IsAuthoritative', False),
                    'initiator': item.get('Initiator', ''),
                    'color_tag': item.get('OData__ColorTag', '')
                }
                
                all_sites.append(site_info)
            
            # Check for next link for pagination
            next_link = data.get('odata.nextLink')
            if next_link:
                print(f"  ✓ Batch {batch_count} complete. Next page available.")
                endpoint = next_link
                # Small delay to avoid rate limiting
                time.sleep(0.5)
                failed_pages = 0  # Reset failure counter on success
            else:
                print(f"  ✓ No more pages available. All items retrieved.")
                has_more_pages = False
                
        except Exception as e:
            print(f"❌ Error processing batch {batch_count}: {str(e)}")
            failed_pages += 1
            
            # Check if it's a view threshold error
            if "view threshold" in str(e).lower() or "5000" in str(e) or "throttle" in str(e).lower():
                print("  ⚠️ List view threshold or throttling issue detected!")
                
                # Try with smaller page size
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
            
            # If too many failures, break
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
    print(f"Total sites retrieved: {len(all_sites)}")
    print(f"Total batches processed: {batch_count}")
    
    # Count deleted sites
    deleted_sites = [s for s in all_sites if s.get('time_deleted')]
    if deleted_sites:
        print(f"⚠️ Found {len(deleted_sites)} sites with TimeDeleted value (soft-deleted sites)")
    
    return all_sites

def save_to_csv(sites, filename):
    """Save sites data to CSV file"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                # IDs
                'ID', 'Item ID', 'GUID',
                
                # Site Information
                'Title', 'Site URL', 'Site ID', 'Template Name', 'Site Creation Source',
                
                # Storage
                'Storage Used (GB)', 'Storage Quota (GB)', 'Storage Used (%)',
                'Storage Used (Bytes)', 'Storage Quota (Bytes)',
                
                # Dates
                'Created', 'Modified', 'Time Created', 'Time Deleted', 
                'Last Activity', 'Operation Start Time',
                
                # Users
                'Created By', 'Created By Email', 'Deleted By',
                'Site Owner Name', 'Site Owner Email',
                
                # Groups and Hubs
                'Group ID', 'Related Group ID', 'Hub Site ID',
                
                # Usage Statistics
                'Number of Files', 'Page Views', 'Pages Visited', 'File Viewed or Edited',
                
                # Sharing
                'External Sharing', 'Allow Guest SignIn', 'Anonymous Link Count',
                'Company Link Count', 'Sensitivity Label',
                
                # State and Status
                'State', 'Archive Status', 'Block Download Policy',
                'Conditional Access Policy', 'IB Mode',
                
                # Channel
                'Channel Type', 'Channel Sites Count', 'Was Segment Applied',
                
                # Locale and Timezone
                'Locale ID', 'Timezone ID',
                
                # Other
                'Content Type ID', 'Compliance Asset ID', 'Is Authoritative',
                'Initiator', 'Color Tag', 'Site Flags'
            ]
            
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for site in sites:
                writer.writerow({
                    'ID': site.get('id', ''),
                    'Item ID': site.get('item_id', ''),
                    'GUID': site.get('guid', ''),
                    
                    'Title': site.get('title', ''),
                    'Site URL': site.get('site_url', ''),
                    'Site ID': site.get('site_id', ''),
                    'Template Name': site.get('template_name', ''),
                    'Site Creation Source': site.get('site_creation_source', ''),
                    
                    'Storage Used (GB)': site.get('storage_used_gb', 0),
                    'Storage Quota (GB)': site.get('storage_quota_gb', 0),
                    'Storage Used (%)': round(site.get('storage_used_percentage', 0), 4),
                    'Storage Used (Bytes)': site.get('storage_used_bytes', 0),
                    'Storage Quota (Bytes)': site.get('storage_quota_bytes', 0),
                    
                    'Created': site.get('created', ''),
                    'Modified': site.get('modified', ''),
                    'Time Created': site.get('time_created', ''),
                    'Time Deleted': site.get('time_deleted', ''),
                    'Last Activity': site.get('last_activity', ''),
                    'Operation Start Time': site.get('operation_start_time', ''),
                    
                    'Created By': site.get('created_by', ''),
                    'Created By Email': site.get('created_by_email', ''),
                    'Deleted By': site.get('deleted_by', ''),
                    'Site Owner Name': site.get('site_owner_name', ''),
                    'Site Owner Email': site.get('site_owner_email', ''),
                    
                    'Group ID': site.get('group_id', ''),
                    'Related Group ID': site.get('related_group_id', ''),
                    'Hub Site ID': site.get('hub_site_id', ''),
                    
                    'Number of Files': site.get('num_of_files', 0),
                    'Page Views': site.get('page_views', 0),
                    'Pages Visited': site.get('pages_visited', 0),
                    'File Viewed or Edited': site.get('file_viewed_or_edited', 0),
                    
                    'External Sharing': site.get('external_sharing', ''),
                    'Allow Guest SignIn': 'Yes' if site.get('allow_guest_signin') else 'No',
                    'Anonymous Link Count': site.get('anonymous_link_count', 0),
                    'Company Link Count': site.get('company_link_count', 0),
                    'Sensitivity Label': site.get('sensitivity_label', ''),
                    
                    'State': site.get('state', 0),
                    'Archive Status': site.get('archive_status', ''),
                    'Block Download Policy': 'Yes' if site.get('block_download_policy') else 'No',
                    'Conditional Access Policy': site.get('conditional_access_policy', 0),
                    'IB Mode': site.get('ib_mode', ''),
                    
                    'Channel Type': site.get('channel_type', 0),
                    'Channel Sites Count': site.get('channel_sites_count', 0),
                    'Was Segment Applied': 'Yes' if site.get('was_segment_applied') else 'No',
                    
                    'Locale ID': site.get('locale_id', 0),
                    'Timezone ID': site.get('timezone_id', 0),
                    
                    'Content Type ID': site.get('content_type_id', ''),
                    'Compliance Asset ID': site.get('compliance_asset_id', ''),
                    'Is Authoritative': 'Yes' if site.get('is_authoritative') else 'No',
                    'Initiator': site.get('initiator', ''),
                    'Color Tag': site.get('color_tag', ''),
                    'Site Flags': site.get('site_flags', 0)
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
    page_size = config.get('page_size', 5000)
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
        
        # Get all sites from the admin list
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
        
        # Count deleted sites
        deleted_sites = [s for s in all_sites if s.get('time_deleted')]
        
        # Count sites by template
        template_counts = {}
        for site in all_sites:
            template = site.get('template_name', 'Unknown')
            template_counts[template] = template_counts.get(template, 0) + 1
        
        print(f"\n{'='*50}")
        print(f"SUMMARY")
        print(f"{'='*50}")
        print(f"Total Sites: {len(all_sites)}")
        print(f"  - Active sites: {len(all_sites) - len(deleted_sites)}")
        print(f"  - Soft-deleted sites: {len(deleted_sites)}")
        print(f"Total Storage Used: {total_storage:.2f} GB")
        print(f"Total Storage Quota: {total_quota:.2f} GB")
        print(f"Total Files: {total_files:,}")
        
        if total_quota > 0:
            print(f"Overall Usage: {(total_storage / total_quota) * 100:.2f}%")
        
        # Show template distribution
        print(f"\nSite Templates:")
        for template, count in sorted(template_counts.items(), key=lambda x: x[1], reverse=True)[:10]:
            print(f"  {template}: {count} sites")
        
        # Show sites created after 2026-01-15 to verify we got them
        sites_after_jan15 = [s for s in all_sites if s.get('created', '') > '2026-01-15']
        if sites_after_jan15:
            print(f"\n✅ Sites created after 2026-01-15: {len(sites_after_jan15)}")
            print(f"  Latest 5:")
            for site in sorted(sites_after_jan15, key=lambda x: x.get('created', ''), reverse=True)[:5]:
                print(f"    • {site['title']} - Created: {site.get('created', 'Unknown')} (ID: {site.get('id', '')})")
        else:
            print(f"\n⚠️ No sites found created after 2026-01-15")
            created_dates = [s.get('created', '') for s in all_sites if s.get('created')]
            if created_dates:
                latest_date = sorted(created_dates, reverse=True)[0]
                print(f"  Latest created date in retrieved data: {latest_date}")
        
        # Show top 5 largest sites
        largest_sites = sorted(all_sites, key=lambda x: x['storage_used_gb'], reverse=True)[:5]
        if largest_sites:
            print(f"\nTop 5 Largest Sites by Storage:")
            for i, site in enumerate(largest_sites, 1):
                print(f"  {i}. {site['title']}: {site['storage_used_gb']:.2f} GB (ID: {site.get('id', '')})")
        
        # Show deleted sites if any
        if deleted_sites:
            print(f"\n⚠️ Soft-Deleted Sites (with TimeDeleted value):")
            for site in deleted_sites[:5]:
                print(f"  • {site['title']} - Deleted: {site.get('time_deleted', 'Unknown')} (ID: {site.get('id', '')})")
            if len(deleted_sites) > 5:
                print(f"  ... and {len(deleted_sites) - 5} more")
        
        print(f"\n{'='*50}")
        print(f"✅ Script completed successfully!")
        print(f"{'='*50}")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

if __name__ == "__main__":
    main()
