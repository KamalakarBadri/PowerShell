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
import sys
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
    
    def get_token(self):
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

def load_config(config_file="config.json"):
    try:
        with open(config_file, 'r') as f:
            config = json.load(f)
        
        config.setdefault('page_size', 100)
        config.setdefault('max_retries', 3)
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
    "page_size": 100,
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

def should_ignore_url(site_url, config):
    """Check if URL should be ignored based on pattern"""
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

def safe_str(value, default=''):
    """Safely convert a value to string, handling None"""
    if value is None:
        return default
    return str(value).strip()

def find_onedrive_by_owner_email(token_manager, sharepoint_admin_url, list_id, search_email, config, page_size=100):
    """
    Search for OneDrive site by owner email.
    Searches through paginated results until exact match is found.
    """
    print(f"\n{'='*60}")
    print(f"🔍 SEARCHING FOR ONEDRIVE SITE")
    print(f"{'='*60}")
    print(f"📧 Search Email: {search_email}")
    print(f"📋 List ID: {list_id}")
    print(f"📄 Page Size: {page_size}")
    
    # Clean up the email for comparison
    search_email = search_email.strip().lower()
    
    base_endpoint = f"{sharepoint_admin_url}/_api/Web/Lists(guid'{list_id}')/items"
    endpoint = f"{base_endpoint}?$top={page_size}"
    
    batch_count = 0
    total_scanned = 0
    found_sites = []
    
    while endpoint:
        batch_count += 1
        try:
            print(f"\n  📄 Searching batch {batch_count}...")
            data = make_sharepoint_request(token_manager, endpoint)
            current_batch = data.get('value', [])
            
            if not current_batch:
                print(f"    No more items found")
                break
            
            print(f"    Scanning {len(current_batch)} items in this batch")
            
            for item in current_batch:
                total_scanned += 1
                
                site_url = safe_str(item.get('SiteUrl', ''))
                owner_email = safe_str(item.get('CreatedByEmail', '')).lower()
                
                # Skip if URL should be ignored
                if should_ignore_url(site_url, config):
                    continue
                
                # Skip if not a OneDrive site
                if not is_onedrive_site(site_url):
                    continue
                
                # Check if this is the site we're looking for
                if owner_email == search_email:
                    site_info = {
                        'site_id': safe_str(item.get('SiteId', '')),
                        'site_url': site_url,
                        'title': safe_str(item.get('Title', '')),
                        'owner_email': owner_email,
                        'created_by': safe_str(item.get('CreatedBy', '')),
                        'template_name': safe_str(item.get('TemplateName', '')),
                        'time_created': safe_str(item.get('TimeCreated', '')),
                        'storage_used': float(item.get('StorageUsed', 0) or 0),
                        'archive_status': safe_str(item.get('ArchiveStatus', '')),
                        'time_deleted': safe_str(item.get('TimeDeleted', ''))
                    }
                    found_sites.append(site_info)
                    print(f"\n  ✅ FOUND MATCH!")
                    print(f"    Site URL: {site_url}")
                    print(f"    Title: {site_info['title']}")
                    print(f"    Site ID: {site_info['site_id']}")
            
            # If we found matches, we can stop searching
            if found_sites:
                print(f"\n  🎯 Found {len(found_sites)} matching site(s). Stopping search.")
                break
            
            # Get next page
            endpoint = data.get('odata.nextLink')
            
            # If no next link, we're done
            if not endpoint:
                print(f"\n  📊 Reached end of list. No more pages to search.")
                break
                
        except Exception as e:
            print(f"  ❌ Error processing batch {batch_count}: {str(e)}")
            break
    
    print(f"\n{'='*60}")
    print(f"📊 SEARCH SUMMARY")
    print(f"{'='*60}")
    print(f"Total items scanned: {total_scanned}")
    print(f"Total batches processed: {batch_count}")
    print(f"Matches found: {len(found_sites)}")
    
    if found_sites:
        print(f"\n✅ FOUND {len(found_sites)} ONEDRIVE SITE(S):")
        for i, site in enumerate(found_sites, 1):
            print(f"\n  Site {i}:")
            print(f"    URL: {site['site_url']}")
            print(f"    Title: {site['title']}")
            print(f"    Site ID: {site['site_id']}")
            print(f"    Owner: {site['owner_email']}")
            print(f"    Created: {site['time_created']}")
            if site['storage_used']:
                storage_gb = site['storage_used'] / (1024**3)
                print(f"    Storage: {storage_gb:.2f} GB")
    else:
        print(f"\n❌ No OneDrive site found for email: {search_email}")
        print(f"   Please check that the email is correct and the user has a OneDrive site.")
    
    return found_sites

def main():
    # Check if email was provided as command line argument
    if len(sys.argv) < 2:
        print("❌ Error: Please provide the owner email to search for.")
        print("\nUsage:")
        print("  python find_onedrive_by_email.py user@domain.com")
        print("\nExample:")
        print("  python find_onedrive_by_email.py john.doe@company.com")
        sys.exit(1)
    
    search_email = sys.argv[1]
    
    # Load configuration
    config = load_config("config.json")
    
    tenant_name = config.get('tenant')
    app_id = config.get('app_id')
    certificate_path = config.get('cert_path')
    private_key_path = config.get('key_path')
    sharepoint_admin_url = config.get('sharepoint_admin_url')
    list_id = config.get('list_id')
    page_size = config.get('page_size', 100)
    
    # Validate required config
    if not sharepoint_admin_url:
        print("❌ Error: sharepoint_admin_url is required in config.json")
        return
    if not list_id:
        print("❌ Error: list_id is required in config.json")
        return
    
    try:
        # Load certificate and get token
        print(f"\n🚀 Starting OneDrive search...")
        print(f"📧 Searching for: {search_email}")
        
        certificate, private_key = load_certificate_and_key(certificate_path, private_key_path)
        print("✅ Certificate and private key loaded successfully")
        
        sharepoint_token_manager = SharePointTokenManager(
            certificate, 
            private_key, 
            tenant_name, 
            app_id, 
            sharepoint_admin_url
        )
        
        # Get initial token
        sharepoint_token_manager.get_token()
        print("✅ SharePoint token retrieved successfully")
        
        # Search for the OneDrive site
        found_sites = find_onedrive_by_owner_email(
            sharepoint_token_manager,
            sharepoint_admin_url,
            list_id,
            search_email,
            config,
            page_size
        )
        
        # If found, output just the URL for easy scripting
        if found_sites:
            print(f"\n{'='*60}")
            print("📋 RESULTS FOR SCRIPTING")
            print(f"{'='*60}")
            print(f"\nFirst match URL:")
            print(found_sites[0]['site_url'])
            
            # Optionally output all URLs
            if len(found_sites) > 1:
                print(f"\nAll matching URLs:")
                for site in found_sites:
                    print(site['site_url'])
            
            # Save to file if needed
            output_file = f"onedrive_site_{search_email.replace('@', '_').replace('.', '_')}.txt"
            with open(output_file, 'w') as f:
                f.write(found_sites[0]['site_url'])
            print(f"\n💾 URL saved to: {output_file}")
        else:
            print(f"\n❌ No OneDrive site found for: {search_email}")
            sys.exit(1)
            
    except Exception as e:
        print(f"\n❌ An error occurred: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main()
