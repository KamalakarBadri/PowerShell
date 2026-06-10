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

def load_config(config_file="config.json"):
    """Load configuration from JSON file"""
    try:
        with open(config_file, 'r') as f:
            config = json.load(f)
        
        config.setdefault('sharepoint_url', None)
        config.setdefault('page_size', 100)
        
        return config
    except FileNotFoundError:
        print(f"Error: Configuration file '{config_file}' not found.")
        print("Please create a config.json file with the required parameters.")
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

def make_sharepoint_request(token, endpoint):
    """Make a request to SharePoint REST API"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.get(endpoint, headers=headers)
        response.raise_for_status()
        return response.json()
    except Exception as err:
        print(f"Error making SharePoint request: {err}")
        raise

def get_site_details(token, web_url):
    """Get storage usage, quota, and last modified for a site"""
    try:
        web_url = web_url.rstrip('/')
        
        site_info = {
            'storage_used_bytes': 0,
            'storage_used_gb': 0,
            'total_quota_bytes': 0,
            'total_quota_gb': 0,
            'storage_percentage': 0,
            'last_modified': ''
        }
        
        # Get storage usage from _api/site/usage
        try:
            usage_endpoint = f"{web_url}/_api/site/usage"
            usage_data = make_sharepoint_request(token, usage_endpoint)
            
            if 'Storage' in usage_data:
                storage_bytes = int(usage_data['Storage'])
                site_info['storage_used_bytes'] = storage_bytes
                site_info['storage_used_gb'] = round(storage_bytes / (1024**3), 2)
            
            if 'StoragePercentageUsed' in usage_data:
                storage_percentage_decimal = float(usage_data['StoragePercentageUsed'])
                site_info['storage_percentage'] = round(storage_percentage_decimal * 100, 2)
                
                # Calculate total quota using formula: Total Quota = Storage / StoragePercentageUsed
                if storage_percentage_decimal > 0 and storage_bytes > 0:
                    total_quota_bytes = storage_bytes / storage_percentage_decimal
                    site_info['total_quota_bytes'] = int(total_quota_bytes)
                    site_info['total_quota_gb'] = round(total_quota_bytes / (1024**3), 2)
        except Exception as e:
            print(f"  Warning: Could not get storage usage for {web_url}: {str(e)}")
        
        # Get last modified from _api/web
        try:
            web_endpoint = f"{web_url}/_api/web"
            web_data = make_sharepoint_request(token, web_endpoint)
            
            if 'LastItemModifiedDate' in web_data:
                site_info['last_modified'] = web_data['LastItemModifiedDate']
            elif 'LastItemUserModifiedDate' in web_data:
                site_info['last_modified'] = web_data['LastItemUserModifiedDate']
            else:
                site_info['last_modified'] = 'Unknown'
        except Exception as e:
            print(f"  Warning: Could not get last modified for {web_url}: {str(e)}")
            site_info['last_modified'] = 'Error'
        
        return site_info
        
    except Exception as e:
        print(f"  Error getting site details: {str(e)}")
        return {
            'storage_used_bytes': 0,
            'storage_used_gb': 0,
            'total_quota_bytes': 0,
            'total_quota_gb': 0,
            'storage_percentage': 0,
            'last_modified': 'Error'
        }

def get_all_sites(token, sharepoint_url, page_size=100):
    """Get all sites with pagination"""
    print(f"\n=== Retrieving SharePoint Sites ===")
    
    all_sites = []
    
    endpoint = f"{sharepoint_url}/_api/v2.0/sites?$top={page_size}"
    batch_count = 0
    
    while endpoint:
        batch_count += 1
        try:
            print(f"Processing batch {batch_count}...")
            sites_data = make_sharepoint_request(token, endpoint)
            
            current_batch = sites_data.get('value', [])
            
            if not current_batch:
                break
            
            print(f"  Found {len(current_batch)} sites in this batch")
            
            for site in current_batch:
                template_name = site.get('template', {}).get('name', 'Unknown')
                web_url = site.get('webUrl')
                site_title = site.get('title', site.get('name'))
                
                print(f"  Fetching details for: {site_title}")
                site_details = get_site_details(token, web_url)
                
                # Add rate limiting
                time.sleep(0.3)
                
                site_info = {
                    'name': site_title,
                    'url': web_url,
                    'created': site.get('createdDateTime', ''),
                    'template': template_name,
                    'is_personal': site.get('isPersonalSite', False),
                    'storage_used_gb': site_details['storage_used_gb'],
                    'total_quota_gb': site_details['total_quota_gb'],
                    'storage_percentage': site_details['storage_percentage'],
                    'last_modified': site_details['last_modified']
                }
                
                all_sites.append(site_info)
                
                site_type = "Personal" if site.get('isPersonalSite', False) else "SharePoint"
                print(f"    {site_type}: {site_details['storage_used_gb']} GB used / {site_details['total_quota_gb']} GB quota ({site_details['storage_percentage']}%)")
            
            # Check for next page
            endpoint = sites_data.get('@odata.nextLink')
            if endpoint:
                print(f"  Next page available")
            else:
                print("  No more pages")
                
        except Exception as e:
            print(f"Error processing batch {batch_count}: {str(e)}")
            break
    
    print(f"\nTotal sites retrieved: {len(all_sites)}")
    return all_sites

def save_to_csv(sites, filename):
    """Save sites data to CSV file"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Site Name', 'Site URL', 'Created Date', 'Template', 'Is Personal Site',
                         'Storage Used (GB)', 'Total Quota (GB)', 'Storage Used (%)', 'Last Modified']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            
            for site in sites:
                writer.writerow({
                    'Site Name': site['name'],
                    'Site URL': site['url'],
                    'Created Date': site['created'],
                    'Template': site['template'],
                    'Is Personal Site': 'Yes' if site['is_personal'] else 'No',
                    'Storage Used (GB)': site['storage_used_gb'],
                    'Total Quota (GB)': site['total_quota_gb'],
                    'Storage Used (%)': site['storage_percentage'],
                    'Last Modified': site['last_modified']
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
    sharepoint_url = config.get('sharepoint_url') or f"https://{tenant_name.split('.')[0]}.sharepoint.com"
    page_size = config.get('page_size', 100)
    
    print(f"Configuration loaded:")
    print(f"  Tenant: {tenant_name}")
    print(f"  SharePoint URL: {sharepoint_url}")
    
    try:
        # Load certificate and key
        certificate, private_key = load_certificate_and_key(certificate_path, private_key_path)
        print("Certificate and private key loaded successfully")
        
        # Get SharePoint token
        sharepoint_scope = f"{sharepoint_url}/.default"
        sharepoint_jwt = get_jwt_token(certificate, private_key, tenant_name, app_id, sharepoint_scope)
        sharepoint_token = get_access_token(sharepoint_jwt, tenant_name, app_id, sharepoint_scope)
        print("SharePoint access token retrieved successfully")
        
        # Get all sites
        all_sites = get_all_sites(sharepoint_token, sharepoint_url, page_size)
        
        # Generate filename and save to CSV
        filename = generate_filename(tenant_name)
        save_to_csv(all_sites, filename)
        
        # Print summary
        sharepoint_count = len([s for s in all_sites if not s['is_personal']])
        personal_count = len([s for s in all_sites if s['is_personal']])
        total_storage = sum(s['storage_used_gb'] for s in all_sites)
        
        print(f"\n=== SUMMARY ===")
        print(f"Total Sites: {len(all_sites)}")
        print(f"  - SharePoint Sites: {sharepoint_count}")
        print(f"  - Personal Sites: {personal_count}")
        print(f"Total Storage Used: {total_storage:.2f} GB")
        
        print(f"\n✅ Script completed successfully!")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

if __name__ == "__main__":
    main()
