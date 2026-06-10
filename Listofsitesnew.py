import csv
import json
import uuid
import base64
import time
import requests
import os
from datetime import datetime, timedelta
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
        
        # Set default values for optional parameters
        config.setdefault('sharepoint_url', None)
        config.setdefault('output_prefix', 'sharepoint_sites')
        config.setdefault('page_size', 100)
        config.setdefault('preview_count', 5)
        
        return config
    except FileNotFoundError:
        print(f"Error: Configuration file '{config_file}' not found.")
        print("Please create a config.json file with the required parameters.")
        raise
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON format in '{config_file}'.")
        raise
    except Exception as e:
        print(f"Error loading configuration: {str(e)}")
        raise

def load_certificate_and_key(certificate_path, private_key_path):
    """Load certificate and private key from PEM files"""
    try:
        # Load certificate
        with open(certificate_path, "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())

        # Load private key
        with open(private_key_path, "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())

        return certificate, private_key

    except Exception as e:
        print(f"Error loading certificate or private key: {str(e)}")
        raise

def get_jwt_token(certificate, private_key, tenant_name, app_id, scope):
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
            "aud": f"https://login.microsoftonline.com/{tenant_name}/oauth2/v2.0/token",
            "exp": expiration,
            "iss": app_id,
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": app_id
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
    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error: {err}")
        print(f"Response: {response.text}")
        raise
    except Exception as err:
        print(f"Error: {err}")
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
    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error: {err}")
        print(f"Response: {response.text}")
        raise
    except Exception as err:
        print(f"Error making SharePoint request: {err}")
        raise

def get_site_storage_and_last_modified(token, web_url):
    """Get storage usage and last modified time for a specific site using _api/site"""
    try:
        # Remove trailing slash if present
        web_url = web_url.rstrip('/')
        
        storage_info = {
            'storage_used_bytes': 0,
            'storage_used_gb': 0,
            'storage_percentage': 0,
            'storage_quota': 0,
            'storage_quota_gb': 0,
            'last_modified': '',
            'error': None
        }
        
        # Get site info including usage and last modified from _api/site
        try:
            site_endpoint = f"{web_url}/_api/site"
            site_data = make_sharepoint_request(token, site_endpoint)
            
            # Get storage usage
            if 'StorageUsage' in site_data:
                storage_bytes = site_data['StorageUsage']
                storage_info['storage_used_bytes'] = storage_bytes
                storage_info['storage_used_gb'] = round(storage_bytes / (1024**3), 2)
            
            # Get storage quota (allocated)
            if 'StorageQuota' in site_data:
                quota_bytes = site_data['StorageQuota'] * 1024  # Convert MB to bytes
                storage_info['storage_quota'] = quota_bytes
                storage_info['storage_quota_gb'] = round(quota_bytes / (1024**3), 2)
            
            # Calculate percentage if both values exist
            if storage_info['storage_used_bytes'] > 0 and storage_info['storage_quota'] > 0:
                storage_info['storage_percentage'] = round(
                    (storage_info['storage_used_bytes'] / storage_info['storage_quota']) * 100, 2
                )
            
            # Get last modified time
            if 'LastModified' in site_data:
                storage_info['last_modified'] = site_data['LastModified']
            else:
                storage_info['last_modified'] = 'Unknown'
                
        except Exception as e:
            storage_info['error'] = f"Site API error: {str(e)}"
        
        return storage_info
        
    except Exception as e:
        return {
            'storage_used_bytes': 0,
            'storage_used_gb': 0,
            'storage_percentage': 0,
            'storage_quota': 0,
            'storage_quota_gb': 0,
            'last_modified': 'Error',
            'error': str(e)
        }

def get_all_sites_with_pagination(token, sharepoint_url, page_size=100):
    """Get all sites with proper pagination handling"""
    print(f"\n=== Getting All Sites with Pagination (page size: {page_size}) ===")
    
    all_sites = []
    sharepoint_sites = []
    personal_sites = []
    
    # Initial endpoint
    endpoint = f"{sharepoint_url}/_api/v2.0/sites?$top={page_size}"
    batch_count = 0
    total_sites_processed = 0
    
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
                
                # Get storage and last modified info for each site
                print(f"    Fetching details for: {site.get('title', site.get('name'))}")
                storage_info = get_site_storage_and_last_modified(token, web_url)
                
                # Add rate limiting to avoid throttling
                time.sleep(0.3)
                
                site_info = {
                    'id': site.get('id'),
                    'name': site.get('name'),
                    'title': site.get('title'),
                    'webUrl': web_url,
                    'createdDateTime': site.get('createdDateTime'),
                    'isPersonalSite': site.get('isPersonalSite', False),
                    'dataLocationCode': site.get('dataLocationCode'),
                    'siteCollection': site.get('siteCollection', {}),
                    'template': template_name,
                    'sensitivityLabel': site.get('sensitivityLabel', {}),
                    'storage_used_bytes': storage_info['storage_used_bytes'],
                    'storage_used_gb': storage_info['storage_used_gb'],
                    'storage_percentage': storage_info['storage_percentage'],
                    'storage_quota': storage_info['storage_quota'],
                    'storage_quota_gb': storage_info['storage_quota_gb'],
                    'last_modified': storage_info['last_modified'],
                    'storage_error': storage_info.get('error')
                }
                
                all_sites.append(site_info)
                total_sites_processed += 1
                
                # Separate personal sites from SharePoint sites
                if site.get('isPersonalSite', False):
                    site_info['type'] = 'Personal Site'
                    personal_sites.append(site_info)
                    print(f"    Personal Site: {site.get('title', site.get('name'))} - Storage: {storage_info['storage_used_gb']} GB / {storage_info['storage_quota_gb']} GB")
                else:
                    site_info['type'] = 'SharePoint Site'
                    sharepoint_sites.append(site_info)
                    print(f"    SharePoint Site: {site.get('title', site.get('name'))} - Storage: {storage_info['storage_used_gb']} GB / {storage_info['storage_quota_gb']} GB")
            
            # Check for next link for pagination
            endpoint = None
            if '@odata.nextLink' in sites_data:
                endpoint = sites_data['@odata.nextLink']
                print(f"  Next page available")
            else:
                print("  No more pages available")
                
        except Exception as e:
            print(f"Error getting sites batch {batch_count}: {str(e)}")
            break
    
    print(f"\nTotal sites retrieved: {len(all_sites)}")
    print(f"SharePoint sites: {len(sharepoint_sites)}")
    print(f"Personal sites: {len(personal_sites)}")
    print(f"Total sites processed with storage info: {total_sites_processed}")
    
    return {
        'all_sites': all_sites,
        'sharepoint_sites': sharepoint_sites,
        'personal_sites': personal_sites
    }

def save_sites_to_file(all_sites_data, filename):
    """Save all sites data to a JSON file as backup"""
    try:
        with open(filename, "w") as f:
            json.dump(all_sites_data, f, indent=2, default=str)
        print(f"JSON backup saved to {filename}")
    except Exception as e:
        print(f"Error saving JSON backup: {str(e)}")

def extract_site_id(full_site_id):
    """Extract the middle GUID from the full site ID"""
    try:
        # Split by comma and get the middle part (index 1)
        # Format: hostname,site_id,web_id
        parts = full_site_id.split(',')
        if len(parts) >= 2:
            return parts[1]  # Return the site GUID
        return full_site_id
    except Exception:
        return full_site_id

def save_sites_to_csv(all_sites_data, filename):
    """Save sites data to CSV with storage and last modified columns"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Site Name', 'Creation Date', 'Last Modified', 'Is Personal Site', 
                         'Web URL', 'Site ID', 'Template', 'Storage Used (GB)', 
                         'Storage Quota (GB)', 'Storage Used (%)', 'Storage Error']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            # Write header
            writer.writeheader()
            
            # Write data for all sites
            for site in all_sites_data.get('all_sites', []):
                # Use title if available, otherwise use name
                site_name = site.get('title') or site.get('name', 'Unknown')
                creation_date = site.get('createdDateTime', '')
                last_modified = site.get('last_modified', '')
                is_personal = site.get('isPersonalSite', False)
                web_url = site.get('webUrl', '')
                site_id = extract_site_id(site.get('id', ''))
                template = site.get('template', 'Unknown')
                storage_used_gb = site.get('storage_used_gb', 0)
                storage_quota_gb = site.get('storage_quota_gb', 0)
                storage_percentage = site.get('storage_percentage', 0)
                storage_error = site.get('storage_error', '')
                
                writer.writerow({
                    'Site Name': site_name,
                    'Creation Date': creation_date,
                    'Last Modified': last_modified,
                    'Is Personal Site': is_personal,
                    'Web URL': web_url,
                    'Site ID': site_id,
                    'Template': template,
                    'Storage Used (GB)': storage_used_gb,
                    'Storage Quota (GB)': storage_quota_gb,
                    'Storage Used (%)': storage_percentage,
                    'Storage Error': storage_error
                })
        
        print(f"\nCSV report saved to {filename}")
        
    except Exception as e:
        print(f"Error saving CSV report: {str(e)}")

def save_filtered_csv(all_sites_data, is_personal_filter, filename):
    """Save filtered CSV (either SharePoint sites or Personal sites only)"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Site Name', 'Creation Date', 'Last Modified', 'Is Personal Site', 
                         'Web URL', 'Site ID', 'Template', 'Storage Used (GB)', 
                         'Storage Quota (GB)', 'Storage Used (%)', 'Storage Error']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            # Write header
            writer.writeheader()
            
            # Write filtered data
            filtered_count = 0
            for site in all_sites_data.get('all_sites', []):
                if site.get('isPersonalSite', False) == is_personal_filter:
                    site_name = site.get('title') or site.get('name', 'Unknown')
                    creation_date = site.get('createdDateTime', '')
                    last_modified = site.get('last_modified', '')
                    is_personal = site.get('isPersonalSite', False)
                    web_url = site.get('webUrl', '')
                    site_id = extract_site_id(site.get('id', ''))
                    template = site.get('template', 'Unknown')
                    storage_used_gb = site.get('storage_used_gb', 0)
                    storage_quota_gb = site.get('storage_quota_gb', 0)
                    storage_percentage = site.get('storage_percentage', 0)
                    storage_error = site.get('storage_error', '')
                    
                    writer.writerow({
                        'Site Name': site_name,
                        'Creation Date': creation_date,
                        'Last Modified': last_modified,
                        'Is Personal Site': is_personal,
                        'Web URL': web_url,
                        'Site ID': site_id,
                        'Template': template,
                        'Storage Used (GB)': storage_used_gb,
                        'Storage Quota (GB)': storage_quota_gb,
                        'Storage Used (%)': storage_percentage,
                        'Storage Error': storage_error
                    })
                    filtered_count += 1
        
        site_type = "Personal" if is_personal_filter else "SharePoint"
        print(f"{site_type} sites CSV saved to {filename} ({filtered_count} sites)")
        
    except Exception as e:
        print(f"Error saving filtered CSV: {str(e)}")

def print_csv_preview(all_sites_data, preview_count=5):
    """Print a preview of the CSV data"""
    print(f"\n=== CSV Report Preview (First {preview_count} Sites) ===")
    print("Site Name | Creation Date | Last Modified | Storage (GB) | Quota (GB) | Usage % | Web URL")
    print("-" * 120)
    
    count = 0
    for site in all_sites_data.get('all_sites', []):
        if count >= preview_count:
            break
            
        site_name = site.get('title') or site.get('name', 'Unknown')
        creation_date = site.get('createdDateTime', '')[:10] if site.get('createdDateTime') else 'N/A'
        last_modified = site.get('last_modified', '')[:10] if site.get('last_modified') and site.get('last_modified') != 'Unknown' else 'N/A'
        storage_gb = site.get('storage_used_gb', 0)
        quota_gb = site.get('storage_quota_gb', 0)
        usage_pct = site.get('storage_percentage', 0)
        web_url = site.get('webUrl', '')
        
        # Truncate long URLs for display
        display_url = web_url[:40] + "..." if len(web_url) > 40 else web_url
        
        print(f"{site_name[:20]:<20} | {creation_date:<12} | {last_modified:<12} | {storage_gb:>10.2f} | {quota_gb:>10.2f} | {usage_pct:>6.1f}% | {display_url}")
        count += 1
    
    total_sites = len(all_sites_data.get('all_sites', []))
    if total_sites > preview_count:
        print(f"... and {total_sites - preview_count} more sites")

def generate_summary_report(all_sites_data):
    """Generate and print a summary report"""
    print(f"\n=== DETAILED SUMMARY REPORT ===")
    
    all_sites = all_sites_data.get('all_sites', [])
    sharepoint_sites = [site for site in all_sites if not site.get('isPersonalSite', False)]
    personal_sites = [site for site in all_sites if site.get('isPersonalSite', False)]
    
    print(f"Total Sites Found: {len(all_sites)}")
    print(f"SharePoint Sites: {len(sharepoint_sites)}")
    print(f"Personal Sites: {len(personal_sites)}")
    
    # Storage summary
    total_storage_used = sum(site.get('storage_used_gb', 0) for site in all_sites)
    total_storage_quota = sum(site.get('storage_quota_gb', 0) for site in all_sites)
    
    print(f"\nStorage Summary:")
    print(f"  Total Storage Used: {total_storage_used:.2f} GB")
    print(f"  Total Storage Quota: {total_storage_quota:.2f} GB")
    if total_storage_quota > 0:
        print(f"  Overall Usage: {(total_storage_used / total
