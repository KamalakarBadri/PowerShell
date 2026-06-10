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

# Enable debug logging
DEBUG = True

def debug_print(message, data=None):
    """Print debug messages with timestamp"""
    if DEBUG:
        timestamp = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        print(f"[DEBUG {timestamp}] {message}")
        if data:
            if isinstance(data, dict) or isinstance(data, list):
                print(json.dumps(data, indent=2, default=str)[:500])  # Limit output size
            else:
                print(f"  {str(data)[:500]}")

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
        
        debug_print(f"Configuration loaded: {json.dumps(config, indent=2)}")
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
        debug_print(f"Loading certificate from: {certificate_path}")
        with open(certificate_path, "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
        debug_print(f"Certificate loaded successfully. Subject: {certificate.subject}")

        debug_print(f"Loading private key from: {private_key_path}")
        with open(private_key_path, "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())
        debug_print(f"Private key loaded successfully")

        return certificate, private_key

    except Exception as e:
        print(f"Error loading certificate or private key: {str(e)}")
        raise

def get_jwt_token(certificate, private_key, tenant_name, app_id, scope):
    """Generate JWT token using certificate and private key"""
    try:
        debug_print(f"Generating JWT token for tenant: {tenant_name}, scope: {scope}")
        
        # Create JWT timestamp for expiration (5 minutes from now)
        now = int(time.time())
        expiration = now + 300  # 5 minutes
        
        # Get certificate thumbprint (x5t)
        thumbprint = certificate.fingerprint(hashes.SHA1())
        x5t = base64.urlsafe_b64encode(thumbprint).decode('utf-8').replace('=', '')
        debug_print(f"Certificate thumbprint (x5t): {x5t}")
        
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
        
        debug_print(f"JWT Header: {json.dumps(jwt_header)}")
        debug_print(f"JWT Payload: {json.dumps(jwt_payload)}")
        
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
        debug_print(f"JWT token generated successfully (length: {len(jwt)})")
        
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
    
    debug_print(f"Requesting access token from: {url}")
    debug_print(f"Request data: scope={scope}, client_id={app_id}")
    
    try:
        response = requests.post(url, headers=headers, data=data)
        debug_print(f"Response status code: {response.status_code}")
        
        response.raise_for_status()
        token_data = response.json()
        debug_print(f"Access token retrieved successfully. Token type: {token_data.get('token_type')}, Expires in: {token_data.get('expires_in')} seconds")
        
        return token_data["access_token"]
    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error: {err}")
        print(f"Response: {response.text}")
        raise
    except Exception as err:
        print(f"Error: {err}")
        raise

def make_sharepoint_request(token, endpoint, method="GET"):
    """Make a request to SharePoint REST API with debug logging"""
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    
    debug_print(f"\n{'='*60}")
    debug_print(f"SharePoint Request: {method} {endpoint}")
    debug_print(f"Headers: Authorization Bearer {token[:50]}...")
    
    try:
        if method == "GET":
            response = requests.get(endpoint, headers=headers)
        else:
            response = requests.post(endpoint, headers=headers)
        
        debug_print(f"Response Status: {response.status_code}")
        
        if response.status_code == 200:
            response_data = response.json()
            debug_print(f"Response Data Preview: {json.dumps(response_data, indent=2)[:500]}")
            return response_data
        else:
            debug_print(f"Error Response: {response.text}")
            response.raise_for_status()
            
    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error: {err}")
        print(f"Response: {response.text}")
        raise
    except Exception as err:
        print(f"Error making SharePoint request: {err}")
        raise

def get_site_usage_and_last_modified(token, web_url):
    """Get storage usage from _api/site/usage and last modified from _api/web"""
    try:
        # Remove trailing slash if present
        web_url = web_url.rstrip('/')
        
        site_info = {
            'storage_used_bytes': 0,
            'storage_used_gb': 0,
            'storage_percentage': 0,
            'last_modified': '',
            'error': None
        }
        
        debug_print(f"\n--- Fetching details for site: {web_url} ---")
        
        # Get storage usage from _api/site/usage
        try:
            usage_endpoint = f"{web_url}/_api/site/usage"
            debug_print(f"Endpoint 1 (Storage): {usage_endpoint}")
            usage_data = make_sharepoint_request(token, usage_endpoint)
            
            debug_print(f"Raw _api/site/usage response:")
            debug_print(None, usage_data)
            
            if 'Storage' in usage_data:
                storage_bytes = int(usage_data['Storage'])
                site_info['storage_used_bytes'] = storage_bytes
                site_info['storage_used_gb'] = round(storage_bytes / (1024**3), 2)
                debug_print(f"  Storage Used: {storage_bytes} bytes ({site_info['storage_used_gb']} GB)")
            
            if 'StoragePercentageUsed' in usage_data:
                # Convert from decimal to percentage (e.g., 2.93e-06 = 0.000293%)
                site_info['storage_percentage'] = float(usage_data['StoragePercentageUsed']) * 100
                debug_print(f"  Storage Percentage: {usage_data['StoragePercentageUsed']} * 100 = {site_info['storage_percentage']:.6f}%")
                
        except Exception as e:
            site_info['error'] = f"Usage API error: {str(e)}"
            debug_print(f"ERROR in _api/site/usage: {str(e)}")
        
        # Get last modified from _api/web (LastItemModifiedDate)
        try:
            web_endpoint = f"{web_url}/_api/web"
            debug_print(f"Endpoint 2 (Last Modified): {web_endpoint}")
            web_data = make_sharepoint_request(token, web_endpoint)
            
            debug_print(f"Raw _api/web response (selected fields):")
            web_summary = {
                'Title': web_data.get('Title'),
                'LastItemModifiedDate': web_data.get('LastItemModifiedDate'),
                'LastItemUserModifiedDate': web_data.get('LastItemUserModifiedDate'),
                'Language': web_data.get('Language'),
                'IsMultilingual': web_data.get('IsMultilingual')
            }
            debug_print(None, web_summary)
            
            # Use LastItemModifiedDate as the primary last modified timestamp
            if 'LastItemModifiedDate' in web_data:
                site_info['last_modified'] = web_data['LastItemModifiedDate']
                debug_print(f"  Last Modified (LastItemModifiedDate): {site_info['last_modified']}")
            elif 'LastItemUserModifiedDate' in web_data:
                # Fallback to LastItemUserModifiedDate if LastItemModifiedDate not available
                site_info['last_modified'] = web_data['LastItemUserModifiedDate']
                debug_print(f"  Last Modified (LastItemUserModifiedDate): {site_info['last_modified']}")
            else:
                site_info['last_modified'] = 'Unknown'
                debug_print(f"  Last Modified: Not found in response")
                
        except Exception as e:
            site_info['error'] = f"Web API error: {str(e)}"
            site_info['last_modified'] = 'Error'
            debug_print(f"ERROR in _api/web: {str(e)}")
        
        return site_info
        
    except Exception as e:
        debug_print(f"ERROR in get_site_usage_and_last_modified: {str(e)}")
        return {
            'storage_used_bytes': 0,
            'storage_used_gb': 0,
            'storage_percentage': 0,
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
    
    debug_print(f"Initial Graph API endpoint: {endpoint}")
    
    while endpoint:
        batch_count += 1
        try:
            print(f"Processing batch {batch_count}...")
            debug_print(f"Batch {batch_count} - Requesting: {endpoint}")
            
            sites_data = make_sharepoint_request(token, endpoint)
            
            # Debug: Print pagination info
            debug_print(f"Batch {batch_count} - Response contains '@odata.nextLink': {'@odata.nextLink' in sites_data}")
            
            current_batch = sites_data.get('value', [])
            
            if not current_batch:
                debug_print(f"Batch {batch_count} - No sites found in response")
                break
            
            print(f"  Found {len(current_batch)} sites in this batch")
            debug_print(f"Batch {batch_count} - Retrieved {len(current_batch)} sites")
            
            for idx, site in enumerate(current_batch):
                template_name = site.get('template', {}).get('name', 'Unknown')
                web_url = site.get('webUrl')
                site_title = site.get('title', site.get('name'))
                
                debug_print(f"\n--- Processing site {total_sites_processed + 1}: {site_title} ---")
                debug_print(f"Site URL: {web_url}")
                debug_print(f"Site Template: {template_name}")
                debug_print(f"Is Personal Site: {site.get('isPersonalSite', False)}")
                
                # Get usage and last modified info for each site
                print(f"    Fetching details for: {site_title}")
                site_details = get_site_usage_and_last_modified(token, web_url)
                
                # Add rate limiting to avoid throttling
                time.sleep(0.3)
                
                site_info = {
                    'id': site.get('id'),
                    'name': site.get('name'),
                    'title': site_title,
                    'webUrl': web_url,
                    'createdDateTime': site.get('createdDateTime'),
                    'isPersonalSite': site.get('isPersonalSite', False),
                    'dataLocationCode': site.get('dataLocationCode'),
                    'siteCollection': site.get('siteCollection', {}),
                    'template': template_name,
                    'sensitivityLabel': site.get('sensitivityLabel', {}),
                    'storage_used_bytes': site_details['storage_used_bytes'],
                    'storage_used_gb': site_details['storage_used_gb'],
                    'storage_percentage': site_details['storage_percentage'],
                    'last_modified': site_details['last_modified'],
                    'storage_error': site_details.get('error')
                }
                
                all_sites.append(site_info)
                total_sites_processed += 1
                
                # Format last modified for display
                last_mod_display = site_details['last_modified']
                if last_mod_display and last_mod_display != 'Unknown' and last_mod_display != 'Error':
                    last_mod_display = last_mod_display[:19]  # Trim to YYYY-MM-DDTHH:MM:SS
                
                # Separate personal sites from SharePoint sites
                if site.get('isPersonalSite', False):
                    site_info['type'] = 'Personal Site'
                    personal_sites.append(site_info)
                    print(f"    ✅ Personal Site: {site_title} - Storage: {site_details['storage_used_gb']} GB ({site_details['storage_percentage']:.6f}%) | Last Modified: {last_mod_display}")
                else:
                    site_info['type'] = 'SharePoint Site'
                    sharepoint_sites.append(site_info)
                    print(f"    ✅ SharePoint Site: {site_title} - Storage: {site_details['storage_used_gb']} GB ({site_details['storage_percentage']:.6f}%) | Last Modified: {last_mod_display}")
                
                debug_print(f"Site processing complete. Storage: {site_details['storage_used_gb']} GB, Last Modified: {site_details['last_modified']}")
            
            # Check for next link for pagination
            endpoint = None
            if '@odata.nextLink' in sites_data:
                endpoint = sites_data['@odata.nextLink']
                print(f"  Next page available")
                debug_print(f"Next page URL: {endpoint}")
            else:
                print("  No more pages available")
                debug_print(f"Pagination complete. No more pages.")
                
        except Exception as e:
            print(f"Error getting sites batch {batch_count}: {str(e)}")
            debug_print(f"FATAL ERROR in batch {batch_count}: {str(e)}")
            break
    
    print(f"\nTotal sites retrieved: {len(all_sites)}")
    print(f"SharePoint sites: {len(sharepoint_sites)}")
    print(f"Personal sites: {len(personal_sites)}")
    print(f"Total sites processed: {total_sites_processed}")
    
    return {
        'all_sites': all_sites,
        'sharepoint_sites': sharepoint_sites,
        'personal_sites': personal_sites
    }

def save_sites_to_file(all_sites_data, filename):
    """Save all sites data to a JSON file as backup"""
    try:
        debug_print(f"Saving JSON backup to: {filename}")
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
        debug_print(f"Saving CSV report to: {filename}")
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Site Name', 'Creation Date', 'Last Modified', 'Is Personal Site', 
                         'Web URL', 'Site ID', 'Template', 'Storage Used (GB)', 
                         'Storage Used (%)', 'Storage Error']
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
                    'Storage Used (%)': f"{storage_percentage:.6f}",
                    'Storage Error': storage_error
                })
        
        print(f"\nCSV report saved to {filename}")
        
    except Exception as e:
        print(f"Error saving CSV report: {str(e)}")

def save_filtered_csv(all_sites_data, is_personal_filter, filename):
    """Save filtered CSV (either SharePoint sites or Personal sites only)"""
    try:
        debug_print(f"Saving filtered CSV to: {filename} (is_personal={is_personal_filter})")
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Site Name', 'Creation Date', 'Last Modified', 'Is Personal Site', 
                         'Web URL', 'Site ID', 'Template', 'Storage Used (GB)', 
                         'Storage Used (%)', 'Storage Error']
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
                        'Storage Used (%)': f"{storage_percentage:.6f}",
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
    print("Site Name | Creation Date | Last Modified | Storage (GB) | Usage % | Web URL")
    print("-" * 120)
    
    count = 0
    for site in all_sites_data.get('all_sites', []):
        if count >= preview_count:
            break
            
        site_name = site.get('title') or site.get('name', 'Unknown')
        creation_date = site.get('createdDateTime', '')[:10] if site.get('createdDateTime') else 'N/A'
        last_modified = site.get('last_modified', '')[:19] if site.get('last_modified') and site.get('last_modified') != 'Unknown' and site.get('last_modified') != 'Error' else 'N/A'
        storage_gb = site.get('storage_used_gb', 0)
        usage_pct = site.get('storage_percentage', 0)
        web_url = site.get('webUrl', '')
        
        # Truncate long URLs for display
        display_url = web_url[:40] + "..." if len(web_url) > 40 else web_url
        
        print(f"{site_name[:20]:<20} | {creation_date:<12} | {last_modified:<19} | {storage_gb:>10.2f} | {usage_pct:>10.6f} | {display_url}")
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
    
    print(f"\nStorage Summary:")
    print(f"  Total Storage Used: {total_storage_used:.2f} GB")
    
    # Sites with storage usage > 0
    sites_with_storage = [site for site in all_sites if site.get('storage_used_gb', 0) > 0]
    print(f"  Sites with storage used: {len(sites_with_storage)}")
    
    # Top 10 largest sites by storage
    largest_sites = sorted(all_sites, key=lambda x: x.get('storage_used_gb', 0), reverse=True)[:10]
    if largest_sites:
        print(f"\nTop 10 Largest Sites by Storage Usage:")
        for i, site in enumerate(largest_sites, 1):
            print(f"  {i}. {site.get('title', site.get('name'))}: {site.get('storage_used_gb', 0):.2f} GB ({site.get('storage_percentage', 0):.6f}%)")
    
    # Recently modified sites (last 30 days)
    try:
        thirty_days_ago = datetime.now() - timedelta(days=30)
        recently_modified = []
        
        for site in all_sites:
            last_modified = site.get('last_modified')
            if last_modified and last_modified != 'Unknown' and last_modified != 'Error':
                try:
                    # Parse ISO date format
                    modified_dt = datetime.fromisoformat(last_modified.replace('Z', '+00:00'))
                    if modified_dt.replace(tzinfo=None) > thirty_days_ago:
                        recently_modified.append(site)
                except:
                    pass
        
        print(f"\nRecently Modified Sites (Last 30 days): {len(recently_modified)}")
        if recently_modified and len(recently_modified) <= 10:
            for site in recently_modified[:5]:
                print(f"  - {site.get('title', site.get('name'))}: {site.get('last_modified')}")
        
    except Exception as e:
        print(f"\nCould not calculate recent modifications: {str(e)}")
    
    # Template breakdown
    template_breakdown = {}
    for site in all_sites:
        template_name = site.get('template', 'Unknown')
        if template_name not in template_breakdown:
            template_breakdown[template_name] = {'total': 0, 'sharepoint': 0, 'personal': 0, 'storage': 0}
        template_breakdown[template_name]['total'] += 1
        template_breakdown[template_name]['storage'] += site.get('storage_used_gb', 0)
        if site.get('isPersonalSite', False):
            template_breakdown[template_name]['personal'] += 1
        else:
            template_breakdown[template_name]['sharepoint'] += 1
    
    print(f"\nTemplate Distribution:")
    for template, counts in sorted(template_breakdown.items()):
        print(f"  {template}: {counts['total']} sites, {counts['storage']:.2f} GB ({counts['sharepoint']} SharePoint, {counts['personal']} Personal)")

def generate_report_filename(tenant_name, report_type, extension="csv"):
    """Generate filename with tenant name and current timestamp"""
    # Extract tenant name without domain if needed
    tenant_clean = tenant_name.split('.')[0] if '.' in tenant_name else tenant_name
    
    # Get current timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Create filename
    filename = f"{tenant_clean}_{report_type}_{timestamp}.{extension}"
    
    return filename

def main():
    # Load configuration from JSON file
    config = load_config("config.json")
    
    # Set up configuration from config
    tenant_name = config.get('tenant')
    app_id = config.get('app_id')
    certificate_path = config.get('cert_path')
    private_key_path = config.get('key_path')
    sharepoint_url = config.get('sharepoint_url') or f"https://{tenant_name.split('.')[0]}.sharepoint.com"
    output_prefix = config.get('output_prefix')
    page_size = config.get('page_size')
    preview_count = config.get('preview_count')
    
    # Define scopes
    scope_graph = "https://graph.microsoft.com/.default"
    scope_sharepoint = f"{sharepoint_url}/.default"
    
    print(f"Configuration loaded from config.json:")
    print(f"  Tenant: {tenant_name}")
    print(f"  App ID: {app_id}")
    print(f"  Certificate: {certificate_path}")
    print(f"  Private Key: {private_key_path}")
    print(f"  SharePoint URL: {sharepoint_url}")
    print(f"  Output Prefix: {output_prefix}")
    print(f"  Page Size: {page_size}")
    print(f"  Preview Count: {preview_count}")
    print(f"  Debug Mode: {DEBUG}")
    
    try:
        # Load certificate and private key
        certificate, private_key = load_certificate_and_key(certificate_path, private_key_path)
        print("Certificate and private key loaded successfully")
        
        # Get Graph token
        graph_jwt = get_jwt_token(certificate, private_key, tenant_name, app_id, scope_graph)
        print("Generated Graph JWT")
        graph_token = get_access_token(graph_jwt, tenant_name, app_id, scope_graph)
        print("Graph access token retrieved successfully")
        
        # Get SharePoint token
        sharepoint_jwt = get_jwt_token(certificate, private_key, tenant_name, app_id, scope_sharepoint)
        print("Generated SharePoint JWT")
        sharepoint_token = get_access_token(sharepoint_jwt, tenant_name, app_id, scope_sharepoint)
        print("SharePoint access token retrieved successfully")
        
        # Save tokens to file with timestamp
        tokens_filename = generate_report_filename(tenant_name, "tokens", "json")
        tokens = {
            "graph_token": graph_token,
            "sharepoint_token": sharepoint_token,
            "timestamp": datetime.now().isoformat(),
            "tenant": tenant_name
        }
        
        with open(tokens_filename, "w") as f:
            json.dump(tokens, f, indent=2)
        
        print(f"Tokens saved to {tokens_filename}")
        
        # Now get all SharePoint sites and personal sites using SharePoint REST API
        print("\n" + "="*50)
        print("RETRIEVING SHAREPOINT SITES AND PERSONAL SITES")
        print("="*50)
        
        all_sites_data = {
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
            "tenant": tenant_name,
            'all_sites': [],
            'sharepoint_sites': [],
            'personal_sites': []
        }
        
        # Use the pagination method to get all sites
        sites_result = get_all_sites_with_pagination(sharepoint_token, sharepoint_url, page_size)
        all_sites_data["all_sites"] = sites_result["all_sites"]
        all_sites_data["sharepoint_sites"] = sites_result["sharepoint_sites"]
        all_sites_data["personal_sites"] = sites_result["personal_sites"]
        
        # Generate filenames with tenant name and timestamp
        report_csv = generate_report_filename(tenant_name, "complete_report", "csv")
        sharepoint_csv = generate_report_filename(tenant_name, "sharepoint_only", "csv")
        personal_csv = generate_report_filename(tenant_name, "personal_only", "csv")
        backup_json = generate_report_filename(tenant_name, "backup", "json")
        
        # Save all data to files
        save_sites_to_file(all_sites_data, backup_json)  # JSON backup
        save_sites_to_csv(all_sites_data, report_csv)   # Main CSV report
        
        # Also create separate CSV files for SharePoint and Personal sites
        save_filtered_csv(all_sites_data, False, sharepoint_csv)
        save_filtered_csv(all_sites_data, True, personal_csv)
        
        # Print CSV preview
        print_csv_preview(all_sites_data, preview_count)
        
        # Generate detailed summary
        generate_summary_report(all_sites_data)
        
        print(f"\n✅ Script completed successfully!")
        print(f"📁 Output files created:")
        print(f"   - {report_csv} (Complete report with all sites)")
        print(f"   - {sharepoint_csv} (SharePoint sites only)")
        print(f"   - {personal_csv} (Personal/OneDrive sites only)")
        print(f"   - {backup_json} (JSON backup)")
        print(f"   - {tokens_filename} (Access tokens)")
        
        if DEBUG:
            print(f"\n💡 Debug mode was enabled. Check console output for detailed endpoint and response logs.")
        
        return all_sites_data
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

if __name__ == "__main__":
    main()
