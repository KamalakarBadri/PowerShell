{
    "tenant": "geekbyteonline.onmicrosoft.com",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "cert_path": "certificate.pem",
    "key_path": "private_key.pem",
    "sharepoint_url": "https://geekbyteonline.sharepoint.com",
    "output_prefix": "sharepoint_sites",
    "page_size": 100,
    "preview_count": 5
}









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

def get_all_sites_from_sharepoint(token, sharepoint_url):
    """Get all sites using SharePoint REST API v2.0"""
    print("\n=== Getting All Sites from SharePoint API ===")
    
    all_sites = []
    sharepoint_sites = []
    personal_sites = []
    
    # Use the SharePoint REST API endpoint
    sites_endpoint = f"{sharepoint_url}/_api/v2.0/sites"
    
    try:
        sites_data = make_sharepoint_request(token, sites_endpoint)
        
        for site in sites_data.get('value', []):
            template_name = site.get('template', {}).get('name', 'Unknown')
            
            site_info = {
                'id': site.get('id'),
                'name': site.get('name'),
                'title': site.get('title'),
                'webUrl': site.get('webUrl'),
                'createdDateTime': site.get('createdDateTime'),
                'isPersonalSite': site.get('isPersonalSite', False),
                'dataLocationCode': site.get('dataLocationCode'),
                'siteCollection': site.get('siteCollection', {}),
                'template': template_name,
                'sensitivityLabel': site.get('sensitivityLabel', {})
            }
            
            all_sites.append(site_info)
            
            # Separate personal sites from SharePoint sites
            if site.get('isPersonalSite', False):
                site_info['type'] = 'Personal Site'
                personal_sites.append(site_info)
                print(f"Personal Site: {site.get('title', site.get('name'))} - {site.get('webUrl')} - Template: {template_name}")
            else:
                site_info['type'] = 'SharePoint Site'
                sharepoint_sites.append(site_info)
                print(f"SharePoint Site: {site.get('title', site.get('name'))} - {site.get('webUrl')} - Template: {template_name}")
        
        return {
            'all_sites': all_sites,
            'sharepoint_sites': sharepoint_sites,
            'personal_sites': personal_sites
        }
        
    except Exception as e:
        print(f"Error getting sites from SharePoint API: {str(e)}")
        return {
            'all_sites': [],
            'sharepoint_sites': [],
            'personal_sites': []
        }

def get_sites_with_pagination(token, sharepoint_url, page_size=100):
    """Get all sites with pagination support"""
    print(f"\n=== Getting All Sites with Pagination (page size: {page_size}) ===")
    
    all_sites = []
    sharepoint_sites = []
    personal_sites = []
    skip = 0
    
    while True:
        try:
            # Use $top and $skip for pagination
            sites_endpoint = f"{sharepoint_url}/_api/v2.0/sites?$top={page_size}&$skip={skip}"
            sites_data = make_sharepoint_request(token, sites_endpoint)
            
            current_batch = sites_data.get('value', [])
            
            if not current_batch:
                break
            
            print(f"Processing batch {skip // page_size + 1} - {len(current_batch)} sites")
            
            for site in current_batch:
                template_name = site.get('template', {}).get('name', 'Unknown')
                
                site_info = {
                    'id': site.get('id'),
                    'name': site.get('name'),
                    'title': site.get('title'),
                    'webUrl': site.get('webUrl'),
                    'createdDateTime': site.get('createdDateTime'),
                    'isPersonalSite': site.get('isPersonalSite', False),
                    'dataLocationCode': site.get('dataLocationCode'),
                    'siteCollection': site.get('siteCollection', {}),
                    'template': template_name,
                    'sensitivityLabel': site.get('sensitivityLabel', {})
                }
                
                all_sites.append(site_info)
                
                # Separate personal sites from SharePoint sites
                if site.get('isPersonalSite', False):
                    site_info['type'] = 'Personal Site'
                    personal_sites.append(site_info)
                    print(f"  Personal Site: {site.get('title', site.get('name'))} - Template: {template_name}")
                else:
                    site_info['type'] = 'SharePoint Site'
                    sharepoint_sites.append(site_info)
                    print(f"  SharePoint Site: {site.get('title', site.get('name'))} - Template: {template_name}")
            
            skip += page_size
            
            # If we got fewer results than requested, we've reached the end
            if len(current_batch) < page_size:
                break
                
        except Exception as e:
            print(f"Error getting sites batch starting at {skip}: {str(e)}")
            break
    
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
    """Save sites data to CSV with specific columns"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Site Name', 'Creation Date', 'Is Personal Site', 'Web URL', 'Site ID', 'Template']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            # Write header
            writer.writeheader()
            
            # Write data for all sites
            for site in all_sites_data.get('all_sites', []):
                # Use title if available, otherwise use name
                site_name = site.get('title') or site.get('name', 'Unknown')
                creation_date = site.get('createdDateTime', '')
                is_personal = site.get('isPersonalSite', False)
                web_url = site.get('webUrl', '')
                site_id = extract_site_id(site.get('id', ''))
                template = site.get('template', 'Unknown')
                
                writer.writerow({
                    'Site Name': site_name,
                    'Creation Date': creation_date,
                    'Is Personal Site': is_personal,
                    'Web URL': web_url,
                    'Site ID': site_id,
                    'Template': template
                })
        
        print(f"\nCSV report saved to {filename}")
        
    except Exception as e:
        print(f"Error saving CSV report: {str(e)}")

def save_filtered_csv(all_sites_data, is_personal_filter, filename):
    """Save filtered CSV (either SharePoint sites or Personal sites only)"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Site Name', 'Creation Date', 'Is Personal Site', 'Web URL', 'Site ID', 'Template']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            # Write header
            writer.writeheader()
            
            # Write filtered data
            filtered_count = 0
            for site in all_sites_data.get('all_sites', []):
                if site.get('isPersonalSite', False) == is_personal_filter:
                    site_name = site.get('title') or site.get('name', 'Unknown')
                    creation_date = site.get('createdDateTime', '')
                    is_personal = site.get('isPersonalSite', False)
                    web_url = site.get('webUrl', '')
                    site_id = extract_site_id(site.get('id', ''))
                    template = site.get('template', 'Unknown')
                    
                    writer.writerow({
                        'Site Name': site_name,
                        'Creation Date': creation_date,
                        'Is Personal Site': is_personal,
                        'Web URL': web_url,
                        'Site ID': site_id,
                        'Template': template
                    })
                    filtered_count += 1
        
        site_type = "Personal" if is_personal_filter else "SharePoint"
        print(f"{site_type} sites CSV saved to {filename} ({filtered_count} sites)")
        
    except Exception as e:
        print(f"Error saving filtered CSV: {str(e)}")

def print_csv_preview(all_sites_data, preview_count=5):
    """Print a preview of the CSV data"""
    print(f"\n=== CSV Report Preview (First {preview_count} Sites) ===")
    print("Site Name | Creation Date | Is Personal Site | Web URL | Site ID | Template")
    print("-" * 120)
    
    count = 0
    for site in all_sites_data.get('all_sites', []):
        if count >= preview_count:
            break
            
        site_name = site.get('title') or site.get('name', 'Unknown')
        creation_date = site.get('createdDateTime', '')[:10] if site.get('createdDateTime') else 'N/A'  # Just date part
        is_personal = 'Yes' if site.get('isPersonalSite', False) else 'No'
        web_url = site.get('webUrl', '')
        site_id = extract_site_id(site.get('id', ''))
        template = site.get('template', 'Unknown')
        
        # Truncate long URLs for display
        display_url = web_url[:40] + "..." if len(web_url) > 40 else web_url
        
        print(f"{site_name[:20]:<20} | {creation_date:<12} | {is_personal:<16} | {display_url:<30} | {site_id:<36} | {template}")
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
    
    # Template breakdown
    template_breakdown = {}
    for site in all_sites:
        template_name = site.get('template', 'Unknown')
        if template_name not in template_breakdown:
            template_breakdown[template_name] = {'total': 0, 'sharepoint': 0, 'personal': 0}
        template_breakdown[template_name]['total'] += 1
        if site.get('isPersonalSite', False):
            template_breakdown[template_name]['personal'] += 1
        else:
            template_breakdown[template_name]['sharepoint'] += 1
    
    print(f"\nTemplate Distribution:")
    for template, counts in sorted(template_breakdown.items()):
        print(f"  {template}: {counts['total']} total ({counts['sharepoint']} SharePoint, {counts['personal']} Personal)")
    
    # Recent sites (created in last 30 days)
    try:
        thirty_days_ago = datetime.now() - timedelta(days=30)
        recent_sites = []
        
        for site in all_sites:
            created_date = site.get('createdDateTime')
            if created_date:
                try:
                    # Parse ISO date format
                    created_dt = datetime.fromisoformat(created_date.replace('Z', '+00:00'))
                    if created_dt.replace(tzinfo=None) > thirty_days_ago:
                        recent_sites.append(site)
                except:
                    pass
        
        print(f"\nRecent Sites (Last 30 days): {len(recent_sites)}")
        
    except Exception as e:
        print(f"\nCould not calculate recent sites: {str(e)}")

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
        
        # Save tokens to file
        tokens = {
            "graph_token": graph_token,
            "sharepoint_token": sharepoint_token
        }
        
        with open(f"{output_prefix}_tokens.json", "w") as f:
            json.dump(tokens, f, indent=2)
        
        print(f"Tokens saved to {output_prefix}_tokens.json")
        
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
        
        # Method 1: Get all sites at once
        try:
            print("Method 1: Getting all sites at once...")
            sites_result = get_all_sites_from_sharepoint(sharepoint_token, sharepoint_url)
            all_sites_data["all_sites"] = sites_result["all_sites"]
            all_sites_data["sharepoint_sites"] = sites_result["sharepoint_sites"]
            all_sites_data["personal_sites"] = sites_result["personal_sites"]
        except Exception as e:
            print(f"Error with Method 1: {str(e)}")
            
            # Method 2: Try with pagination if the first method fails
            try:
                print("Method 2: Getting sites with pagination...")
                sites_result = get_sites_with_pagination(sharepoint_token, sharepoint_url, page_size)
                all_sites_data["all_sites"] = sites_result["all_sites"]
                all_sites_data["sharepoint_sites"] = sites_result["sharepoint_sites"]
                all_sites_data["personal_sites"] = sites_result["personal_sites"]
            except Exception as e2:
                print(f"Error with Method 2: {str(e2)}")
        
        # Save all data to files
        save_sites_to_file(all_sites_data, f"{output_prefix}_backup.json")  # JSON backup
        save_sites_to_csv(all_sites_data, f"{output_prefix}_report.csv")   # Main CSV report
        
        # Also create separate CSV files for SharePoint and Personal sites
        save_filtered_csv(all_sites_data, False, f"{output_prefix}_sharepoint_only.csv")
        save_filtered_csv(all_sites_data, True, f"{output_prefix}_personal_only.csv")
        
        # Print CSV preview
        print_csv_preview(all_sites_data, preview_count)
        
        # Generate detailed summary
        generate_summary_report(all_sites_data)
        
        return all_sites_data
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

if __name__ == "__main__":
    main()
