import json
import uuid
import base64
import time
import requests
import csv
from datetime import datetime
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend

# Configuration
tenant_name = "geekbyteonline.onmicrosoft.com"
app_id = "73efa35d-6188-42d4-b258-838a977eb149"
scope_graph = "https://graph.microsoft.com/.default"
scope_sharepoint = "https://geekbyteonline.sharepoint.com/.default"

# Certificate paths (update these with your actual file paths)
CERTIFICATE_PATH = "certificate.pem"
PRIVATE_KEY_PATH = "private_key.pem"

# Base URL for getting subsites
BASE_SUBSITES_URL = "https://geekbyteonline.sharepoint.com/sites/New365/XYX/_api/web/webs"

def load_certificate_and_key():
    """Load certificate and private key from PEM files"""
    try:
        # Load certificate
        with open(CERTIFICATE_PATH, "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())

        # Load private key
        with open(PRIVATE_KEY_PATH, "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())

        return certificate, private_key

    except Exception as e:
        print(f"Error loading certificate or private key: {str(e)}")
        raise

def get_jwt_token(certificate, private_key, scope):
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

def get_access_token(jwt, scope):
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

def get_subsites(base_url, sharepoint_token, level=0, parent_title="Root"):
    """Get all subsites from the specified SharePoint location (recursive)"""
    headers = {
        "Authorization": f"Bearer {sharepoint_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json"
    }
    
    # Construct the API URL for getting subsites
    api_url = f"{base_url}/_api/web/webs"
    indent = "  " * level
    
    try:
        print(f"{indent}Fetching subsites from: {api_url}")
        response = requests.get(api_url, headers=headers)
        response.raise_for_status()
        
        data = response.json()
        subsites = []
        
        if 'd' in data and 'results' in data['d']:
            for subsite in data['d']['results']:
                title = subsite.get('Title', 'Unknown')
                url = subsite.get('Url', '')
                created = subsite.get('Created', 'Unknown')
                
                subsite_info = {
                    'Title': title,
                    'Url': url,
                    'Created': created,
                    'Level': level,
                    'ParentTitle': parent_title,
                    'FullPath': f"{parent_title} > {title}" if parent_title != "Root" else title
                }
                
                subsites.append(subsite_info)
                print(f"{indent}Found subsite (Level {level}): {title} - {url} (Created: {created})")
        
        print(f"{indent}Total subsites found at level {level}: {len(subsites)}")
        return subsites
        
    except requests.exceptions.HTTPError as err:
        print(f"{indent}HTTP Error getting subsites from {api_url}: {err}")
        if response.status_code == 404:
            print(f"{indent}No subsites found or endpoint not accessible")
            return []
        else:
            print(f"{indent}Response: {response.text}")
            raise
    except Exception as err:
        print(f"{indent}Error getting subsites from {api_url}: {err}")
        return []

def get_all_subsites_recursive(sharepoint_token, max_depth=5):
    """Get all subsites recursively from the base location"""
    all_subsites = []
    processed_urls = set()  # Prevent infinite loops
    
    print("=== Starting Recursive Subsite Discovery ===")
    
    # Start with the base URL
    base_site_info = {
        'Title': 'XYX (Root Site)',
        'Url': 'https://geekbyteonline.sharepoint.com/sites/New365/XYX',
        'Created': 'Unknown',
        'Level': -1,
        'ParentTitle': 'Root',
        'FullPath': 'XYX (Root Site)'
    }
    
    # Queue for processing sites (site_info, current_level)
    sites_to_process = [(base_site_info, 0)]
    
    while sites_to_process and len(sites_to_process[0]) > 0:
        current_site, current_level = sites_to_process.pop(0)
        
        if current_level > max_depth:
            print(f"Reached maximum depth ({max_depth}), skipping deeper levels")
            continue
            
        if current_site['Url'] in processed_urls:
            print(f"Already processed {current_site['Url']}, skipping to avoid loops")
            continue
            
        processed_urls.add(current_site['Url'])
        
        # Add current site to results if it's not the root
        if current_level >= 0:
            all_subsites.append(current_site)
        
        # Get subsites of current site
        try:
            subsites = get_subsites(
                current_site['Url'], 
                sharepoint_token, 
                current_level, 
                current_site['Title']
            )
            
            # Add found subsites to processing queue
            for subsite in subsites:
                subsite['Level'] = current_level
                sites_to_process.append((subsite, current_level + 1))
                
            # Small delay to avoid throttling
            time.sleep(0.3)
            
        except Exception as e:
            print(f"Error processing {current_site['Url']}: {str(e)}")
            continue
    
    print(f"\n=== Recursive Discovery Complete ===")
    print(f"Total sites found across all levels: {len(all_subsites)}")
    
    # Print hierarchy summary
    level_counts = {}
    for site in all_subsites:
        level = site['Level']
        level_counts[level] = level_counts.get(level, 0) + 1
    
    print("Sites by level:")
    for level in sorted(level_counts.keys()):
        print(f"  Level {level}: {level_counts[level]} sites")
    
    return all_subsites

def get_site_policy_and_properties(site_url, sharepoint_token):
    """Get policy name and additional properties for a specific site"""
    # Construct the API endpoint for getting site properties
    api_url = f"{site_url}/_api/web/allproperties"
    
    headers = {
        "Authorization": f"Bearer {sharepoint_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json"
    }
    
    try:
        print(f"Getting properties for: {site_url}")
        response = requests.get(api_url, headers=headers)
        response.raise_for_status()
        
        data = response.json()
        
        # Initialize default values
        site_properties = {
            'PolicyName': "No Policy Found",
            'CloseDate': "Not Set",
            'DeleteDate': "Not Set",
            'SiteClosed': "Unknown"
        }
        
        # Look for properties in different possible locations
        if 'd' in data:
            properties = data['d']
            
            # Check common property names for policy
            policy_fields = ['PolicyName', 'policyname', 'SitePolicy', 'sitepolicy', 'Policy', 'policy']
            
            for field in policy_fields:
                if field in properties and properties[field]:
                    site_properties['PolicyName'] = properties[field]
                    break
            
            # Check for close date fields
            close_date_fields = ['CloseDate', 'closedate', 'SiteCloseDate', 'siteclosedate']
            for field in close_date_fields:
                if field in properties and properties[field]:
                    site_properties['CloseDate'] = properties[field]
                    break
            
            # Check for delete date fields
            delete_date_fields = ['DeleteDate', 'deletedate', 'SiteDeleteDate', 'sitedeletedate']
            for field in delete_date_fields:
                if field in properties and properties[field]:
                    site_properties['DeleteDate'] = properties[field]
                    break
            
            # Check for site closed status
            closed_fields = ['SiteClosed', 'siteclosed', 'IsClosed', 'isclosed', 'Closed', 'closed']
            for field in closed_fields:
                if field in properties and properties[field] is not None:
                    site_properties['SiteClosed'] = str(properties[field])
                    break
            
            # If not found in direct properties, check in custom properties or AllProperties
            if 'AllProperties' in properties:
                all_props = properties['AllProperties']
                
                # Check policy in AllProperties
                if site_properties['PolicyName'] == "No Policy Found":
                    for field in policy_fields:
                        if field in all_props and all_props[field]:
                            site_properties['PolicyName'] = all_props[field]
                            break
                
                # Check close date in AllProperties
                if site_properties['CloseDate'] == "Not Set":
                    for field in close_date_fields:
                        if field in all_props and all_props[field]:
                            site_properties['CloseDate'] = all_props[field]
                            break
                
                # Check delete date in AllProperties
                if site_properties['DeleteDate'] == "Not Set":
                    for field in delete_date_fields:
                        if field in all_props and all_props[field]:
                            site_properties['DeleteDate'] = all_props[field]
                            break
                
                # Check site closed in AllProperties
                if site_properties['SiteClosed'] == "Unknown":
                    for field in closed_fields:
                        if field in all_props and all_props[field] is not None:
                            site_properties['SiteClosed'] = str(all_props[field])
                            break
        
        print(f"Properties found - Policy: {site_properties['PolicyName']}, " +
              f"CloseDate: {site_properties['CloseDate']}, " +
              f"DeleteDate: {site_properties['DeleteDate']}, " +
              f"SiteClosed: {site_properties['SiteClosed']}")
        
        return site_properties
        
    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error getting properties for {site_url}: {err}")
        error_msg = f"Error: {response.status_code}"
        if response.status_code == 404:
            error_msg = "Site Not Found"
        elif response.status_code == 403:
            error_msg = "Access Denied"
        
        return {
            'PolicyName': error_msg,
            'CloseDate': error_msg,
            'DeleteDate': error_msg,
            'SiteClosed': error_msg
        }
    except Exception as err:
        print(f"Error getting properties for {site_url}: {err}")
        error_msg = "Error Retrieving Properties"
        return {
            'PolicyName': error_msg,
            'CloseDate': error_msg,
            'DeleteDate': error_msg,
            'SiteClosed': error_msg
        }

def generate_report(subsites_data, output_format='both'):
    """Generate report in JSON and/or CSV format"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Generate JSON report
    if output_format in ['json', 'both']:
        json_filename = f"sharepoint_policy_report_{timestamp}.json"
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump({
                "report_generated": datetime.now().isoformat(),
                "total_sites": len(subsites_data),
                "sites": subsites_data
            }, f, indent=2, ensure_ascii=False)
        print(f"JSON report saved: {json_filename}")
    
    # Generate CSV report
    if output_format in ['csv', 'both']:
        csv_filename = f"sharepoint_policy_report_{timestamp}.csv"
        with open(csv_filename, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(['Site Title', 'Site URL', 'Created Date', 'Policy Name', 'Close Date', 'Delete Date', 'Site Closed'])
            for site in subsites_data:
                writer.writerow([
                    site['Title'], 
                    site['Url'], 
                    site['Created'],
                    site['PolicyName'], 
                    site['CloseDate'], 
                    site['DeleteDate'], 
                    site['SiteClosed']
                ])
        print(f"CSV report saved: {csv_filename}")
    
    return subsites_data

def main():
    try:
        print("=== SharePoint Subsites Policy Report Generator ===")
        print("Step 1: Loading certificates and authenticating...")
        
        # Load certificate and private key
        certificate, private_key = load_certificate_and_key()
        print("Certificate and private key loaded successfully")
        
        # Get SharePoint token
        sharepoint_jwt = get_jwt_token(certificate, private_key, scope_sharepoint)
        sharepoint_token = get_access_token(sharepoint_jwt, scope_sharepoint)
        print("SharePoint access token retrieved successfully")
        
        print("\nStep 2: Fetching all subsites...")
        # Get all subsites
        subsites = get_subsites(sharepoint_token)
        
        if not subsites:
            print("No subsites found!")
            return
        
        print(f"\nStep 3: Getting policy and property information for {len(subsites)} subsites...")
        # Get policy and properties for each subsite
        subsites_with_properties = []
        
        for i, subsite in enumerate(subsites, 1):
            print(f"\nProcessing {i}/{len(subsites)}: {subsite['Title']}")
            site_properties = get_site_policy_and_properties(subsite['Url'], sharepoint_token)
            
            subsites_with_properties.append({
                'Title': subsite['Title'],
                'Url': subsite['Url'],
                'Created': subsite['Created'],
                'PolicyName': site_properties['PolicyName'],
                'CloseDate': site_properties['CloseDate'],
                'DeleteDate': site_properties['DeleteDate'],
                'SiteClosed': site_properties['SiteClosed']
            })
            
            # Add a small delay to avoid throttling
            time.sleep(0.5)
        
        print("\nStep 4: Generating reports...")
        # Generate reports
        generate_report(subsites_with_properties, 'both')
        
        print("\n=== Summary ===")
        print(f"Total subsites processed: {len(subsites_with_properties)}")
        
        # Summary statistics
        policy_counts = {}
        closed_counts = {'True': 0, 'False': 0, 'Unknown': 0}
        
        for site in subsites_with_properties:
            policy = site['PolicyName']
            policy_counts[policy] = policy_counts.get(policy, 0) + 1
            
            closed_status = site['SiteClosed'].lower()
            if closed_status in ['true', '1', 'yes']:
                closed_counts['True'] += 1
            elif closed_status in ['false', '0', 'no']:
                closed_counts['False'] += 1
            else:
                closed_counts['Unknown'] += 1
        
        print("\nPolicy distribution:")
        for policy, count in policy_counts.items():
            print(f"  {policy}: {count} sites")
        
        print(f"\nSite closure status:")
        print(f"  Closed sites: {closed_counts['True']}")
        print(f"  Open sites: {closed_counts['False']}")
        print(f"  Unknown status: {closed_counts['Unknown']}")
        
        print("\nReport generation completed successfully!")
        return subsites_with_properties
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

if __name__ == "__main__":
    main()
