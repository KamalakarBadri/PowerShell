#!/usr/bin/env python3
"""
SharePoint Site Admin Report Generator
Generates a CSV report of site administrators for a list of SharePoint sites
Detects groups from LoginName pattern (xxxxx|xxxxxxx|<groupid>_o) and expands group owners via Graph API
"""

import requests
import json
import uuid
import base64
import time
import csv
import os
import logging
import xml.etree.ElementTree as ET
from datetime import datetime
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend
from typing import List, Dict, Any, Optional, Set

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Configuration - Update these values
CONFIG = {
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "tenant_name": "geekbyteonline.onmicrosoft.com",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "client_secret": "CyG8Q~FYHuCMSyVmt4sNxt5IejrMc2c24Ziz4a.t",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    "scopes": {
        "graph": "https://graph.microsoft.com/.default",
        "sharepoint": "https://geekbyteonline.sharepoint.com/.default"
    }
}

# Cache for group owners to avoid repeated API calls
group_owner_cache = {}

def get_token_with_certificate(scope: str) -> Optional[str]:
    """Get access token using certificate-based authentication"""
    try:
        if not os.path.exists(CONFIG['certificate_path']) or not os.path.exists(CONFIG['private_key_path']):
            logger.warning("Certificate files not found, falling back to client secret")
            return None
            
        with open(CONFIG['certificate_path'], "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
        with open(CONFIG['private_key_path'], "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())

        now = int(time.time())
        jwt_header = {
            "alg": "RS256",
            "typ": "JWT",
            "x5t": base64.urlsafe_b64encode(certificate.fingerprint(hashes.SHA1())).decode().rstrip('=')
        }
        jwt_payload = {
            "aud": f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            "exp": now + 300,
            "iss": CONFIG['app_id'],
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": CONFIG['app_id']
        }

        encoded_header = base64.urlsafe_b64encode(json.dumps(jwt_header).encode()).decode().rstrip('=')
        encoded_payload = base64.urlsafe_b64encode(json.dumps(jwt_payload).encode()).decode().rstrip('=')
        jwt_unsigned = f"{encoded_header}.{encoded_payload}"
        signature = private_key.sign(jwt_unsigned.encode(), padding.PKCS1v15(), hashes.SHA256())
        encoded_signature = base64.urlsafe_b64encode(signature).decode().rstrip('=')
        jwt = f"{jwt_unsigned}.{encoded_signature}"

        token_response = requests.post(
            f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            data={
                "client_id": CONFIG['app_id'],
                "client_assertion": jwt,
                "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                "scope": scope,
                "grant_type": "client_credentials"
            }
        )

        if token_response.status_code == 200:
            logger.info("Successfully obtained token using certificate")
            return token_response.json()["access_token"]
        else:
            logger.error(f"Certificate token request failed: {token_response.text}")
            return None
            
    except Exception as e:
        logger.exception("Certificate authentication failed")
        return None

def get_token_with_secret(scope: str) -> Optional[str]:
    """Get access token using client secret authentication"""
    try:
        token_url = f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token"
        
        token_data = {
            "client_id": CONFIG['app_id'],
            "client_secret": CONFIG['client_secret'],
            "scope": scope,
            "grant_type": "client_credentials"
        }
        
        token_response = requests.post(token_url, data=token_data)

        if token_response.status_code == 200:
            logger.info("Successfully obtained token using client secret")
            return token_response.json()["access_token"]
        else:
            logger.error(f"Client secret token request failed: {token_response.text}")
            return None
            
    except Exception as e:
        logger.exception("Client secret authentication failed")
        return None

def get_graph_token() -> Optional[str]:
    """Get Graph API access token"""
    token = get_token_with_certificate(CONFIG['scopes']['graph'])
    if not token:
        token = get_token_with_secret(CONFIG['scopes']['graph'])
    return token

def get_sharepoint_token() -> Optional[str]:
    """Get SharePoint access token"""
    token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
    if not token:
        token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
    return token

def extract_group_id_from_loginname(login_name: str) -> Optional[str]:
    """
    Extract Group ID from LoginName.
    
    Format: "xxxxx|xxxxxxx|<groupid>_o"
    
    Examples:
    - "c:0o.c|tenant|a1b2c3d4-e5f6-7890-abcd-ef1234567890_o" → "a1b2c3d4-e5f6-7890-abcd-ef1234567890"
    - "i:0#.f|membership|group123_o" → "group123"
    - "c:0o.c|system|xyz_o" → "xyz"
    """
    if not login_name:
        return None
    
    try:
        # Split by '|' and get the last part
        parts = login_name.split('|')
        if len(parts) >= 1:
            last_part = parts[-1]  # Gets "<groupid>_o" or just "<groupid>"
            
            # Remove the "_o" suffix if present
            if last_part.endswith('_o'):
                group_id = last_part[:-2]  # Remove "_o"
            else:
                group_id = last_part
            
            # Clean up any extra characters
            if group_id:
                group_id = group_id.strip()
                return group_id
    except Exception as e:
        logger.error(f"Failed to extract group ID from login_name: {login_name}")
        return None
    
    return None

def get_group_owners(group_id: str) -> List[Dict[str, Any]]:
    """Get owners of a Microsoft 365 group using Graph API"""
    try:
        # Check cache first
        if group_id in group_owner_cache:
            logger.info(f"Using cached group owners for {group_id}")
            return group_owner_cache[group_id]
        
        token = get_graph_token()
        if not token:
            logger.error("Failed to get Graph API token")
            return []
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        # Get group owners
        url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/owners"
        logger.info(f"Fetching owners for group: {group_id}")
        
        response = requests.get(url, headers=headers)
        
        if response.status_code != 200:
            logger.error(f"Failed to get group owners: {response.text}")
            return []
        
        data = response.json()
        owners = []
        
        for owner in data.get('value', []):
            # Get user details including email
            user_details = {
                'user_id': owner.get('id'),
                'title': owner.get('displayName', ''),
                'email': owner.get('userPrincipalName', ''),
                'login_name': owner.get('userPrincipalName', ''),
                'is_site_admin': True,
                'is_group_member': True,
                'group_id': group_id
            }
            owners.append(user_details)
        
        # Cache the results
        group_owner_cache[group_id] = owners
        logger.info(f"Found {len(owners)} owners for group {group_id}")
        
        return owners
        
    except Exception as e:
        logger.exception(f"Failed to get owners for group {group_id}")
        return []

def parse_site_users_xml(xml_content: str) -> List[Dict[str, Any]]:
    """Parse SharePoint site users XML and return all admins with group expansion"""
    try:
        root = ET.fromstring(xml_content)
        
        namespaces = {
            'd': 'http://schemas.microsoft.com/ado/2007/08/dataservices',
            'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata',
            'atom': 'http://www.w3.org/2005/Atom'
        }
        
        entries = root.findall('.//atom:entry', namespaces)
        logger.debug(f"Found {len(entries)} entries in XML")
        
        admins = []
        group_ids_to_expand = []
        
        for entry in entries:
            content = entry.find('atom:content', namespaces)
            if content is not None:
                properties = content.find('m:properties', namespaces)
                if properties is not None:
                    # Extract user details
                    user_id_elem = properties.find('d:Id', namespaces)
                    title_elem = properties.find('d:Title', namespaces)  
                    email_elem = properties.find('d:Email', namespaces)
                    login_name_elem = properties.find('d:LoginName', namespaces)
                    is_site_admin_elem = properties.find('d:IsSiteAdmin', namespaces)
                    principal_type_elem = properties.find('d:PrincipalType', namespaces)
                    
                    user_id = user_id_elem.text if user_id_elem is not None else None
                    title = title_elem.text if title_elem is not None else None
                    email = email_elem.text if email_elem is not None else None
                    login_name = login_name_elem.text if login_name_elem is not None else None
                    is_site_admin = is_site_admin_elem.text == 'true' if is_site_admin_elem is not None else False
                    principal_type = int(principal_type_elem.text) if principal_type_elem is not None else None
                    
                    if is_site_admin:
                        # Check if this is a group using multiple methods
                        is_group = False
                        group_id = None
                        detection_method = ""
                        
                        # Method 1: Check PrincipalType = 4 (Microsoft 365 Group)
                        if principal_type == 4:
                            is_group = True
                            detection_method = "PrincipalType=4"
                            # Try to extract group ID from login_name
                            if login_name:
                                group_id = extract_group_id_from_loginname(login_name)
                        
                        # Method 2: Check if title contains "owners" or "owner"
                        if not is_group and title:
                            title_lower = title.lower()
                            if 'owners' in title_lower or 'owner' in title_lower:
                                is_group = True
                                detection_method = "Title contains 'owners'"
                                if login_name:
                                    group_id = extract_group_id_from_loginname(login_name)
                        
                        # Method 3: Check if login_name ends with "_o"
                        if not is_group and login_name and login_name.endswith('_o'):
                            is_group = True
                            detection_method = "LoginName ends with '_o'"
                            group_id = extract_group_id_from_loginname(login_name)
                        
                        # Method 4: Check if login_name has group ID pattern
                        if not is_group and login_name:
                            extracted_id = extract_group_id_from_loginname(login_name)
                            if extracted_id:
                                is_group = True
                                detection_method = "LoginName pattern matched"
                                group_id = extracted_id
                        
                        admin_entry = {
                            'user_id': user_id,
                            'title': title,
                            'email': email,
                            'login_name': login_name,
                            'is_site_admin': is_site_admin,
                            'principal_type': principal_type,
                            'is_group': is_group,
                            'group_id': group_id,
                            'detection_method': detection_method
                        }
                        
                        if is_group and group_id:
                            logger.info(f"Found group: {title} (ID: {group_id}) - Detected by: {detection_method}")
                            group_ids_to_expand.append({
                                'group_id': group_id,
                                'group_name': title,
                                'admin_entry': admin_entry
                            })
                        else:
                            admins.append(admin_entry)
        
        # Expand groups and add their owners
        for group_info in group_ids_to_expand:
            group_id = group_info['group_id']
            group_name = group_info['group_name']
            logger.info(f"Expanding group: {group_name} ({group_id})")
            
            owners = get_group_owners(group_id)
            
            if owners:
                logger.info(f"Added {len(owners)} owners from group {group_name}")
                for owner in owners:
                    # Check if owner is already in the list (avoid duplicates)
                    if not any(a.get('email') == owner.get('email') for a in admins):
                        admins.append(owner)
            else:
                logger.warning(f"No owners found for group {group_name}")
                # If no owners found, keep the group itself in the list
                admins.append(group_info['admin_entry'])
        
        logger.info(f"Final admin list has {len(admins)} users (including expanded groups)")
        return admins
        
    except Exception as e:
        logger.exception("Failed to parse site users XML")
        return []

def get_site_admins(site_url: str) -> List[Dict[str, Any]]:
    """Get all site administrators for a SharePoint site, expanding groups"""
    try:
        if not site_url.endswith('/'):
            site_url += '/'
        
        site_users_url = f"{site_url}_api/web/siteusers"
        logger.info(f"Fetching site admins from URL: {site_users_url}")
        
        token = get_sharepoint_token()
        if not token:
            raise Exception("Failed to obtain SharePoint access token")
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/xml"
        }
        
        response = requests.get(site_users_url, headers=headers)
        
        if response.status_code != 200:
            logger.error(f"SharePoint API failed: {response.status_code} - {response.text}")
            return []
        
        admins = parse_site_users_xml(response.text)
        return admins
        
    except Exception as e:
        logger.exception(f"Failed to get site admins from {site_url}")
        return []

def normalize_site_url(site_input: str) -> str:
    """Normalize site URL from various input formats"""
    site_input = site_input.strip()
    
    if site_input.startswith('http'):
        return site_input
    
    if not site_input.startswith('sites/'):
        site_input = f"sites/{site_input}"
    
    tenant_name = CONFIG['tenant_name'].split('.')[0]
    return f"https://{tenant_name}.sharepoint.com/{site_input}"

def generate_report(site_list: List[Dict[str, str]], output_file: str = "sharepoint_admin_report.csv"):
    """Generate CSV report for a list of SharePoint sites with group expansion"""
    
    report_data = []
    
    for site_entry in site_list:
        site_input = site_entry.get('site_url') or site_entry.get('Site URL') or site_entry.get('Site') or site_entry.get('site')
        site_name = site_entry.get('site_name') or site_entry.get('Site Name') or site_entry.get('Name') or site_input
        
        if not site_input:
            logger.warning(f"Skipping entry with no site URL: {site_entry}")
            continue
        
        try:
            site_url = normalize_site_url(site_input)
            logger.info(f"\n{'='*60}")
            logger.info(f"Processing site: {site_url}")
            logger.info(f"{'='*60}")
            
            admins = get_site_admins(site_url)
            
            if not admins:
                report_data.append({
                    'Site URL': site_url,
                    'Site Name': site_name,
                    'Admin Emails': '',
                    'Admin Names': '',
                    'Admin Count': 0,
                    'Groups Found': '',
                    'Group Members Expanded': 0,
                    'Error': 'No admins found or site inaccessible',
                    'Additional Info': site_entry.get('additional_info', '')
                })
            else:
                # Separate users and groups for reporting
                users = [a for a in admins if a.get('principal_type') != 4 and not a.get('is_group_member')]
                group_members = [a for a in admins if a.get('is_group_member', False)]
                groups = [a for a in admins if a.get('is_group') and not a.get('is_group_member')]
                
                # Extract admin emails and names
                admin_emails = [a.get('email', a.get('login_name', '')) for a in admins 
                              if a.get('email') or a.get('login_name')]
                admin_names = [a.get('title', a.get('login_name', '')) for a in admins 
                             if a.get('title') or a.get('login_name')]
                
                # Extract group names
                group_names = [g.get('title', g.get('login_name', '')) for g in groups if g.get('title') or g.get('login_name')]
                
                report_data.append({
                    'Site URL': site_url,
                    'Site Name': site_name,
                    'Admin Emails': ', '.join(admin_emails) if admin_emails else '',
                    'Admin Names': ', '.join(admin_names) if admin_names else '',
                    'Admin Count': len(admins),
                    'Groups Found': ', '.join(group_names) if group_names else 'None',
                    'Group Members Expanded': len(group_members),
                    'Error': '',
                    'Additional Info': site_entry.get('additional_info', '')
                })
                
                logger.info(f"  ✅ Found {len(admins)} admins:")
                logger.info(f"     - Direct users: {len(users)}")
                logger.info(f"     - Groups: {len(groups)}")
                logger.info(f"     - Group members expanded: {len(group_members)}")
                
        except Exception as e:
            logger.error(f"Error processing {site_input}: {str(e)}")
            report_data.append({
                'Site URL': site_input,
                'Site Name': site_name,
                'Admin Emails': '',
                'Admin Names': '',
                'Admin Count': 0,
                'Groups Found': '',
                'Group Members Expanded': 0,
                'Error': str(e),
                'Additional Info': site_entry.get('additional_info', '')
            })
    
    # Write to CSV
    try:
        output_headers = [
            'Site URL', 
            'Site Name', 
            'Admin Emails', 
            'Admin Names', 
            'Admin Count',
            'Groups Found',
            'Group Members Expanded',
            'Error', 
            'Additional Info'
        ]
        
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=output_headers)
            writer.writeheader()
            for row in report_data:
                writer.writerow(row)
        
        logger.info(f"Report generated successfully: {output_file}")
        print(f"\n✅ Report generated: {output_file}")
        print(f"📊 Processed {len(site_list)} sites")
        
        successful = sum(1 for row in report_data if not row['Error'])
        print(f"✅ Successful: {successful}")
        print(f"❌ Failed: {len(site_list) - successful}")
        
        # Print group expansion summary
        total_groups = sum(1 for row in report_data if row.get('Groups Found') and row['Groups Found'] != 'None')
        if total_groups > 0:
            print(f"👥 Groups found: {total_groups}")
        
    except Exception as e:
        logger.error(f"Failed to write CSV report: {str(e)}")
        raise

def read_sites_from_csv(csv_file: str) -> List[Dict[str, str]]:
    """Read sites from input CSV file"""
    sites = []
    
    try:
        with open(csv_file, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            
            if not reader.fieldnames:
                print("❌ CSV file has no headers")
                return []
            
            print(f"\n📋 Found headers: {', '.join(reader.fieldnames)}")
            print(f"   Expected headers: 'Site URL', 'Site Name' (or variations)")
            print()
            
            for row in reader:
                # Try to find site URL column
                site_url = None
                for col in ['Site URL', 'Site', 'site_url', 'site', 'URL', 'url']:
                    if col in row and row[col]:
                        site_url = row[col]
                        break
                
                if site_url:
                    site_entry = {'site_url': site_url}
                    
                    # Try to find site name column
                    for col in ['Site Name', 'Name', 'site_name', 'name']:
                        if col in row and row[col]:
                            site_entry['site_name'] = row[col]
                            break
                    
                    # Store any additional columns as additional_info
                    additional_info = []
                    for key, value in row.items():
                        if key not in ['Site URL', 'Site', 'site_url', 'site', 'URL', 'url', 
                                       'Site Name', 'Name', 'site_name', 'name'] and value:
                            additional_info.append(f"{key}: {value}")
                    
                    if additional_info:
                        site_entry['additional_info'] = ' | '.join(additional_info)
                    else:
                        site_entry['additional_info'] = ''
                    
                    sites.append(site_entry)
                else:
                    logger.warning(f"Skipping row with no site URL: {row}")
        
        return sites
        
    except Exception as e:
        logger.error(f"Failed to read CSV file: {str(e)}")
        return []

def main():
    """Main function to run the report generator"""
    
    print("=" * 70)
    print("SharePoint Site Admin Report Generator (with Group Expansion)")
    print("=" * 70)
    print()
    print("🔍 This tool will:")
    print("   - Find all site administrators")
    print("   - Detect groups using:")
    print("     • PrincipalType = 4")
    print("     • Title containing 'owners'")
    print("     • LoginName pattern: xxxxx|xxxxxxx|<groupid>_o")
    print("   - Expand groups to show all group owners")
    print("   - Generate a comprehensive CSV report")
    print()
    
    print("Choose input method:")
    print("1. Interactive input (enter sites manually)")
    print("2. CSV file input")
    print()
    
    choice = input("Enter your choice (1 or 2): ").strip()
    
    site_list = []
    
    if choice == '2':
        csv_file = input("Enter input CSV file path: ").strip()
        
        if not os.path.exists(csv_file):
            print(f"❌ File not found: {csv_file}")
            return
        
        print(f"\n📄 Reading sites from: {csv_file}")
        site_list = read_sites_from_csv(csv_file)
        
        if not site_list:
            print("❌ No sites found in CSV file")
            return
        
        print(f"\n✅ Found {len(site_list)} sites in CSV file")
        
    else:
        # Interactive input
        print("\nEnter SharePoint sites (one per line, press Enter twice to finish):")
        print("Examples:")
        print("  - Full URL: https://tenant.sharepoint.com/sites/projectx")
        print("  - Short name: projectx")
        print("  - Path: sites/projectx")
        print()
        
        while True:
            line = input().strip()
            if not line:
                if site_list:
                    break
                else:
                    print("Please enter at least one site")
                    continue
            site_list.append({'site_url': line, 'site_name': line, 'additional_info': ''})
            print(f"  Added: {line}")
    
    if not site_list:
        print("No sites provided. Exiting.")
        return
    
    print(f"\n📋 Processing {len(site_list)} sites...")
    print()
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"sharepoint_admin_report_{timestamp}.csv"
    
    generate_report(site_list, output_file)

if __name__ == "__main__":
    main()
