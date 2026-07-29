#!/usr/bin/env python3
"""
SharePoint Site Admin Report Generator (Multi-Threaded)
Generates CSV report with separate columns for Direct Users and Group Owners
Extracts Group ID from LoginName field (c:0o.c|tenant|groupid_o)
Supports ignoring specific users and groups (exact matches only)
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
from typing import List, Dict, Any, Optional, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(threadName)s - %(levelname)s - %(message)s')
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
    },
    "max_workers": 5,  # Number of threads for parallel processing
    "ignore_users": [],  # List of users to ignore (email addresses or display names)
    "ignore_groups": []  # List of groups to ignore (group names)
}

# Cache for group owners to avoid repeated API calls
group_owner_cache = {}
cache_lock = Lock()
stats_lock = Lock()

# Statistics tracking
stats = {
    'total_sites': 0,
    'processed': 0,
    'failed': 0,
    'total_direct_users': 0,
    'total_group_owners': 0,
    'total_groups': 0,
    'ignored_users': 0,
    'ignored_groups': 0
}

def should_ignore_user(user_name: str, user_email: str = None) -> bool:
    """
    Check if a user should be ignored based on configured ignore lists
    """
    if not user_name and not user_email:
        return False
    
    # Check exact matches in ignore_users list
    if user_email and user_email in CONFIG['ignore_users']:
        logger.info(f"Ignoring user (exact email match): {user_email}")
        return True
    
    if user_name and user_name in CONFIG['ignore_users']:
        logger.info(f"Ignoring user (exact name match): {user_name}")
        return True
    
    return False

def should_ignore_group(group_name: str) -> bool:
    """
    Check if a group should be ignored based on configured ignore lists
    """
    if not group_name:
        return False
    
    # Check exact matches in ignore_groups list
    if group_name in CONFIG['ignore_groups']:
        logger.info(f"Ignoring group (exact match): {group_name}")
        return True
    
    return False

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
    Extract Group ID from LoginName format: "xxxxx|xxxxxxx|<groupid>_o"
    
    Examples:
    - "c:0o.c|geekbyteonline|a1b2c3d4-e5f6-7890_o" → "a1b2c3d4-e5f6-7890"
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

def get_group_owners(group_id: str) -> Tuple[List[Dict[str, Any]], str]:
    """
    Get owners of a Microsoft 365 group using Graph API
    
    Args:
        group_id: The Group ID extracted from LoginName (e.g., "a1b2c3d4-e5f6-7890")
    
    Returns:
        Tuple of (list of owners, group title)
    """
    try:
        # Check cache first
        with cache_lock:
            if group_id in group_owner_cache:
                logger.info(f"Using cached group owners for {group_id}")
                return group_owner_cache[group_id]
        
        token = get_graph_token()
        if not token:
            logger.error("Failed to get Graph API token")
            return [], ""
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        # First get group details to get the display name
        group_url = f"https://graph.microsoft.com/v1.0/groups/{group_id}"
        logger.info(f"Fetching group details: {group_id}")
        group_response = requests.get(group_url, headers=headers)
        group_title = group_id  # Default to ID if title not found
        
        if group_response.status_code == 200:
            group_data = group_response.json()
            group_title = group_data.get('displayName', group_id)
            logger.info(f"Group title: {group_title}")
        else:
            logger.warning(f"Could not get group title: {group_response.text}")
        
        # Get group owners
        url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/owners"
        logger.info(f"Fetching owners for group: {group_id}")
        
        response = requests.get(url, headers=headers)
        
        if response.status_code != 200:
            logger.error(f"Failed to get group owners: {response.text}")
            return [], group_title
        
        data = response.json()
        owners = []
        
        for owner in data.get('value', []):
            user_details = {
                'user_id': owner.get('id'),
                'title': owner.get('displayName', ''),
                'email': owner.get('userPrincipalName', ''),
                'login_name': owner.get('userPrincipalName', ''),
                'is_site_admin': True,
                'is_group_member': True,
                'group_id': group_id,
                'group_title': group_title
            }
            owners.append(user_details)
        
        # Cache the results
        with cache_lock:
            group_owner_cache[group_id] = (owners, group_title)
        
        logger.info(f"Found {len(owners)} owners for group {group_id} ({group_title})")
        return owners, group_title
        
    except Exception as e:
        logger.exception(f"Failed to get owners for group {group_id}")
        return [], group_id if 'group_id' in locals() else "Unknown"

def parse_site_users_xml(xml_content: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, str]]]:
    """
    Parse SharePoint site users XML and return direct users and group owners separately
    
    Important: Group ID is extracted from LoginName, NOT from Id field
    """
    try:
        root = ET.fromstring(xml_content)
        
        namespaces = {
            'd': 'http://schemas.microsoft.com/ado/2007/08/dataservices',
            'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata',
            'atom': 'http://www.w3.org/2005/Atom'
        }
        
        entries = root.findall('.//atom:entry', namespaces)
        logger.debug(f"Found {len(entries)} entries in XML")
        
        direct_users = []
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
                        # GROUP: PrincipalType = 4
                        if principal_type == 4:
                            logger.info(f"Found group: {title} (LoginName: {login_name})")
                            
                            # Check if group should be ignored
                            if title and should_ignore_group(title):
                                with stats_lock:
                                    stats['ignored_groups'] += 1
                                logger.info(f"Ignoring group: {title}")
                                continue
                            
                            # Extract Group ID from LoginName (NOT from Id field!)
                            if login_name:
                                group_id = extract_group_id_from_loginname(login_name)
                                if group_id:
                                    logger.info(f"Extracted Group ID: {group_id} from LoginName: {login_name}")
                                    group_ids_to_expand.append({
                                        'group_id': group_id,  # ← This comes from LoginName
                                        'group_name': title,
                                        'group_title': title
                                    })
                                else:
                                    logger.warning(f"Could not extract group ID from LoginName: {login_name}")
                            else:
                                logger.warning(f"No LoginName found for group: {title}")
                        
                        # USER: PrincipalType = 1 (or any other value that's not 4)
                        else:
                            logger.info(f"Found direct user: {title} ({email})")
                            
                            # Check if user should be ignored
                            if should_ignore_user(title, email):
                                with stats_lock:
                                    stats['ignored_users'] += 1
                                logger.info(f"Ignoring user: {title} ({email})")
                                continue
                            
                            direct_users.append({
                                'user_id': user_id,
                                'title': title,
                                'email': email,
                                'login_name': login_name,
                                'is_site_admin': is_site_admin,
                                'principal_type': principal_type,
                                'is_group_member': False,
                                'group_title': ''
                            })
        
        # Expand groups and get all owners
        group_owners = []
        groups_info = []
        
        for group_info in group_ids_to_expand:
            group_id = group_info['group_id']  # This came from LoginName
            group_name = group_info['group_name']
            logger.info(f"Expanding group: {group_name} (Group ID from LoginName: {group_id}")
            
            # Get owners using the Group ID from LoginName
            owners, actual_group_title = get_group_owners(group_id)
            
            # Use the actual group title from Graph API if available, otherwise use SharePoint title
            group_display_title = actual_group_title if actual_group_title and actual_group_title != group_id else group_name
            
            if owners:
                logger.info(f"Found {len(owners)} owners from group {group_display_title}")
                # Filter out ignored users from group owners
                filtered_owners = []
                for owner in owners:
                    if should_ignore_user(owner.get('title'), owner.get('email')):
                        with stats_lock:
                            stats['ignored_users'] += 1
                        logger.info(f"Ignoring group owner: {owner.get('title')} ({owner.get('email')})")
                        continue
                    owner['group_title'] = group_display_title
                    filtered_owners.append(owner)
                
                if filtered_owners:
                    group_owners.extend(filtered_owners)
                    groups_info.append({
                        'group_id': group_id,
                        'group_title': group_display_title,
                        'member_count': len(filtered_owners)
                    })
                else:
                    logger.warning(f"All owners were ignored for group {group_display_title}")
                    groups_info.append({
                        'group_id': group_id,
                        'group_title': group_display_title,
                        'member_count': 0
                    })
            else:
                logger.warning(f"No owners found for group {group_display_title}")
                # Keep the group itself as an entry (but check if group should be ignored)
                if not should_ignore_group(group_display_title):
                    group_owners.append({
                        'user_id': group_id,
                        'title': group_display_title,
                        'email': '',
                        'login_name': '',
                        'is_site_admin': True,
                        'is_group_member': False,
                        'is_group_itself': True,
                        'group_id': group_id,
                        'group_title': group_display_title
                    })
                    groups_info.append({
                        'group_id': group_id,
                        'group_title': group_display_title,
                        'member_count': 0
                    })
        
        logger.info(f"Final - Direct Users: {len(direct_users)}, Group Owners: {len(group_owners)}")
        return direct_users, group_owners, groups_info
        
    except Exception as e:
        logger.exception("Failed to parse site users XML")
        return [], [], []

def get_site_admins(site_url: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], List[Dict[str, str]]]:
    """Get site administrators, separating direct users and group owners"""
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
            return [], [], []
        
        direct_users, group_owners, groups_info = parse_site_users_xml(response.text)
        return direct_users, group_owners, groups_info
        
    except Exception as e:
        logger.exception(f"Failed to get site admins from {site_url}")
        return [], [], []

def normalize_site_url(site_input: str) -> str:
    """Normalize site URL from various input formats"""
    site_input = site_input.strip()
    
    if site_input.startswith('http'):
        return site_input
    
    if not site_input.startswith('sites/'):
        site_input = f"sites/{site_input}"
    
    tenant_name = CONFIG['tenant_name'].split('.')[0]
    return f"https://{tenant_name}.sharepoint.com/{site_input}"

def process_site(site_entry: Dict[str, str]) -> Dict[str, Any]:
    """Process a single site - used by thread pool"""
    site_input = site_entry.get('site_url') or site_entry.get('Site URL') or site_entry.get('Site') or site_entry.get('site')
    site_name = site_entry.get('site_name') or site_entry.get('Site Name') or site_entry.get('Name') or site_input
    
    result = {
        'site_input': site_input,
        'site_name': site_name,
        'success': False,
        'data': None,
        'error': None,
        'additional_info': site_entry.get('additional_info', '')
    }
    
    if not site_input:
        result['error'] = "No site URL provided"
        return result
    
    try:
        site_url = normalize_site_url(site_input)
        logger.info(f"Processing site: {site_url}")
        
        direct_users, group_owners, groups_info = get_site_admins(site_url)
        
        # Extract direct user details
        direct_emails = [u.get('email', u.get('login_name', '')) for u in direct_users 
                       if u.get('email') or u.get('login_name')]
        direct_names = [u.get('title', u.get('login_name', '')) for u in direct_users 
                      if u.get('title') or u.get('login_name')]
        
        # Extract group owner details
        owner_emails = [o.get('email', o.get('login_name', '')) for o in group_owners 
                      if o.get('email') or o.get('login_name')]
        owner_names = [o.get('title', o.get('login_name', '')) for o in group_owners 
                     if o.get('title') or o.get('login_name')]
        
        # Get group titles and their member counts
        group_titles = [g.get('group_title', '') for g in groups_info if g.get('group_title')]
        group_member_counts = [str(g.get('member_count', 0)) for g in groups_info]
        
        total_count = len(direct_users) + len(group_owners)
        
        result['success'] = True
        result['data'] = {
            'Site URL': site_url,
            'Site Name': site_name,
            'Direct Admin Emails': ', '.join(direct_emails) if direct_emails else '',
            'Direct Admin Names': ', '.join(direct_names) if direct_names else '',
            'Direct Admin Count': len(direct_users),
            'Group Owners Emails': ', '.join(owner_emails) if owner_emails else '',
            'Group Owners Names': ', '.join(owner_names) if owner_names else '',
            'Group Owners Count': len(group_owners),
            'Total Admin Count': total_count,
            'Group Titles': ' | '.join(group_titles) if group_titles else 'None',
            'Group Members Count': ' | '.join(group_member_counts) if group_member_counts else '0',
            'Error': '',
            'Additional Info': site_entry.get('additional_info', '')
        }
        
        # Update statistics
        with stats_lock:
            stats['processed'] += 1
            stats['total_direct_users'] += len(direct_users)
            stats['total_group_owners'] += len(group_owners)
            stats['total_groups'] += len(groups_info)
        
        logger.info(f"✅ Completed: {site_url} - Direct: {len(direct_users)}, Group Owners: {len(group_owners)}, Groups: {len(groups_info)}")
        
    except Exception as e:
        logger.error(f"Error processing {site_input}: {str(e)}")
        result['error'] = str(e)
        with stats_lock:
            stats['failed'] += 1
        
        result['data'] = {
            'Site URL': site_input,
            'Site Name': site_name,
            'Direct Admin Emails': '',
            'Direct Admin Names': '',
            'Direct Admin Count': 0,
            'Group Owners Emails': '',
            'Group Owners Names': '',
            'Group Owners Count': 0,
            'Total Admin Count': 0,
            'Group Titles': '',
            'Group Members Count': '',
            'Error': str(e),
            'Additional Info': site_entry.get('additional_info', '')
        }
    
    return result

def generate_report(site_list: List[Dict[str, str]], output_file: str = "sharepoint_admin_report.csv"):
    """Generate CSV report using multi-threading"""
    
    # Update total sites
    stats['total_sites'] = len(site_list)
    
    print(f"\n🚀 Starting processing with {CONFIG['max_workers']} threads...")
    print(f"📋 Total sites to process: {len(site_list)}")
    print(f"👤 Ignoring {len(CONFIG['ignore_users'])} users")
    print(f"📁 Ignoring {len(CONFIG['ignore_groups'])} groups")
    print()
    
    results = []
    start_time = time.time()
    
    # Use ThreadPoolExecutor for parallel processing
    with ThreadPoolExecutor(max_workers=CONFIG['max_workers']) as executor:
        # Submit all tasks
        future_to_site = {executor.submit(process_site, site): site for site in site_list}
        
        # Process completed tasks
        completed = 0
        for future in as_completed(future_to_site):
            completed += 1
            result = future.result()
            results.append(result)
            
            # Show progress
            status = "✅" if result['success'] else "❌"
            print(f"[{completed}/{len(site_list)}] {status} {result.get('site_input', 'Unknown')}")
    
    end_time = time.time()
    elapsed_time = end_time - start_time
    
    # Sort results by site name for consistent output
    results.sort(key=lambda x: x.get('site_name', ''))
    
    # Write to CSV
    try:
        output_headers = [
            'Site URL',
            'Site Name',
            'Direct Admin Emails',
            'Direct Admin Names',
            'Direct Admin Count',
            'Group Owners Emails',
            'Group Owners Names',
            'Group Owners Count',
            'Total Admin Count',
            'Group Titles',
            'Group Members Count',
            'Error',
            'Additional Info'
        ]
        
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=output_headers)
            writer.writeheader()
            for result in results:
                if result.get('data'):
                    writer.writerow(result['data'])
        
        logger.info(f"Report generated successfully: {output_file}")
        
        # Print summary
        print(f"\n{'='*80}")
        print("📊 PROCESSING SUMMARY")
        print(f"{'='*80}")
        print(f"✅ Report generated: {output_file}")
        print(f"⏱️  Time taken: {elapsed_time:.2f} seconds")
        print(f"\n📈 Statistics:")
        print(f"   Total sites: {stats['total_sites']}")
        print(f"   ✅ Successful: {stats['processed']}")
        print(f"   ❌ Failed: {stats['failed']}")
        print(f"   👤 Direct Users: {stats['total_direct_users']}")
        print(f"   👥 Group Owners: {stats['total_group_owners']}")
        print(f"   📁 Groups Found: {stats['total_groups']}")
        print(f"   📊 Total Admins: {stats['total_direct_users'] + stats['total_group_owners']}")
        if stats['ignored_users'] > 0 or stats['ignored_groups'] > 0:
            print(f"\n🚫 Ignored Items:")
            if stats['ignored_users'] > 0:
                print(f"   👤 Ignored Users: {stats['ignored_users']}")
            if stats['ignored_groups'] > 0:
                print(f"   📁 Ignored Groups: {stats['ignored_groups']}")
        print(f"{'='*80}")
        
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
            print()
            
            for row in reader:
                site_url = None
                for col in ['Site URL', 'Site', 'site_url', 'site', 'URL', 'url']:
                    if col in row and row[col]:
                        site_url = row[col]
                        break
                
                if site_url:
                    site_entry = {'site_url': site_url}
                    
                    for col in ['Site Name', 'Name', 'site_name', 'name']:
                        if col in row and row[col]:
                            site_entry['site_name'] = row[col]
                            break
                    
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
    
    print("=" * 80)
    print("🚀 SharePoint Site Admin Report Generator (Multi-Threaded)")
    print("=" * 80)
    print()
    print("🔍 This tool will:")
    print("   - Find direct site administrators (PrincipalType=1)")
    print("   - Find groups with admin rights (PrincipalType=4)")
    print("   - Extract Group ID from LoginName field")
    print("   - Expand groups to get all owners via Graph API")
    print("   - Include group titles in the report")
    print("   - Ignore specified users and groups (exact matches only)")
    print("   - Process multiple sites in parallel for speed")
    print()
    print(f"⚙️  Using {CONFIG['max_workers']} threads for parallel processing")
    print()
    
    # Show current ignore configuration
    if CONFIG['ignore_users']:
        print(f"👤 Users to ignore: {', '.join(CONFIG['ignore_users'])}")
        print()
    
    if CONFIG['ignore_groups']:
        print(f"📁 Groups to ignore: {', '.join(CONFIG['ignore_groups'])}")
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
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = f"sharepoint_admin_report_{timestamp}.csv"
    
    generate_report(site_list, output_file)

if __name__ == "__main__":
    main()
