#!/usr/bin/env python3
"""
SharePoint Site Admin Report Generator
Generates a CSV report of site administrators for a list of SharePoint sites
Supports input from CSV file or interactive input
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
from typing import List, Dict, Any, Optional

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

def get_sharepoint_token() -> Optional[str]:
    """Get SharePoint access token"""
    token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
    if not token:
        token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
    return token

def parse_site_users_xml(xml_content: str) -> List[Dict[str, Any]]:
    """Parse SharePoint site users XML and return all admins"""
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
        
        for entry in entries:
            content = entry.find('atom:content', namespaces)
            if content is not None:
                properties = content.find('m:properties', namespaces)
                if properties is not None:
                    user_id_elem = properties.find('d:Id', namespaces)
                    title_elem = properties.find('d:Title', namespaces)  
                    email_elem = properties.find('d:Email', namespaces)
                    login_name_elem = properties.find('d:LoginName', namespaces)
                    is_site_admin_elem = properties.find('d:IsSiteAdmin', namespaces)
                    
                    user_id = user_id_elem.text if user_id_elem is not None else None
                    title = title_elem.text if title_elem is not None else None
                    email = email_elem.text if email_elem is not None else None
                    login_name = login_name_elem.text if login_name_elem is not None else None
                    is_site_admin = is_site_admin_elem.text == 'true' if is_site_admin_elem is not None else False
                    
                    if is_site_admin:
                        admins.append({
                            'user_id': user_id,
                            'title': title,
                            'email': email,
                            'login_name': login_name,
                            'is_site_admin': is_site_admin
                        })
        
        logger.info(f"Found {len(admins)} site administrators")
        return admins
        
    except Exception as e:
        logger.exception("Failed to parse site users XML")
        return []

def get_site_admins(site_url: str) -> List[Dict[str, Any]]:
    """Get all site administrators for a SharePoint site"""
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
    """Generate CSV report for a list of SharePoint sites"""
    
    report_data = []
    
    for site_entry in site_list:
        site_input = site_entry.get('site_url') or site_entry.get('Site URL') or site_entry.get('Site') or site_entry.get('site')
        site_name = site_entry.get('site_name') or site_entry.get('Site Name') or site_entry.get('Name') or site_input
        
        if not site_input:
            logger.warning(f"Skipping entry with no site URL: {site_entry}")
            continue
        
        try:
            site_url = normalize_site_url(site_input)
            logger.info(f"Processing site: {site_url}")
            
            admins = get_site_admins(site_url)
            
            if not admins:
                report_data.append({
                    'Site URL': site_url,
                    'Site Name': site_name,
                    'Admin Emails': '',
                    'Admin Names': '',
                    'Admin Count': 0,
                    'Error': 'No admins found or site inaccessible',
                    'Additional Info': site_entry.get('additional_info', '')
                })
            else:
                admin_emails = [admin.get('email', admin.get('login_name', '')) for admin in admins if admin.get('email') or admin.get('login_name')]
                admin_names = [admin.get('title', admin.get('login_name', '')) for admin in admins if admin.get('title') or admin.get('login_name')]
                
                report_data.append({
                    'Site URL': site_url,
                    'Site Name': site_name,
                    'Admin Emails': ', '.join(admin_emails) if admin_emails else '',
                    'Admin Names': ', '.join(admin_names) if admin_names else '',
                    'Admin Count': len(admins),
                    'Error': '',
                    'Additional Info': site_entry.get('additional_info', '')
                })
                
        except Exception as e:
            logger.error(f"Error processing {site_input}: {str(e)}")
            report_data.append({
                'Site URL': site_input,
                'Site Name': site_name,
                'Admin Emails': '',
                'Admin Names': '',
                'Admin Count': 0,
                'Error': str(e),
                'Additional Info': site_entry.get('additional_info', '')
            })
    
    # Write to CSV
    try:
        output_headers = ['Site URL', 'Site Name', 'Admin Emails', 'Admin Names', 'Admin Count', 'Error', 'Additional Info']
        
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
        
    except Exception as e:
        logger.error(f"Failed to write CSV report: {str(e)}")
        raise

def read_sites_from_csv(csv_file: str) -> List[Dict[str, str]]:
    """Read sites from input CSV file"""
    sites = []
    
    try:
        with open(csv_file, 'r', encoding='utf-8') as file:
            reader = csv.DictReader(file)
            
            # Check if headers exist
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
    print("SharePoint Site Admin Report Generator")
    print("=" * 70)
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
