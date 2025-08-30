import requests
import json
import uuid
import base64
import time
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend
import logging
import os
import xml.etree.ElementTree as ET
import csv
from datetime import datetime

# Configure minimal logging (only to console)
logging.basicConfig(
    level=logging.WARNING,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# Configuration
CONFIG = {
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "tenant_name": "geekbyteonline.onmicrosoft.com",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "client_secret": "t",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    "scopes": {
        "graph": "https://graph.microsoft.com/.default",
        "sharepoint": "https://geekbyteonline.sharepoint.com/.default"
    }
}

# Token cache
TOKEN_CACHE = {
    "graph": {"token": None, "expires": 0},
    "sharepoint": {"token": None, "expires": 0}
}

def get_token_with_certificate(scope):
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
            return token_response.json()["access_token"]
        else:
            logger.error(f"Certificate token request failed: {token_response.text}")
            return None
            
    except Exception as e:
        logger.exception("Certificate authentication failed")
        return None

def get_token_with_secret(scope):
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
            return token_response.json()["access_token"]
        else:
            logger.error(f"Client secret token request failed: {token_response.text}")
            return None
            
    except Exception as e:
        logger.exception("Client secret authentication failed")
        return None

def get_cached_token(scope_type):
    """Get cached token if it's still valid, otherwise get a new one"""
    cache = TOKEN_CACHE[scope_type]
    
    # If token exists and hasn't expired (with 5 minute buffer)
    if cache["token"] and cache["expires"] > time.time() + 300:
        return cache["token"]
    
    # Get new token
    scope = CONFIG['scopes'][scope_type]
    
    # Try certificate first
    token = get_token_with_certificate(scope)
    if not token:
        # Fall back to client secret
        token = get_token_with_secret(scope)
    
    if token:
        # Cache the token with expiration (assuming 1 hour lifetime)
        cache["token"] = token
        cache["expires"] = time.time() + 3600
        return token
    
    return None

def get_site_owner(site_url):
    """Get site owner details from SharePoint site"""
    try:
        # Ensure site URL ends properly
        if not site_url.endswith('/'):
            site_url += '/'
        
        # Construct SharePoint API URL for site owner
        site_owner_url = f"{site_url}_api/site/owner"
        
        # Get SharePoint token
        sharepoint_token = get_cached_token("sharepoint")
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Call SharePoint API to get site owner
        sharepoint_headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/xml"
        }
        
        site_owner_response = requests.get(site_owner_url, headers=sharepoint_headers)
        
        if site_owner_response.status_code != 200:
            raise Exception(f"Failed to get site owner from {site_owner_url}: {site_owner_response.text}")
        
        # Parse XML response
        owner_info = parse_site_owner_xml(site_owner_response.text)
        return owner_info
        
    except Exception as e:
        logger.exception(f"Failed to get site owner from {site_url}")
        raise

def parse_site_owner_xml(xml_content):
    """Parse SharePoint site owner XML and return owner details"""
    try:
        # Parse XML
        root = ET.fromstring(xml_content)
        
        # Register namespaces to handle default namespace
        ns = {
            'atom': 'http://www.w3.org/2005/Atom',
            'd': 'http://schemas.microsoft.com/ado/2007/08/dataservices',
            'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'
        }
        
        # Find the content element
        content = root.find('.//atom:content', ns)
        if content is None:
            raise Exception("No content found in owner XML response")
        
        # Find properties
        properties = content.find('.//m:properties', ns)
        if properties is None:
            raise Exception("No properties found in owner XML response")
        
        # Extract owner details
        user_id_elem = properties.find('.//d:Id', ns)
        title_elem = properties.find('.//d:Title', ns)  
        email_elem = properties.find('.//d:Email', ns)
        login_name_elem = properties.find('.//d:LoginName', ns)
        user_principal_name_elem = properties.find('.//d:UserPrincipalName', ns)
        
        # Get values safely
        owner_info = {
            'user_id': user_id_elem.text if user_id_elem is not None else None,
            'title': title_elem.text if title_elem is not None else None,
            'email': email_elem.text if email_elem is not None else None,
            'login_name': login_name_elem.text if login_name_elem is not None else None,
            'user_principal_name': user_principal_name_elem.text if user_principal_name_elem is not None else None
        }
        
        return owner_info
        
    except Exception as e:
        logger.exception("Failed to parse site owner XML")
        raise Exception(f"Failed to parse owner XML response: {str(e)}")

def parse_site_users_xml(xml_content):
    """Parse SharePoint site users XML and return all users with admin privileges"""
    try:
        # Parse XML
        root = ET.fromstring(xml_content)
        
        # Register namespaces
        ns = {
            'atom': 'http://www.w3.org/2005/Atom',
            'd': 'http://schemas.microsoft.com/ado/2007/08/dataservices',
            'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'
        }
        
        # Find all entries
        entries = root.findall('.//atom:entry', ns)
        
        users = []
        
        for entry in entries:
            content = entry.find('.//atom:content', ns)
            if content is not None:
                properties = content.find('.//m:properties', ns)
                if properties is not None:
                    # Extract user details
                    user_id_elem = properties.find('.//d:Id', ns)
                    title_elem = properties.find('.//d:Title', ns)  
                    email_elem = properties.find('.//d:Email', ns)
                    login_name_elem = properties.find('.//d:LoginName', ns)
                    is_site_admin_elem = properties.find('.//d:IsSiteAdmin', ns)
                    user_principal_name_elem = properties.find('.//d:UserPrincipalName', ns)
                    
                    # Get values safely
                    user_id = user_id_elem.text if user_id_elem is not None else None
                    title = title_elem.text if title_elem is not None else None
                    email = email_elem.text if email_elem is not None else None
                    login_name = login_name_elem.text if login_name_elem is not None else None
                    is_site_admin = is_site_admin_elem.text == 'true' if is_site_admin_elem is not None else False
                    user_principal_name = user_principal_name_elem.text if user_principal_name_elem is not None else None
                    
                    users.append({
                        'user_id': user_id,
                        'title': title,
                        'email': email,
                        'login_name': login_name,
                        'is_site_admin': is_site_admin,
                        'user_principal_name': user_principal_name
                    })
        
        return users
        
    except Exception as e:
        logger.exception("Failed to parse site users XML")
        raise Exception(f"Failed to parse XML response: {str(e)}")

def get_site_users(site_url):
    """Get all site users for a SharePoint site/OneDrive"""
    try:
        # Ensure site URL ends properly
        if not site_url.endswith('/'):
            site_url += '/'
        
        # Construct SharePoint API URL
        site_users_url = f"{site_url}_api/web/siteusers"
        
        # Get SharePoint token
        sharepoint_token = get_cached_token("sharepoint")
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Call SharePoint API to get site users
        sharepoint_headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/xml"
        }
        
        site_users_response = requests.get(site_users_url, headers=sharepoint_headers)
        
        if site_users_response.status_code != 200:
            raise Exception(f"Failed to get site users from {site_users_url}: {site_users_response.text}")
        
        # Parse XML and get all users
        users = parse_site_users_xml(site_users_response.text)
        return users
        
    except Exception as e:
        logger.exception(f"Failed to get site users from {site_url}")
        raise

def filter_admins_exclude_owner(users, owner_info):
    """Filter out admins excluding the owner"""
    try:
        admins = [user for user in users if user['is_site_admin']]
        
        # Filter out owner from admins list
        owner_login = owner_info.get('login_name')
        owner_email = owner_info.get('email')
        owner_upn = owner_info.get('user_principal_name')
        
        other_admins = []
        for admin in admins:
            admin_login = admin.get('login_name')
            admin_email = admin.get('email')
            admin_upn = admin.get('user_principal_name')
            
            # Check if this admin is not the owner
            is_owner = False
            if owner_login and admin_login and owner_login == admin_login:
                is_owner = True
            elif owner_email and admin_email and owner_email == admin_email:
                is_owner = True
            elif owner_upn and admin_upn and owner_upn == admin_upn:
                is_owner = True
            
            if not is_owner:
                other_admins.append(admin)
        
        return other_admins
        
    except Exception as e:
        logger.exception("Failed to filter admins")
        return []

def read_onedrive_urls_from_csv(file_path):
    """Read OneDrive URLs from a CSV file with 'Web URL' header"""
    urls = []
    try:
        with open(file_path, 'r', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            
            # Check if 'Web URL' column exists
            if 'Web URL' not in reader.fieldnames:
                raise Exception(f"CSV file must have a 'Web URL' column. Found columns: {reader.fieldnames}")
            
            for row in reader:
                url = row['Web URL'].strip()
                if url:  # Skip empty URLs
                    urls.append(url)
                    
        print(f"Read {len(urls)} OneDrive URLs from {file_path}")
        return urls
    except Exception as e:
        print(f"Failed to read URLs from {file_path}: {str(e)}")
        raise

def update_csv_dynamically(results, filename):
    """Update CSV file dynamically as results come in"""
    try:
        # Check if file exists to determine if we need to write header
        file_exists = os.path.exists(filename)
        
        with open(filename, 'a', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'onedrive_url', 'status', 'error',
                'owner_name', 'owner_email', 'owner_login_name', 'owner_upn',
                'additional_admin_count',
                'admin_names', 'admin_emails', 'admin_login_names', 'admin_upns'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            if not file_exists:
                writer.writeheader()
            
            for result in results:
                if result['status'] == 'success':
                    owner = result.get('owner', {})
                    admins = result.get('additional_admins', [])
                    
                    # Combine all admin details into semicolon-separated strings
                    admin_names = []
                    admin_emails = []
                    admin_login_names = []
                    admin_upns = []
                    
                    for admin in admins:
                        admin_names.append(str(admin.get('title', '')) if admin.get('title') is not None else '')
                        admin_emails.append(str(admin.get('email', '')) if admin.get('email') is not None else '')
                        admin_login_names.append(str(admin.get('login_name', '')) if admin.get('login_name') is not None else '')
                        admin_upns.append(str(admin.get('user_principal_name', '')) if admin.get('user_principal_name') is not None else '')
                    
                    writer.writerow({
                        'onedrive_url': result['onedrive_url'],
                        'status': result['status'],
                        'error': '',
                        'owner_name': str(owner.get('title', '')) if owner.get('title') is not None else '',
                        'owner_email': str(owner.get('email', '')) if owner.get('email') is not None else '',
                        'owner_login_name': str(owner.get('login_name', '')) if owner.get('login_name') is not None else '',
                        'owner_upn': str(owner.get('user_principal_name', '')) if owner.get('user_principal_name') is not None else '',
                        'additional_admin_count': len(admins),
                        'admin_names': '; '.join(filter(None, admin_names)),
                        'admin_emails': '; '.join(filter(None, admin_emails)),
                        'admin_login_names': '; '.join(filter(None, admin_login_names)),
                        'admin_upns': '; '.join(filter(None, admin_upns))
                    })
                else:
                    # Write error row
                    writer.writerow({
                        'onedrive_url': result['onedrive_url'],
                        'status': result['status'],
                        'error': str(result.get('error', '')) if result.get('error') is not None else '',
                        'owner_name': '',
                        'owner_email': '',
                        'owner_login_name': '',
                        'owner_upn': '',
                        'additional_admin_count': 0,
                        'admin_names': '',
                        'admin_emails': '',
                        'admin_login_names': '',
                        'admin_upns': ''
                    })
                
                # Flush to ensure data is written immediately
                csvfile.flush()
                
    except Exception as e:
        print(f"Failed to update CSV file {filename}: {str(e)}")

def process_onedrive_urls(urls, output_csv):
    """Process a list of OneDrive URLs and get owner and admin information"""
    results = []
    
    for i, url in enumerate(urls, 1):
        print(f"Processing {i}/{len(urls)}: {url}")
        
        try:
            # Clean up URL
            if url.endswith('/Documents'):
                url = url[:-10]
            
            # Get site owner
            owner_info = get_site_owner(url)
            
            # Get all site users
            all_users = get_site_users(url)
            
            # Filter admins excluding owner
            additional_admins = filter_admins_exclude_owner(all_users, owner_info)
            
            result = {
                'onedrive_url': url,
                'status': 'success',
                'owner': owner_info,
                'additional_admins': additional_admins
            }
            
            print(f"✓ Success: {url} - Owner: {owner_info.get('email', 'Unknown')}, Additional admins: {len(additional_admins)}")
            
        except Exception as e:
            print(f"✗ Failed: {url}: {str(e)}")
            result = {
                'onedrive_url': url,
                'status': 'error',
                'error': str(e),
                'owner': {},
                'additional_admins': []
            }
        
        results.append(result)
        
        # Update CSV dynamically after each result
        update_csv_dynamically([result], output_csv)
        
        # Add a small delay to avoid rate limiting
        time.sleep(0.3)  # Reduced delay for faster processing
    
    return results

def print_summary(results):
    """Print a summary of the results"""
    total_urls = len(results)
    successful = sum(1 for r in results if r['status'] == 'success')
    failed = total_urls - successful
    total_additional_admins = sum(len(r['additional_admins']) for r in results if r['status'] == 'success')
    urls_with_additional_admins = sum(1 for r in results if r['status'] == 'success' and len(r['additional_admins']) > 0)
    
    print("\n" + "="*60)
    print("PROCESSING SUMMARY")
    print("="*60)
    print(f"Total OneDrive URLs processed: {total_urls}")
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")
    print(f"OneDrive sites with additional admins: {urls_with_additional_admins}")
    print(f"Total additional administrators found: {total_additional_admins}")
    print("="*60)
    
    if failed > 0:
        print("\nFailed URLs:")
        for result in results:
            if result['status'] == 'error':
                print(f"  - {result['onedrive_url']}: {result['error']}")
    
    if urls_with_additional_admins > 0:
        print("\nOneDrive sites with additional admins:")
        for result in results:
            if result['status'] == 'success' and len(result['additional_admins']) > 0:
                owner_email = result['owner'].get('email', 'Unknown')
                admin_count = len(result['additional_admins'])
                print(f"  - {result['onedrive_url']}")
                print(f"    Owner: {owner_email}")
                print(f"    Additional admins: {admin_count}")
                for admin in result['additional_admins']:
                    admin_email = admin.get('email', admin.get('login_name', 'Unknown'))
                    print(f"      - {admin_email}")

def main():
    """Main function to process OneDrive URLs and generate admin report"""
    try:
        # Input CSV file containing OneDrive URLs with 'Web URL' header
        input_file = "onedrive_urls.csv"
        
        # Output file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_csv = f"onedrive_admin_report_{timestamp}.csv"
        
        # Check if input file exists
        if not os.path.exists(input_file):
            # Create sample input CSV file
            sample_data = [
                ["Web URL"],
                ["https://tenant-my.sharepoint.com/personal/user1_domain_com/"],
                ["https://tenant-my.sharepoint.com/personal/user2_domain_com/Documents"],
                ["https://tenant-my.sharepoint.com/personal/user3_domain_com/"]
            ]
            
            with open(input_file, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerows(sample_data)
            
            print(f"Created sample input file: {input_file}")
            print("Please add your OneDrive URLs to this CSV file and run the script again.")
            print("The CSV should have a 'Web URL' column containing the OneDrive URLs.")
            return
        
        # Read OneDrive URLs from CSV
        urls = read_onedrive_urls_from_csv(input_file)
        
        if not urls:
            print(f"No URLs found in {input_file}. Please add OneDrive URLs and try again.")
            return
        
        print(f"Starting processing of {len(urls)} OneDrive URLs...")
        print(f"Results will be saved dynamically to: {output_csv}")
        print("-" * 60)
        
        # Process URLs with dynamic CSV updates
        results = process_onedrive_urls(urls, output_csv)
        
        # Print final summary
        print_summary(results)
        
        print(f"\nFinal results saved to: {output_csv}")
        
    except Exception as e:
        print(f"Script failed: {str(e)}")

if __name__ == "__main__":
    main()
