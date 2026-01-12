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
    "tenant_id": "0e439a1fb-9e6b-4e582e203607",
    "tenant_name": "geekbyteonline.onmicrosoft.com",
    "app_id": "73efa35d-618a977eb149",
    "client_secret": "t",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    "scopes": {
        "graph": "https://graph.microsoft.com/.default",
        "sharepoint": "https://ne.sharepoint.com/.default"
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
        is_site_admin_elem = properties.find('.//d:IsSiteAdmin', ns)
        
        # Get values safely
        owner_info = {
            'user_id': user_id_elem.text if user_id_elem is not None else None,
            'title': title_elem.text if title_elem is not None else None,
            'email': email_elem.text if email_elem is not None else None,
            'login_name': login_name_elem.text if login_name_elem is not None else None,
            'user_principal_name': user_principal_name_elem.text if user_principal_name_elem is not None else None,
            'is_site_admin': is_site_admin_elem.text == 'true' if is_site_admin_elem is not None else False
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
    """Get all site users for a SharePoint site"""
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

def get_all_admins(site_url):
    """Get all site administrators for a SharePoint site"""
    try:
        print(f"Getting admins for site: {site_url}")
        
        # Get all site users
        all_users = get_site_users(site_url)
        
        # Filter only site administrators
        admins = [user for user in all_users if user['is_site_admin']]
        
        print(f"Found {len(admins)} administrators for {site_url}")
        return admins
        
    except Exception as e:
        logger.exception(f"Failed to get admins from {site_url}")
        raise

def format_admins_for_csv(admins):
    """Format administrators list into comma-separated strings for CSV output"""
    if not admins:
        return {
            'admin_names': '',
            'admin_emails': '',
            'admin_login_names': '',
            'admin_upns': '',
            'admin_count': 0
        }
    
    # Extract all admin details
    admin_names = []
    admin_emails = []
    admin_login_names = []
    admin_upns = []
    
    for admin in admins:
        # Add name/display name
        if admin.get('title'):
            admin_names.append(str(admin['title']))
        elif admin.get('user_principal_name'):
            admin_names.append(str(admin['user_principal_name']))
        elif admin.get('email'):
            admin_names.append(str(admin['email']))
        elif admin.get('login_name'):
            admin_names.append(str(admin['login_name']))
        
        # Add email
        if admin.get('email'):
            admin_emails.append(str(admin['email']))
        
        # Add login name
        if admin.get('login_name'):
            admin_login_names.append(str(admin['login_name']))
        
        # Add user principal name
        if admin.get('user_principal_name'):
            admin_upns.append(str(admin['user_principal_name']))
    
    return {
        'admin_names': ', '.join(admin_names),
        'admin_emails': ', '.join(admin_emails),
        'admin_login_names': ', '.join(admin_login_names),
        'admin_upns': ', '.join(admin_upns),
        'admin_count': len(admins)
    }

def read_sharepoint_urls_from_csv(file_path):
    """Read SharePoint site URLs from a CSV file"""
    urls = []
    try:
        with open(file_path, 'r', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            
            # Look for common column names
            possible_columns = ['Site URL', 'URL', 'Web URL', 'SharePoint URL', 'Site']
            
            found_column = None
            for col in possible_columns:
                if col in reader.fieldnames:
                    found_column = col
                    break
            
            if not found_column:
                # Try case-insensitive search
                for col in reader.fieldnames:
                    if 'url' in col.lower() or 'site' in col.lower():
                        found_column = col
                        break
            
            if not found_column:
                raise Exception(f"CSV file must have a URL column. Found columns: {reader.fieldnames}")
            
            print(f"Using column '{found_column}' for SharePoint URLs")
            
            for row in reader:
                url = row[found_column].strip()
                if url:  # Skip empty URLs
                    urls.append(url)
                    
        print(f"Read {len(urls)} SharePoint site URLs from {file_path}")
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
                'site_url', 'status', 'error',
                'admin_count',
                'admin_names', 'admin_emails', 'admin_login_names', 'admin_upns'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            if not file_exists:
                writer.writeheader()
            
            for result in results:
                if result['status'] == 'success':
                    formatted_admins = format_admins_for_csv(result['admins'])
                    
                    writer.writerow({
                        'site_url': result['site_url'],
                        'status': result['status'],
                        'error': '',
                        'admin_count': formatted_admins['admin_count'],
                        'admin_names': formatted_admins['admin_names'],
                        'admin_emails': formatted_admins['admin_emails'],
                        'admin_login_names': formatted_admins['admin_login_names'],
                        'admin_upns': formatted_admins['admin_upns']
                    })
                else:
                    # Write error row
                    writer.writerow({
                        'site_url': result['site_url'],
                        'status': result['status'],
                        'error': str(result.get('error', '')) if result.get('error') is not None else '',
                        'admin_count': 0,
                        'admin_names': '',
                        'admin_emails': '',
                        'admin_login_names': '',
                        'admin_upns': ''
                    })
                
                # Flush to ensure data is written immediately
                csvfile.flush()
                
    except Exception as e:
        print(f"Failed to update CSV file {filename}: {str(e)}")

def process_sharepoint_sites(urls, output_csv):
    """Process a list of SharePoint site URLs and get all administrators"""
    results = []
    
    for i, url in enumerate(urls, 1):
        print(f"\nProcessing {i}/{len(urls)}: {url}")
        
        try:
            # Clean up URL if needed
            url = url.strip()
            if not url.startswith('http'):
                print(f"Warning: {url} doesn't start with http/https, skipping")
                continue
                
            # Ensure URL ends with /
            if not url.endswith('/'):
                url += '/'
            
            # Get all admins for the site
            admins = get_all_admins(url)
            
            result = {
                'site_url': url,
                'status': 'success',
                'admins': admins
            }
            
            print(f"✓ Success: {url}")
            print(f"  Found {len(admins)} administrators")
            
            if admins:
                for admin in admins:
                    admin_info = admin.get('email') or admin.get('user_principal_name') or admin.get('login_name') or admin.get('title')
                    print(f"    - {admin_info}")
            
        except Exception as e:
            print(f"✗ Failed: {url}: {str(e)}")
            result = {
                'site_url': url,
                'status': 'error',
                'error': str(e),
                'admins': []
            }
        
        results.append(result)
        
        # Update CSV dynamically after each result
        update_csv_dynamically([result], output_csv)
        
        # Add a small delay to avoid rate limiting
        time.sleep(0.5)
    
    return results

def print_summary(results):
    """Print a summary of the results"""
    total_sites = len(results)
    successful = sum(1 for r in results if r['status'] == 'success')
    failed = total_sites - successful
    total_admins = sum(len(r['admins']) for r in results if r['status'] == 'success')
    
    print("\n" + "="*60)
    print("PROCESSING SUMMARY")
    print("="*60)
    print(f"Total SharePoint sites processed: {total_sites}")
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")
    print(f"Total administrators found: {total_admins}")
    print("="*60)
    
    if failed > 0:
        print("\nFailed sites:")
        for result in results:
            if result['status'] == 'error':
                print(f"  - {result['site_url']}: {result['error']}")

def create_sample_input_file():
    """Create a sample input CSV file"""
    sample_data = [
        ["Site URL"],
        ["https://geekbyteonline.sharepoint.com/sites/YourSiteName1"],
        ["https://geekbyteonline.sharepoint.com/sites/YourSiteName2"],
        ["https://geekbyteonline.sharepoint.com/sites/TeamSite"],
        ["https://geekbyteonline.sharepoint.com/sites/ProjectSite"]
    ]
    
    input_file = "sharepoint_sites.csv"
    
    with open(input_file, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerows(sample_data)
    
    print(f"Created sample input file: {input_file}")
    print("Please add your SharePoint site URLs to this CSV file and run the script again.")
    print("The CSV should have a 'Site URL' column (or similar) containing the SharePoint site URLs.")
    return input_file

def main():
    """Main function to process SharePoint sites and generate admin report"""
    try:
        # Input CSV file containing SharePoint site URLs
        input_file = "sharepoint_sites.csv"
        
        # Check if input file exists
        if not os.path.exists(input_file):
            input_file = create_sample_input_file()
            return
        
        # Output file
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_csv = f"sharepoint_admins_report_{timestamp}.csv"
        
        # Read SharePoint site URLs from CSV
        urls = read_sharepoint_urls_from_csv(input_file)
        
        if not urls:
            print(f"No URLs found in {input_file}. Please add SharePoint site URLs and try again.")
            return
        
        print(f"\nStarting processing of {len(urls)} SharePoint sites...")
        print(f"Results will be saved dynamically to: {output_csv}")
        print("-" * 60)
        
        # Process URLs with dynamic CSV updates
        results = process_sharepoint_sites(urls, output_csv)
        
        # Print final summary
        print_summary(results)
        
        print(f"\nFinal results saved to: {output_csv}")
        print("\nCSV Columns:")
        print("  - site_url: The SharePoint site URL")
        print("  - status: Success or error")
        print("  - error: Error message (if any)")
        print("  - admin_count: Number of administrators found")
        print("  - admin_names: All admin display names (comma-separated)")
        print("  - admin_emails: All admin emails (comma-separated)")
        print("  - admin_login_names: All admin login names (comma-separated)")
        print("  - admin_upns: All admin UPNs (comma-separated)")
        
    except Exception as e:
        print(f"Script failed: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
