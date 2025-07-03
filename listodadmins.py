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

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('onedrive_admin_report.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Configuration
CONFIG = {


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
            logger.info("Successfully obtained token using certificate")
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
            logger.info("Successfully obtained token using client secret")
            return token_response.json()["access_token"]
        else:
            logger.error(f"Client secret token request failed: {token_response.text}")
            return None
            
    except Exception as e:
        logger.exception("Client secret authentication failed")
        return None

def parse_site_users_xml(xml_content):
    """Parse SharePoint site users XML and return all admins"""
    try:
        logger.debug("Parsing XML for site admins")
        
        # Parse XML
        root = ET.fromstring(xml_content)
        
        # Define namespaces
        namespaces = {
            'd': 'http://schemas.microsoft.com/ado/2007/08/dataservices',
            'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata',
            'atom': 'http://www.w3.org/2005/Atom'
        }
        
        # Find all entries
        entries = root.findall('.//atom:entry', namespaces)
        logger.debug(f"Found {len(entries)} entries in XML")
        
        admins = []
        
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
                    
                    # Get values safely
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
        raise Exception(f"Failed to parse XML response: {str(e)}")

def get_site_admins(site_url):
    """Get all site administrators for a SharePoint site/OneDrive"""
    try:
        # Ensure site URL ends properly
        if not site_url.endswith('/'):
            site_url += '/'
        
        # Construct SharePoint API URL
        site_users_url = f"{site_url}_api/web/siteusers"
        logger.info(f"Fetching site admins from URL: {site_users_url}")
        
        # Get SharePoint token
        sharepoint_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sharepoint_token:
            sharepoint_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Call SharePoint API to get site users
        sharepoint_headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/xml"
        }
        
        logger.info(f"Making request to: {site_users_url}")
        site_users_response = requests.get(site_users_url, headers=sharepoint_headers)
        
        if site_users_response.status_code != 200:
            logger.error(f"SharePoint API failed: {site_users_response.status_code} - {site_users_response.text}")
            raise Exception(f"Failed to get site users from {site_users_url}: {site_users_response.text}")
        
        # Parse XML and find all admins
        admins = parse_site_users_xml(site_users_response.text)
        logger.info(f"Found admins: {[a['email'] or a['login_name'] for a in admins]}")
        return admins
        
    except Exception as e:
        logger.exception(f"Failed to get site admins from {site_url}")
        raise

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
                    
        logger.info(f"Read {len(urls)} OneDrive URLs from {file_path}")
        return urls
    except Exception as e:
        logger.exception(f"Failed to read URLs from {file_path}")
        raise

def extract_owner_from_url(onedrive_url):
    """Extract owner information from OneDrive URL"""
    try:
        # Extract the personal folder part from OneDrive URL
        # Example: https://tenant-my.sharepoint.com/personal/username_domain_com/
        if '/personal/' in onedrive_url:
            parts = onedrive_url.split('/personal/')
            if len(parts) > 1:
                owner_part = parts[1].split('/')[0]
                # Convert back to email format
                owner_email = owner_part.replace('_', '.').replace('@', '_', 1).replace('_', '@', 1)
                return owner_email
        return "Unknown"
    except Exception as e:
        logger.warning(f"Could not extract owner from URL {onedrive_url}: {e}")
        return "Unknown"

def save_to_csv(results, filename):
    """Save results to CSV file"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['onedrive_url', 'owner', 'admin_count', 'admin_name', 'admin_email', 'admin_login_name', 'status', 'error']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for result in results:
                if result['status'] == 'success':
                    if result['admins']:
                        for admin in result['admins']:
                            writer.writerow({
                                'onedrive_url': result['onedrive_url'],
                                'owner': result['owner'],
                                'admin_count': len(result['admins']),
                                'admin_name': admin.get('title', ''),
                                'admin_email': admin.get('email', ''),
                                'admin_login_name': admin.get('login_name', ''),
                                'status': result['status'],
                                'error': ''
                            })
                    else:
                        writer.writerow({
                            'onedrive_url': result['onedrive_url'],
                            'owner': result['owner'],
                            'admin_count': 0,
                            'admin_name': '',
                            'admin_email': '',
                            'admin_login_name': '',
                            'status': result['status'],
                            'error': 'No admins found'
                        })
                else:
                    writer.writerow({
                        'onedrive_url': result['onedrive_url'],
                        'owner': result['owner'],
                        'admin_count': 0,
                        'admin_name': '',
                        'admin_email': '',
                        'admin_login_name': '',
                        'status': result['status'],
                        'error': result.get('error', '')
                    })
        
        logger.info(f"Results saved to {filename}")
        
    except Exception as e:
        logger.exception(f"Failed to save results to {filename}")
        raise

def process_onedrive_urls(urls):
    """Process a list of OneDrive URLs and get admin information"""
    results = []
    
    for i, url in enumerate(urls, 1):
        logger.info(f"Processing {i}/{len(urls)}: {url}")
        
        try:
            # Clean up URL
            if url.endswith('/Documents'):
                url = url[:-10]
            
            # Extract owner information
            owner = extract_owner_from_url(url)
            
            # Get site admins
            admins = get_site_admins(url)
            
            result = {
                'onedrive_url': url,
                'owner': owner,
                'status': 'success',
                'admins': admins
            }
            
            logger.info(f"✓ Successfully processed {url} - Found {len(admins)} admins")
            
        except Exception as e:
            logger.error(f"✗ Failed to process {url}: {str(e)}")
            result = {
                'onedrive_url': url,
                'owner': extract_owner_from_url(url),
                'status': 'error',
                'error': str(e),
                'admins': []
            }
        
        results.append(result)
        
        # Add a small delay to avoid rate limiting
        time.sleep(0.5)
    
    return results

def print_summary(results):
    """Print a summary of the results"""
    total_urls = len(results)
    successful = sum(1 for r in results if r['status'] == 'success')
    failed = total_urls - successful
    total_admins = sum(len(r['admins']) for r in results if r['status'] == 'success')
    
    print("\n" + "="*50)
    print("PROCESSING SUMMARY")
    print("="*50)
    print(f"Total OneDrive URLs processed: {total_urls}")
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")
    print(f"Total administrators found: {total_admins}")
    print("="*50)
    
    if failed > 0:
        print("\nFailed URLs:")
        for result in results:
            if result['status'] == 'error':
                print(f"  - {result['onedrive_url']}: {result['error']}")

def main():
    """Main function to process OneDrive URLs and generate admin report"""
    try:
        # Input CSV file containing OneDrive URLs with 'Web URL' header
        input_file = "onedrive_urls.csv"
        
        # Output files
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
        
        print(f"Found {len(urls)} OneDrive URLs to process")
        
        # Process URLs
        results = process_onedrive_urls(urls)
        
        # Save results to CSV
        save_to_csv(results, output_csv)
        
        # Print summary
        print_summary(results)
        
        print(f"\nDetailed results saved to: {output_csv}")
        print(f"Log file saved to: onedrive_admin_report.log")
        
    except Exception as e:
        logger.exception("Script execution failed")
        print(f"Script failed: {str(e)}")

if __name__ == "__main__":
    main()
