{
    "sharepoint_main_site": "https://geekbyteonline.sharepoint.com/sites/New365",
    "tenant_name": "geekbyteonline.onmicrosoft.com",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "scope_sharepoint": "https://geekbyteonline.sharepoint.com/.default",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem"
}



import requests
import json
from urllib.parse import urljoin
import csv
from datetime import datetime
import uuid
import base64
import time
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
        return config
    except Exception as e:
        print(f"Error loading configuration: {str(e)}")
        raise

def load_certificate_and_key(private_key_path, certificate_path):
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

def authenticate_sharepoint(config):
    """Authenticate and get SharePoint access token using certificate"""
    try:
        print("Loading certificate and private key...")
        certificate, private_key = load_certificate_and_key(
            config["private_key_path"], 
            config["certificate_path"]
        )
        
        print("Generating JWT token...")
        jwt = get_jwt_token(
            certificate, 
            private_key, 
            config["tenant_name"], 
            config["app_id"], 
            config["scope_sharepoint"]
        )
        
        print("Getting access token...")
        access_token = get_access_token(
            jwt, 
            config["tenant_name"], 
            config["app_id"], 
            config["scope_sharepoint"]
        )
        
        print("Successfully authenticated to SharePoint")
        return access_token
        
    except Exception as e:
        print(f"Authentication failed: {str(e)}")
        return None

def make_sharepoint_request(url, access_token, method='GET', headers=None):
    """Make a request to SharePoint REST API"""
    default_headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json"
    }
    
    if headers:
        default_headers.update(headers)
    
    try:
        response = requests.request(method, url, headers=default_headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Request failed for {url}: {str(e)}")
        return None

def get_all_subsites(site_url, access_token):
    """Get all subsites from SharePoint site"""
    subsites_url = urljoin(site_url + "/", "_api/web/webs")
    response = make_sharepoint_request(subsites_url, access_token)
    
    if response and 'd' in response and 'results' in response['d']:
        return response['d']['results']
    return []

def get_subsite_details(subsite_url, access_token):
    """Get additional details for a subsite including WelcomePage"""
    details_url = urljoin(subsite_url + "/", "_api/web")
    response = make_sharepoint_request(details_url, access_token)
    
    if response and 'd' in response:
        return response['d']
    return {}

def process_subsite(subsite, access_token):
    """Extract relevant details from a subsite including WelcomePage"""
    # Get additional details for the subsite
    subsite_details = get_subsite_details(subsite['Url'], access_token)
    
    details = {
        'Title': subsite.get('Title', 'N/A'),
        'URL': subsite.get('Url', 'N/A'),
        'Description': subsite.get('Description', 'N/A'),
        'Created': subsite.get('Created', 'N/A'),
        'Language': subsite.get('Language', 'N/A'),
        'WebTemplate': subsite.get('WebTemplate', 'N/A'),
        'Configuration': subsite.get('Configuration', 'N/A'),
        'WelcomePage': subsite_details.get('WelcomePage', 'N/A')
    }
    
    return details

def main():
    # Load configuration from JSON file
    try:
        config = load_config()
    except Exception as e:
        print(f"Failed to load configuration: {str(e)}")
        return
    
    # Authenticate to SharePoint
    access_token = authenticate_sharepoint(config)
    if not access_token:
        print("Failed to authenticate to SharePoint. Exiting.")
        return
    
    # Get all subsites
    print("Retrieving subsites...")
    subsites = get_all_subsites(config["sharepoint_main_site"], access_token)
    
    if not subsites:
        print("No subsites found.")
        return
    
    # Prepare CSV output
    output_csv = f"SharePointSubsites_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.csv"
    
    with open(output_csv, 'w', newline='', encoding='utf-8') as csvfile:
        fieldnames = [
            'Title', 'URL', 'Description', 'Created', 
            'Language', 'WebTemplate', 'Configuration', 'WelcomePage'
        ]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        
        # Process each subsite
        for i, subsite in enumerate(subsites, 1):
            try:
                print(f"Processing subsite {i}/{len(subsites)}: {subsite.get('Title', 'Unknown')}")
                details = process_subsite(subsite, access_token)
                writer.writerow(details)
                
                # Print to console
                print(f"  - Title: {details['Title']}")
                print(f"  - URL: {details['URL']}")
                print(f"  - WelcomePage: {details['WelcomePage']}")
                
            except Exception as e:
                print(f"Error processing subsite: {str(e)}")
    
    print(f"\nSubsites report generated successfully: {output_csv}")

if __name__ == "__main__":
    main()
