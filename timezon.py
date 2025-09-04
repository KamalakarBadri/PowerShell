import json
import uuid
import base64
import time
import requests
import csv
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend

# Configuration
tenant_name = "geekbyteonline.onmicrosoft.com"
app_id = "73efa35d-6188-42d4-b258-838a977eb149"
scope_sharepoint = "https://geekbyteonline.sharepoint.com/.default"
sharepoint_base_url = "https://geekbyteonline.sharepoint.com"

# Certificate paths
CERTIFICATE_PATH = "certificate.pem"
PRIVATE_KEY_PATH = "private_key.pem"
CSV_INPUT_PATH = "sites.csv"
CSV_OUTPUT_PATH = "site_timezones.csv"

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

def get_site_timezone(access_token, site_name):
    """Get timezone information for a SharePoint site"""
    timezone_url = f"{sharepoint_base_url}/sites/{site_name}/_api/web/RegionalSettings/TimeZone"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose"
    }
    
    try:
        response = requests.get(timezone_url, headers=headers)
        response.raise_for_status()
        
        if response.status_code == 200:
            data = response.json()
            if 'd' in data:
                return {
                    'timezone_id': data['d']['Id'],
                    'timezone_description': data['d']['Description']
                }
        return None
    except requests.exceptions.HTTPError as err:
        print(f"Error getting timezone for {site_name}: {err}")
        return None
    except Exception as err:
        print(f"Unexpected error getting timezone for {site_name}: {err}")
        return None

def get_site_creation_date(access_token, site_name):
    """Get creation date for a SharePoint site"""
    web_url = f"{sharepoint_base_url}/sites/{site_name}/_api/web"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose"
    }
    
    try:
        response = requests.get(web_url, headers=headers)
        response.raise_for_status()
        
        if response.status_code == 200:
            data = response.json()
            if 'd' in data and 'Created' in data['d']:
                return data['d']['Created']
        return None
    except requests.exceptions.HTTPError as err:
        print(f"Error getting creation date for {site_name}: {err}")
        return None
    except Exception as err:
        print(f"Unexpected error getting creation date for {site_name}: {err}")
        return None

def read_sites_from_csv(csv_path):
    """Read site names from CSV file"""
    sites = []
    try:
        with open(csv_path, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.DictReader(csvfile)
            for row in reader:
                if 'SiteName' in row and row['SiteName'].strip():
                    sites.append(row['SiteName'].strip())
        return sites
    except FileNotFoundError:
        print(f"Error: CSV file not found at {csv_path}")
        raise
    except Exception as e:
        print(f"Error reading CSV file: {e}")
        raise

def write_results_to_csv(results, csv_path):
    """Write results to CSV file"""
    try:
        with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['SiteName', 'TimezoneID', 'TimezoneDescription', 'CreatedDate', 'Status']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for result in results:
                writer.writerow(result)
        
        print(f"Results written to {csv_path}")
    except Exception as e:
        print(f"Error writing to CSV file: {e}")
        raise

def main():
    try:
        print("Starting SharePoint Site Timezone Retrieval...")
        
        # Load certificate and private key
        print("Loading certificate and private key...")
        certificate, private_key = load_certificate_and_key()
        print("✓ Certificate and private key loaded successfully")
        
        # Generate JWT and get access token
        print("Generating JWT token...")
        sharepoint_jwt = get_jwt_token(certificate, private_key, scope_sharepoint)
        print("✓ JWT token generated")
        
        print("Getting access token...")
        access_token = get_access_token(sharepoint_jwt, scope_sharepoint)
        print("✓ Access token obtained successfully")
        
        # Read sites from CSV
        print(f"Reading sites from {CSV_INPUT_PATH}...")
        sites = read_sites_from_csv(CSV_INPUT_PATH)
        print(f"✓ Found {len(sites)} sites to process")
        
        results = []
        
        # Process each site
        for i, site_name in enumerate(sites, 1):
            print(f"\nProcessing site {i} of {len(sites)}: {site_name}")
            
            # Get timezone information
            timezone_info = get_site_timezone(access_token, site_name)
            
            if timezone_info:
                # Get creation date
                created_date = get_site_creation_date(access_token, site_name)
                
                result = {
                    'SiteName': site_name,
                    'TimezoneID': timezone_info['timezone_id'],
                    'TimezoneDescription': timezone_info['timezone_description'],
                    'CreatedDate': created_date if created_date else 'Not Available',
                    'Status': 'Success'
                }
                
                results.append(result)
                print(f"  ✓ Success: Timezone ID {timezone_info['timezone_id']}")
            else:
                result = {
                    'SiteName': site_name,
                    'TimezoneID': 'N/A',
                    'TimezoneDescription': 'N/A',
                    'CreatedDate': 'N/A',
                    'Status': 'Failed'
                }
                
                results.append(result)
                print("  ✗ Failed to retrieve timezone information")
            
            # Add a small delay to avoid rate limiting
            time.sleep(0.2)
        
        # Write results to CSV
        write_results_to_csv(results, CSV_OUTPUT_PATH)
        
        # Display summary
        print("\n=== Processing Summary ===")
        print(f"Total sites processed: {len(results)}")
        success_count = sum(1 for r in results if r['Status'] == 'Success')
        print(f"Successful: {success_count}")
        print(f"Failed: {len(results) - success_count}")
        
        # Display sample results
        print("\n=== Sample Results ===")
        for result in results[:3]:
            print(f"Site: {result['SiteName']}, Timezone ID: {result['TimezoneID']}, Created: {result['CreatedDate']}")
        
        return results
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

if __name__ == "__main__":
    main()