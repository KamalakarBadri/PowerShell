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
scope_graph = "https://graph.microsoft.com/.default"
scope_sharepoint = "https://geekbyteonline.sharepoint.com/.default"

# Certificate paths
CERTIFICATE_PATH = "certificate.pem"
PRIVATE_KEY_PATH = "private_key.pem"
CSV_INPUT_PATH = "sites.csv"
CSV_OUTPUT_PATH = "site_timezone_changes.csv"

# SharePoint admin URLs
sharepoint_admin_url = "https://geekbyteonline-admin.sharepoint.com"
graph_base_url = "https://graph.microsoft.com/v1.0"

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

def get_site_id_from_graph(access_token, site_name):
    """Get site ID from Graph API"""
    graph_url = f"{graph_base_url}/sites/geekbyteonline.sharepoint.com:/sites/{site_name}"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.get(graph_url, headers=headers)
        response.raise_for_status()
        
        if response.status_code == 200:
            data = response.json()
            return data.get('id')  # Returns the site ID
        return None
    except requests.exceptions.HTTPError as err:
        print(f"Error getting site ID for {site_name}: {err}")
        return None
    except Exception as err:
        print(f"Unexpected error getting site ID for {site_name}: {err}")
        return None

def get_current_timezone(sharepoint_access_token, site_id):
    """Get current timezone from SharePoint Admin API"""
    timezone_url = f"{sharepoint_admin_url}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/Sites('{site_id}')"
    
    headers = {
        "Authorization": f"Bearer {sharepoint_access_token}",
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.get(timezone_url, headers=headers)
        response.raise_for_status()
        
        if response.status_code == 200:
            data = response.json()
            # Extract timezone ID from the response
            # The exact path might need adjustment based on actual API response
            timezone_id = data.get('TimeZoneId')
            return timezone_id
        return None
    except requests.exceptions.HTTPError as err:
        print(f"Error getting current timezone: {err}")
        return None
    except Exception as err:
        print(f"Unexpected error getting current timezone: {err}")
        return None

def update_timezone(sharepoint_access_token, site_id, new_timezone_id=13):
    """Update timezone using SharePoint Admin API"""
    update_url = f"{sharepoint_admin_url}/_api/Microsoft.Online.SharePoint.TenantAdministration.Tenant/Sites('{site_id}')"
    
    headers = {
        "Authorization": f"Bearer {sharepoint_access_token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
        "X-HTTP-Method": "MERGE"
    }
    
    body = {
        "TimeZoneId": new_timezone_id
    }
    
    try:
        response = requests.post(update_url, headers=headers, json=body)
        
        # SharePoint Admin API might return 204 No Content for successful updates
        if response.status_code in [200, 204]:
            print(f"✓ Successfully updated timezone to {new_timezone_id}")
            return True
        else:
            print(f"✗ Failed to update timezone. Status: {response.status_code}")
            print(f"Response: {response.text}")
            return False
            
    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error updating timezone: {err}")
        print(f"Response: {response.text if 'response' in locals() else 'N/A'}")
        return False
    except Exception as err:
        print(f"Unexpected error updating timezone: {err}")
        return False

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
            fieldnames = ['SiteName', 'SiteID', 'CurrentTimezone', 'NewTimezone', 'Status', 'Message']
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
        print("Starting SharePoint Site Timezone Update...")
        
        # Load certificate and private key
        print("Loading certificate and private key...")
        certificate, private_key = load_certificate_and_key()
        print("✓ Certificate and private key loaded successfully")
        
        # Get Graph access token for site ID lookup
        print("Getting Graph access token...")
        graph_jwt = get_jwt_token(certificate, private_key, scope_graph)
        graph_token = get_access_token(graph_jwt, scope_graph)
        print("✓ Graph access token obtained")
        
        # Get SharePoint access token for admin operations
        print("Getting SharePoint access token...")
        sharepoint_jwt = get_jwt_token(certificate, private_key, scope_sharepoint)
        sharepoint_token = get_access_token(sharepoint_jwt, scope_sharepoint)
        print("✓ SharePoint access token obtained")
        
        # Read sites from CSV
        print(f"Reading sites from {CSV_INPUT_PATH}...")
        sites = read_sites_from_csv(CSV_INPUT_PATH)
        print(f"✓ Found {len(sites)} sites to process")
        
        results = []
        target_timezone = 13
        
        # Process each site
        for i, site_name in enumerate(sites, 1):
            print(f"\nProcessing site {i} of {len(sites)}: {site_name}")
            
            # Get site ID from Graph API
            site_id = get_site_id_from_graph(graph_token, site_name)
            
            if not site_id:
                result = {
                    'SiteName': site_name,
                    'SiteID': 'N/A',
                    'CurrentTimezone': 'N/A',
                    'NewTimezone': 'N/A',
                    'Status': 'Failed',
                    'Message': 'Could not retrieve site ID'
                }
                results.append(result)
                print("  ✗ Failed to get site ID")
                continue
            
            print(f"  Site ID: {site_id}")
            
            # Get current timezone
            current_timezone = get_current_timezone(sharepoint_token, site_id)
            
            if current_timezone is None:
                result = {
                    'SiteName': site_name,
                    'SiteID': site_id,
                    'CurrentTimezone': 'N/A',
                    'NewTimezone': 'N/A',
                    'Status': 'Failed',
                    'Message': 'Could not retrieve current timezone'
                }
                results.append(result)
                print("  ✗ Failed to get current timezone")
                continue
            
            print(f"  Current timezone: {current_timezone}")
            
            # Check if timezone needs to be changed (from 23 to 13)
            if current_timezone == 23:
                print(f"  Timezone needs update (23 → {target_timezone})")
                
                # Update timezone
                success = update_timezone(sharepoint_token, site_id, target_timezone)
                
                if success:
                    result = {
                        'SiteName': site_name,
                        'SiteID': site_id,
                        'CurrentTimezone': current_timezone,
                        'NewTimezone': target_timezone,
                        'Status': 'Updated',
                        'Message': f'Successfully changed from {current_timezone} to {target_timezone}'
                    }
                    print(f"  ✓ Timezone updated successfully")
                else:
                    result = {
                        'SiteName': site_name,
                        'SiteID': site_id,
                        'CurrentTimezone': current_timezone,
                        'NewTimezone': 'N/A',
                        'Status': 'Failed',
                        'Message': 'Failed to update timezone'
                    }
                    print("  ✗ Failed to update timezone")
            else:
                result = {
                    'SiteName': site_name,
                    'SiteID': site_id,
                    'CurrentTimezone': current_timezone,
                    'NewTimezone': current_timezone,
                    'Status': 'NoChange',
                    'Message': f'Timezone is already {current_timezone} (no update needed)'
                }
                print(f"  ✓ Timezone is already {current_timezone} (no update needed)")
            
            results.append(result)
            
            # Add a small delay to avoid rate limiting
            time.sleep(1)
        
        # Write results to CSV
        write_results_to_csv(results, CSV_OUTPUT_PATH)
        
        # Display summary
        print("\n=== Processing Summary ===")
        print(f"Total sites processed: {len(results)}")
        updated_count = sum(1 for r in results if r['Status'] == 'Updated')
        nochange_count = sum(1 for r in results if r['Status'] == 'NoChange')
        failed_count = sum(1 for r in results if r['Status'] == 'Failed')
        
        print(f"Updated: {updated_count}")
        print(f"No change needed: {nochange_count}")
        print(f"Failed: {failed_count}")
        
        return results
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

if __name__ == "__main__":
    main()