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

def get_subsites(sharepoint_token):
    """Get all subsites from the specified SharePoint location"""
    headers = {
        "Authorization": f"Bearer {sharepoint_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json"
    }
    
    try:
        print(f"Fetching subsites from: {BASE_SUBSITES_URL}")
        response = requests.get(BASE_SUBSITES_URL, headers=headers)
        response.raise_for_status()
        
        data = response.json()
        subsites = []
        
        if 'd' in data and 'results' in data['d']:
            for subsite in data['d']['results']:
                title = subsite.get('Title', 'Unknown')
                url = subsite.get('Url', '')
                subsites.append({
                    'Title': title,
                    'Url': url
                })
                print(f"Found subsite: {title} - {url}")
        
        print(f"Total subsites found: {len(subsites)}")
        return subsites
        
    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error getting subsites: {err}")
        print(f"Response: {response.text}")
        raise
    except Exception as err:
        print(f"Error getting subsites: {err}")
        raise

def get_site_policy(site_url, sharepoint_token):
    """Get policy name for a specific site"""
    # Construct the API endpoint for getting site properties
    api_url = f"{site_url}/_api/web/allproperties"
    
    headers = {
        "Authorization": f"Bearer {sharepoint_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json"
    }
    
    try:
        print(f"Getting policy for: {site_url}")
        response = requests.get(api_url, headers=headers)
        response.raise_for_status()
        
        data = response.json()
        policy_name = "No Policy Found"
        
        # Look for policy name in different possible locations
        if 'd' in data:
            properties = data['d']
            
            # Check common property names for policy
            policy_fields = ['PolicyName', 'policyname', 'SitePolicy', 'sitepolicy', 'Policy', 'policy']
            
            for field in policy_fields:
                if field in properties and properties[field]:
                    policy_name = properties[field]
                    break
            
            # If not found in direct properties, check in custom properties
            if policy_name == "No Policy Found" and 'AllProperties' in properties:
                all_props = properties['AllProperties']
                for field in policy_fields:
                    if field in all_props and all_props[field]:
                        policy_name = all_props[field]
                        break
        
        print(f"Policy found: {policy_name}")
        return policy_name
        
    except requests.exceptions.HTTPError as err:
        print(f"HTTP Error getting policy for {site_url}: {err}")
        if response.status_code == 404:
            return "Site Not Found"
        elif response.status_code == 403:
            return "Access Denied"
        else:
            return f"Error: {response.status_code}"
    except Exception as err:
        print(f"Error getting policy for {site_url}: {err}")
        return "Error Retrieving Policy"

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
            writer.writerow(['Site Title', 'Site URL', 'Policy Name'])
            for site in subsites_data:
                writer.writerow([site['Title'], site['Url'], site['PolicyName']])
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
        
        print(f"\nStep 3: Getting policy information for {len(subsites)} subsites...")
        # Get policy for each subsite
        subsites_with_policy = []
        
        for i, subsite in enumerate(subsites, 1):
            print(f"\nProcessing {i}/{len(subsites)}: {subsite['Title']}")
            policy_name = get_site_policy(subsite['Url'], sharepoint_token)
            
            subsites_with_policy.append({
                'Title': subsite['Title'],
                'Url': subsite['Url'],
                'PolicyName': policy_name
            })
            
            # Add a small delay to avoid throttling
            time.sleep(0.5)
        
        print("\nStep 4: Generating reports...")
        # Generate reports
        generate_report(subsites_with_policy, 'both')
        
        print("\n=== Summary ===")
        print(f"Total subsites processed: {len(subsites_with_policy)}")
        
        # Summary statistics
        policy_counts = {}
        for site in subsites_with_policy:
            policy = site['PolicyName']
            policy_counts[policy] = policy_counts.get(policy, 0) + 1
        
        print("\nPolicy distribution:")
        for policy, count in policy_counts.items():
            print(f"  {policy}: {count} sites")
        
        print("\nReport generation completed successfully!")
        return subsites_with_policy
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None

if __name__ == "__main__":
    main()
