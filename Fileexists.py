import json
import uuid
import base64
import time
import requests
import csv
from datetime import datetime
from urllib.parse import urlparse, quote
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

# Input and output files
CSV_INPUT_FILE = "files_to_delete.csv"
OUTPUT_FILE = "deletion_results.csv"

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

def extract_site_info_from_url(file_url):
    """
    Extract site base URL and server relative path from a full SharePoint file URL
    Returns: (base_url, server_relative_path)
    """
    try:
        parsed_url = urlparse(file_url)
        path_parts = parsed_url.path.split('/')
        
        # Find the "sites" part of the path
        if "sites" in path_parts:
            sites_index = path_parts.index("sites")
            site_name = path_parts[sites_index + 1]
            
            # Reconstruct the base URL
            base_url = f"{parsed_url.scheme}://{parsed_url.netloc}/sites/{site_name}"
            
            # Extract the server relative path
            server_relative_path = '/'.join(path_parts[sites_index:])
            
            return base_url, server_relative_path
        else:
            # Handle root site URLs
            base_url = f"{parsed_url.scheme}://{parsed_url.netloc}"
            server_relative_path = parsed_url.path
            return base_url, server_relative_path
            
    except Exception as e:
        print(f"Error parsing URL {file_url}: {str(e)}")
        return None, None

def check_file_exists(access_token, base_url, server_relative_path):
    """
    Check if a file exists using SharePoint REST API
    Returns: (exists, status_code, error_message)
    """
    try:
        # URL encode the server relative path
        encoded_path = quote(server_relative_path)
        
        # Construct the API endpoint
        api_url = f"{base_url}/_api/web/GetFileByServerRelativeUrl('{encoded_path}')"
        
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json"
        }
        
        response = requests.get(api_url, headers=headers)
        
        if response.status_code == 200:
            return True, response.status_code, "File exists"
        else:
            return False, response.status_code, response.text
            
    except Exception as e:
        return False, "EXCEPTION", str(e)

def delete_file(access_token, base_url, server_relative_path):
    """
    Delete a file using SharePoint REST API (permanent deletion)
    Returns: (success, status_code, error_message)
    """
    try:
        # URL encode the server relative path
        encoded_path = quote(server_relative_path)
        
        # Construct the API endpoint for permanent deletion
        api_url = f"{base_url}/_api/web/GetFileByServerRelativeUrl('{encoded_path}')"
        
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json",
            "X-HTTP-Method": "DELETE",
            "If-Match": "*"
        }
        
        response = requests.post(api_url, headers=headers)
        
        if response.status_code in [200, 204]:
            return True, response.status_code, "Success"
        else:
            return False, response.status_code, response.text
            
    except Exception as e:
        return False, "EXCEPTION", str(e)

def read_csv_file(filename):
    """Read file URLs from CSV file"""
    file_urls = []
    try:
        with open(filename, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if row and row[0].strip():  # Skip empty rows
                    file_urls.append(row[0].strip())
        return file_urls
    except Exception as e:
        print(f"Error reading CSV file: {str(e)}")
        return []

def write_results_to_csv(results, filename):
    """Write deletion results to CSV file"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            # Write header
            writer.writerow(['File URL', 'Status', 'Exists', 'Status Code', 'Error Message', 'Timestamp'])
            # Write results
            for result in results:
                writer.writerow(result)
        print(f"Results written to {filename}")
    except Exception as e:
        print(f"Error writing results to CSV: {str(e)}")

def main():
    try:
        print("Starting file deletion process...")
        
        # Load certificate and private key
        certificate, private_key = load_certificate_and_key()
        print("Certificate and private key loaded successfully")
        
        # Get SharePoint token
        sharepoint_jwt = get_jwt_token(certificate, private_key, scope_sharepoint)
        print("Generated SharePoint JWT")
        sharepoint_token = get_access_token(sharepoint_jwt, scope_sharepoint)
        print("SharePoint access token retrieved successfully")
        
        # Read file URLs from CSV
        file_urls = read_csv_file(CSV_INPUT_FILE)
        if not file_urls:
            print(f"No URLs found in {CSV_INPUT_FILE}")
            return
        
        print(f"Found {len(file_urls)} files to process")
        
        # Process each file
        results = []
        for i, file_url in enumerate(file_urls, 1):
            print(f"Processing file {i}/{len(file_urls)}: {file_url}")
            
            # Extract site info from the URL
            base_url, server_relative_path = extract_site_info_from_url(file_url)
            if not base_url or not server_relative_path:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                results.append([file_url, "FAILED", "UNKNOWN", "URL_PARSE_ERROR", f"Failed to parse URL: {file_url}", timestamp])
                print(f"✗ Failed to parse URL: {file_url}")
                continue
            
            # First check if the file exists
            exists, status_code, error_message = check_file_exists(sharepoint_token, base_url, server_relative_path)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            if not exists:
                # File doesn't exist
                results.append([file_url, "SKIPPED", "NO", status_code, f"File does not exist: {error_message}", timestamp])
                print(f"↷ File does not exist: {file_url}")
                continue
            
            # File exists, proceed with deletion
            success, status_code, error_message = delete_file(sharepoint_token, base_url, server_relative_path)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            if success:
                print(f"✓ Successfully deleted: {file_url}")
                results.append([file_url, "SUCCESS", "YES", status_code, "", timestamp])
            else:
                print(f"✗ Failed to delete: {file_url} - {error_message}")
                results.append([file_url, "FAILED", "YES", status_code, error_message, timestamp])
            
            # Small delay to avoid rate limiting
            time.sleep(0.5)
        
        # Write results to output CSV
        write_results_to_csv(results, OUTPUT_FILE)
        
        # Print summary
        success_count = sum(1 for r in results if r[1] == "SUCCESS")
        skipped_count = sum(1 for r in results if r[1] == "SKIPPED")
        failed_count = sum(1 for r in results if r[1] == "FAILED")
        
        print(f"\nProcessing completed!")
        print(f"Successfully deleted: {success_count}")
        print(f"Skipped (not found): {skipped_count}")
        print(f"Failed: {failed_count}")
        print(f"Results saved to: {OUTPUT_FILE}")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
