import json
import uuid
import base64
import time
import requests
import csv
import os
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
scope_sharepoint = "https://geekbyteonline.sharepoint.com/.default"

# Certificate paths
CERTIFICATE_PATH = "certificate.pem"
PRIVATE_KEY_PATH = "private_key.pem"

# Input file
CSV_INPUT_FILE = "files_to_process.csv"

# Operation mode: "check" for status check only, "delete" to delete files
OPERATION_MODE = "delete"  # Change to "delete" to enable deletion
# OPERATION_MODE = "delete"  # Change to "delete" to enable deletion

# Token management
access_token = None
token_expiry_time = 0

# Global output file variable
OUTPUT_FILE = None

def get_output_filename():
    """Generate output filename with timestamp"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    if OPERATION_MODE == "check":
        return f"file_check_results_{timestamp}.csv"
    else:
        return f"file_deletion_results_{timestamp}.csv"

def load_certificate_and_key():
    """Load certificate and private key from PEM files"""
    try:
        with open(CERTIFICATE_PATH, "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())

        with open(PRIVATE_KEY_PATH, "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())

        return certificate, private_key
    except Exception as e:
        print(f"Error loading certificate or private key: {str(e)}")
        raise

def get_jwt_token(certificate, private_key, scope):
    """Generate JWT token using certificate and private key"""
    try:
        now = int(time.time())
        expiration = now + 2700  # 45 minutes
        
        thumbprint = certificate.fingerprint(hashes.SHA1())
        x5t = base64.urlsafe_b64encode(thumbprint).decode('utf-8').replace('=', '')
        
        jwt_header = {"alg": "RS256", "typ": "JWT", "x5t": x5t}
        jwt_payload = {
            "aud": f"https://login.microsoftonline.com/{tenant_name}/oauth2/v2.0/token",
            "exp": expiration, "iss": app_id, "jti": str(uuid.uuid4()),
            "nbf": now, "sub": app_id
        }
        
        encoded_header = base64.urlsafe_b64encode(
            json.dumps(jwt_header, separators=(',', ':')).encode('utf-8')
        ).decode('utf-8').replace('=', '')
        
        encoded_payload = base64.urlsafe_b64encode(
            json.dumps(jwt_payload, separators=(',', ':')).encode('utf-8')
        ).decode('utf-8').replace('=', '')
        
        jwt_unsigned = f"{encoded_header}.{encoded_payload}"
        
        signature = private_key.sign(
            jwt_unsigned.encode('utf-8'),
            padding.PKCS1v15(),
            hashes.SHA256()
        )
        encoded_signature = base64.urlsafe_b64encode(signature).decode('utf-8').replace('=', '')
        
        return f"{jwt_unsigned}.{encoded_signature}"

    except Exception as e:
        print(f"Error generating JWT: {str(e)}")
        raise

def get_access_token():
    """Get access token from Microsoft Identity Platform (with caching)"""
    global access_token, token_expiry_time
    
    # Check if token is still valid (with 5 minute buffer)
    if access_token and time.time() < token_expiry_time - 300:
        return access_token
    
    try:
        certificate, private_key = load_certificate_and_key()
        sharepoint_jwt = get_jwt_token(certificate, private_key, scope_sharepoint)
        
        url = f"https://login.microsoftonline.com/{tenant_name}/oauth2/v2.0/token"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        data = {
            "client_id": app_id,
            "client_assertion": sharepoint_jwt,
            "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            "scope": scope_sharepoint,
            "grant_type": "client_credentials"
        }
        
        response = requests.post(url, headers=headers, data=data)
        response.raise_for_status()
        
        token_data = response.json()
        access_token = token_data["access_token"]
        token_expiry_time = time.time() + 2700  # 45 minutes
        
        print("New access token generated (valid for 45 minutes)")
        return access_token
        
    except Exception as err:
        print(f"Error getting access token: {err}")
        raise

def extract_base_url_from_file_url(file_url):
    """Extract the base URL from a full SharePoint file URL"""
    try:
        parsed_url = urlparse(file_url)
        path_parts = parsed_url.path.split('/')
        
        if "sites" in path_parts:
            sites_index = path_parts.index("sites")
            site_name = path_parts[sites_index + 1]
            base_url = f"{parsed_url.scheme}://{parsed_url.netloc}/sites/{site_name}"
            server_relative_path = '/'.join(path_parts[sites_index:])
            return base_url, server_relative_path
        else:
            base_url = f"{parsed_url.scheme}://{parsed_url.netloc}"
            server_relative_path = parsed_url.path
            return base_url, server_relative_path
            
    except Exception as e:
        print(f"Error parsing URL {file_url}: {str(e)}")
        return None, None

def get_next_subsite_api_url(file_url, current_api_url, attempt_count):
    """Build subsite API URL by adding next subsite path"""
    try:
        parsed_url = urlparse(file_url)
        path_parts = parsed_url.path.split('/')
        
        if "sites" in path_parts:
            sites_index = path_parts.index("sites")
            potential_subsite_segments = []
            
            for i in range(sites_index + 2, len(path_parts)):
                segment = path_parts[i]
                if segment not in ['_api', '_vti_bin', '_layouts', 'Shared%20Documents', 'Shared Documents']:
                    potential_subsite_segments.append(segment)
            
            if attempt_count <= len(potential_subsite_segments):
                base_domain = f"{parsed_url.scheme}://{parsed_url.netloc}"
                site_name = path_parts[sites_index + 1]
                
                subsite_path_parts = [site_name]
                for j in range(attempt_count):
                    if j < len(potential_subsite_segments):
                        subsite_path_parts.append(potential_subsite_segments[j])
                
                subsite_base_url = f"{base_domain}/sites/{'/'.join(subsite_path_parts)}"
                old_base_url = current_api_url.split('/_api/')[0]
                new_api_url = current_api_url.replace(old_base_url, subsite_base_url)
                
                print(f"Trying subsite level {attempt_count}")
                return new_api_url
        
        return current_api_url
        
    except Exception as e:
        print(f"Error building subsite URL: {str(e)}")
        return current_api_url

def check_file_exists(access_token, api_url):
    """Check if file exists using SharePoint REST API"""
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json"
        }
        
        response = requests.get(api_url, headers=headers, timeout=30)
        
        if response.status_code == 200:
            return True, response.status_code, "File exists", api_url
        else:
            return False, response.status_code, response.text, api_url
            
    except Exception as e:
        return False, "EXCEPTION", str(e), api_url

def delete_file(access_token, api_url):
    """Delete a file using SharePoint REST API with recycle option"""
    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json",
            "X-HTTP-Method": "DELETE"
        }
        
        response = requests.post(api_url, headers=headers, timeout=30)
        
        if response.status_code in [200, 204]:
            return True, response.status_code, "Success - File moved to recycle bin", api_url
        else:
            return False, response.status_code, response.text, api_url
            
    except Exception as e:
        return False, "EXCEPTION", str(e), api_url

def build_api_url(base_url, server_relative_path, operation):
    """Build API URL for check or delete operation"""
    encoded_path = quote(server_relative_path)
    if operation == "check":
        return f"{base_url}/_api/web/GetFileByServerRelativeUrl('/{encoded_path}')"
    else:
        return f"{base_url}/_api/web/GetFileByServerRelativeUrl('/{encoded_path}')/recycle"

def try_with_subsites(access_token, file_url, original_api_url, operation_func, max_attempts=5):
    """Try the operation with progressively deeper subsite levels"""
    attempt_count = 1
    current_api_url = original_api_url
    
    while attempt_count <= max_attempts:
        if operation_func.__name__ == 'check_file_exists':
            exists, status_code, error_message, api_url_used = operation_func(access_token, current_api_url)
        else:
            success, status_code, error_message, api_url_used = operation_func(access_token, current_api_url)
        
        # If success or file not found, return the result
        if (operation_func.__name__ == 'check_file_exists' and exists) or \
           (operation_func.__name__ == 'delete_file' and success):
            return (exists if operation_func.__name__ == 'check_file_exists' else success), status_code, error_message, api_url_used
        
        # If access denied, try next subsite level
        if "Access denied" in error_message and attempt_count < max_attempts:
            attempt_count += 1
            next_api_url = get_next_subsite_api_url(file_url, current_api_url, attempt_count - 1)
            if next_api_url != current_api_url:
                current_api_url = next_api_url
                continue
        
        # If not access denied or no more subsites to try, return the result
        return (exists if operation_func.__name__ == 'check_file_exists' else success), status_code, error_message, api_url_used
    
    return (exists if operation_func.__name__ == 'check_file_exists' else success), status_code, error_message, api_url_used

def read_csv_file(filename):
    """Read file URLs from CSV file"""
    file_urls = []
    try:
        with open(filename, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if row and row[0].strip():
                    file_urls.append(row[0].strip())
        return file_urls
    except Exception as e:
        print(f"Error reading CSV file: {str(e)}")
        return []

def initialize_output_csv(filename, mode):
    """Initialize the output CSV file with headers"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            if mode == "check":
                writer.writerow([
                    'File URL', 'Status', 'HTTP Status Code', 'Error Message', 'API URL', 'Timestamp'
                ])
            else:
                writer.writerow([
                    'File URL', 'Status', 'HTTP Status Code', 'Error Message', 'API URL', 'Timestamp'
                ])
        print(f"Created new output file: {filename}")
    except Exception as e:
        print(f"Error initializing output CSV: {str(e)}")

def append_result_to_csv(filename, result):
    """Append a single result to the CSV file"""
    try:
        with open(filename, 'a', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(result)
        print(f"✓ Result saved to CSV for: {result[0]}")
    except Exception as e:
        print(f"Error appending to CSV: {str(e)}")

def list_previous_output_files():
    """List previous output files in the current directory"""
    try:
        csv_files = [f for f in os.listdir('.') if f.endswith('.csv') and ('file_check_results_' in f or 'file_deletion_results_' in f)]
        if csv_files:
            print("\nPrevious output files found:")
            for file in sorted(csv_files, reverse=True)[:5]:  # Show only latest 5
                print(f"  - {file}")
        return csv_files
    except Exception as e:
        print(f"Error listing previous files: {str(e)}")
        return []

def main():
    global OUTPUT_FILE
    
    try:
        print(f"Starting file processing in '{OPERATION_MODE}' mode...")
        
        # Show previous output files
        list_previous_output_files()
        
        if OPERATION_MODE == "delete":
            print("\nWARNING: FILES WILL BE MOVED TO RECYCLE BIN")
            confirmation = input("Type 'YES' to confirm you want to delete files: ")
            if confirmation != "YES":
                print("Operation cancelled")
                return
        
        # Generate output filename with timestamp
        OUTPUT_FILE = get_output_filename()
        print(f"\nOutput will be saved to: {OUTPUT_FILE}")
        
        # Get access token (will be cached for 45 minutes)
        access_token = get_access_token()
        
        # Read file URLs from CSV
        file_urls = read_csv_file(CSV_INPUT_FILE)
        if not file_urls:
            print(f"No URLs found in {CSV_INPUT_FILE}")
            return
        
        # Initialize output CSV
        initialize_output_csv(OUTPUT_FILE, OPERATION_MODE)
        
        print(f"Found {len(file_urls)} files to process")
        print("=" * 80)
        
        # Process each file
        for i, file_url in enumerate(file_urls, 1):
            print(f"\nProcessing file {i}/{len(file_urls)}: {file_url}")
            
            # Extract base URL and server relative path
            base_url, server_relative_path = extract_base_url_from_file_url(file_url)
            if not base_url or not server_relative_path:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                result_row = [
                    file_url, "URL_PARSE_ERROR", "N/A", f"Failed to parse URL: {file_url}", "N/A", timestamp
                ]
                append_result_to_csv(OUTPUT_FILE, result_row)
                print(f"✗ Failed to parse URL: {file_url}")
                continue
            
            if OPERATION_MODE == "check":
                # Check-only mode: just verify file existence
                check_api_url = build_api_url(base_url, server_relative_path, "check")
                exists, status_code, error_message, api_url_used = try_with_subsites(
                    access_token, file_url, check_api_url, check_file_exists
                )
                
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                status = "EXISTS" if exists else "NOT_FOUND" if status_code == 404 else "ERROR"
                result_row = [
                    file_url,
                    status,
                    status_code,
                    error_message,
                    api_url_used,
                    timestamp
                ]
                
                append_result_to_csv(OUTPUT_FILE, result_row)
                
                if exists:
                    print(f"✓ File exists: {file_url}")
                elif status_code == 404:
                    print(f"↷ File not found: {file_url}")
                else:
                    print(f"✗ Error checking file: {file_url}")
                    print(f"  → Error: {error_message}")
                    
            else:
                # Delete mode: check if file exists, then delete
                check_api_url = build_api_url(base_url, server_relative_path, "check")
                exists, status_code, error_message, check_api_url_used = try_with_subsites(
                    access_token, file_url, check_api_url, check_file_exists
                )
                
                if not exists:
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    result_row = [
                        file_url, "SKIPPED", status_code, error_message, check_api_url_used, timestamp
                    ]
                    append_result_to_csv(OUTPUT_FILE, result_row)
                    print(f"↷ File not found, skipping: {file_url}")
                    continue
                
                # File exists, proceed with deletion
                delete_api_url = build_api_url(base_url, server_relative_path, "delete")
                success, status_code, error_message, delete_api_url_used = try_with_subsites(
                    access_token, file_url, delete_api_url, delete_file
                )
                
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                result_row = [
                    file_url,
                    "SUCCESS" if success else "FAILED",
                    status_code,
                    error_message,
                    delete_api_url_used,
                    timestamp
                ]
                
                append_result_to_csv(OUTPUT_FILE, result_row)
                
                if success:
                    print(f"✓ Successfully deleted: {file_url}")
                else:
                    print(f"✗ Failed to delete: {file_url}")
                    print(f"  → Error: {error_message}")
            
            # Small delay to avoid rate limiting
            time.sleep(0.5)
        
        # Print final summary
        print("\n" + "=" * 80)
        print("Processing completed!")
        print(f"Results saved to: {OUTPUT_FILE}")
        
        # Show quick summary from output file
        if os.path.exists(OUTPUT_FILE):
            with open(OUTPUT_FILE, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                rows = list(reader)
                if len(rows) > 1:  # More than just header
                    success_count = sum(1 for row in rows[1:] if row[1] in ["EXISTS", "SUCCESS"])
                    skipped_count = sum(1 for row in rows[1:] if row[1] in ["NOT_FOUND", "SKIPPED"])
                    error_count = sum(1 for row in rows[1:] if row[1] in ["ERROR", "FAILED", "URL_PARSE_ERROR"])
                    
                    if OPERATION_MODE == "check":
                        print(f"Files found: {success_count}")
                        print(f"Files not found: {skipped_count}")
                        print(f"Errors: {error_count}")
                    else:
                        print(f"Successfully deleted: {success_count}")
                        print(f"Skipped (not found): {skipped_count}")
                        print(f"Failed: {error_count}")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        if OUTPUT_FILE:
            print(f"Partial results saved to: {OUTPUT_FILE}")

if __name__ == "__main__":
    main()
