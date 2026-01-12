import requests
import json
import uuid
import base64
import time
import os
from datetime import datetime
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend
import csv

# Configuration - Update these values
CONFIG = {
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    "scope": "https://geekbyteonline.sharepoint.com/.default"
}

# Token cache
TOKEN_CACHE = {
    "token": None,
    "expires": 0
}

def get_token_with_certificate():
    """Get access token using certificate-based authentication"""
    try:
        if not os.path.exists(CONFIG['certificate_path']) or not os.path.exists(CONFIG['private_key_path']):
            raise Exception("Certificate files not found")
            
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
                "scope": CONFIG['scope'],
                "grant_type": "client_credentials"
            }
        )

        if token_response.status_code == 200:
            return token_response.json()["access_token"]
        else:
            print(f"Certificate token request failed: {token_response.text}")
            return None
            
    except Exception as e:
        print(f"Certificate authentication failed: {str(e)}")
        raise

def get_cached_token():
    """Get cached token if it's still valid, otherwise get a new one"""
    cache = TOKEN_CACHE
    
    # If token exists and hasn't expired (with 5 minute buffer)
    if cache["token"] and cache["expires"] > time.time() + 300:
        return cache["token"]
    
    # Get new token
    token = get_token_with_certificate()
    
    if token:
        # Cache the token with expiration (assuming 1 hour lifetime)
        cache["token"] = token
        cache["expires"] = time.time() + 3600
        return token
    
    return None

def get_file_versions(api_url):
    """Get all file versions using SharePoint REST API with GET method"""
    try:
        # Get SharePoint token
        sharepoint_token = get_cached_token()
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Call SharePoint API to get versions
        headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        
        # Make the GET request
        response = requests.get(api_url, headers=headers)
        
        if response.status_code == 200:
            data = response.json()
            versions = data.get("value", [])
            
            # Extract version info
            version_list = []
            for version in versions:
                version_info = {
                    "ID": version.get("ID", 0),
                    "VersionLabel": version.get("VersionLabel", ""),
                    "Created": version.get("Created", ""),
                    "IsCurrentVersion": version.get("IsCurrentVersion", False),
                    "Size": version.get("Size", 0),
                    "Length": version.get("Length", "0"),
                    "CheckInComment": version.get("CheckInComment", ""),
                    "Url": version.get("Url", "")
                }
                version_list.append(version_info)
            
            return {
                "success": True,
                "versions": version_list,
                "total": len(version_list),
                "status_code": response.status_code
            }
        else:
            error_message = response.text
            try:
                error_json = response.json()
                if "error" in error_json:
                    error_message = error_json["error"].get("message", error_message)
            except:
                pass
            
            return {
                "success": False,
                "message": error_message,
                "status_code": response.status_code,
                "versions": [],
                "total": 0
            }
        
    except Exception as e:
        return {
            "success": False,
            "message": f"Exception occurred: {str(e)}",
            "status_code": 0,
            "versions": [],
            "total": 0
        }

def delete_file_version(api_url, version_label):
    """Delete a specific file version using SharePoint REST API with DELETE method"""
    try:
        # Get SharePoint token
        sharepoint_token = get_cached_token()
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Call SharePoint API to delete version
        headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        
        # Make the DELETE request
        response = requests.delete(api_url, headers=headers)
        
        if response.status_code == 200:
            return {
                "success": True,
                "message": f"Version {version_label} deleted successfully",
                "status_code": response.status_code
            }
        else:
            error_message = response.text
            try:
                error_json = response.json()
                if "error" in error_json:
                    error_message = error_json["error"].get("message", error_message)
            except:
                pass
            
            return {
                "success": False,
                "message": error_message,
                "status_code": response.status_code
            }
        
    except Exception as e:
        return {
            "success": False,
            "message": f"Exception occurred: {str(e)}",
            "status_code": 0
        }

def parse_version_range(start_label, end_label):
    """Parse version range and generate all labels between start and end"""
    labels = []
    
    try:
        # Handle float versions (e.g., 1.0, 2.0)
        if '.' in start_label or '.' in end_label:
            # Convert to float
            start_ver = float(start_label) if '.' in start_label else float(f"{start_label}.0")
            end_ver = float(end_label) if '.' in end_label else float(f"{end_label}.0")
            
            # Generate all versions between start and end
            current = start_ver
            while current <= end_ver:
                if current.is_integer():
                    labels.append(f"{int(current)}.0")
                else:
                    labels.append(f"{current:.1f}")
                current += 1.0
        else:
            # Integer labels
            start_num = int(start_label)
            end_num = int(end_label)
            
            for i in range(start_num, end_num + 1):
                labels.append(f"{i}.0")
                
    except ValueError:
        # If parsing fails, return the labels as-is
        labels = [start_label]
        if start_label != end_label:
            labels.append(end_label)
    
    return labels

def process_file_operations(site_name, file_path, start_version, end_version, is_onedrive=False):
    """Process file operations: get versions and delete specified range"""
    
    # Construct API URLs
    if is_onedrive:
        base_url = f"https://geekbyteonline.sharepoint.com/personal/{site_name}"
        versions_api_url = f"{base_url}/_api/web/GetFileByServerRelativeUrl('/personal/{site_name}/{file_path}')/versions"
        delete_base_url = base_url
        delete_relative_path = f"/personal/{site_name}/{file_path}"
    else:
        base_url = f"https://geekbyteonline.sharepoint.com/sites/{site_name}"
        versions_api_url = f"{base_url}/_api/web/GetFileByServerRelativeUrl('/sites/{site_name}/{file_path}')/versions"
        delete_base_url = base_url
        delete_relative_path = f"/sites/{site_name}/{file_path}"
    
    print(f"\nProcessing: {file_path}")
    print(f"Site: {site_name}")
    print(f"Type: {'OneDrive' if is_onedrive else 'SharePoint'}")
    print("-" * 80)
    
    # Step 1: Get all available versions
    print("Step 1: Getting all file versions...", end="", flush=True)
    versions_result = get_file_versions(versions_api_url)
    
    if not versions_result["success"]:
        print(f" ✗ (Error: {versions_result.get('message', 'Unknown error')})")
        return {
            "success": False,
            "message": versions_result.get('message', 'Failed to get versions'),
            "available_versions": [],
            "deletion_results": []
        }
    
    available_versions = versions_result["versions"]
    print(f" ✓ Found {len(available_versions)} versions")
    
    # Display available versions
    if available_versions:
        print("\nAvailable versions:")
        for v in available_versions:
            current_marker = " (Current)" if v["IsCurrentVersion"] else ""
            print(f"  ID: {v['ID']}, Label: {v['VersionLabel']}{current_marker}, Created: {v['Created']}, Size: {v['Size']} bytes")
    
    # Step 2: Parse deletion range
    print(f"\nStep 2: Parsing deletion range {start_version} to {end_version}...")
    labels_to_delete = parse_version_range(start_version, end_version)
    print(f"  Versions to delete: {', '.join(labels_to_delete)}")
    
    # Step 3: Delete versions in range
    print("\nStep 3: Deleting versions...")
    deletion_results = []
    deleted_count = 0
    failed_count = 0
    skipped_count = 0
    
    # Create a map of available version labels for quick lookup
    available_labels = {v["VersionLabel"]: v["ID"] for v in available_versions}
    
    for label in labels_to_delete:
        if label not in available_labels:
            print(f"  Version {label} not found, skipping...")
            result = {
                "success": False,
                "message": f"Version {label} does not exist",
                "status_code": 404,
                "skipped": True,
                "version_label": label,
                "version_id": 0
            }
            skipped_count += 1
        else:
            version_id = available_labels[label]
            delete_api_url = f"{delete_base_url}/_api/web/GetFileByServerRelativeUrl('{delete_relative_path}')/versions/DeleteByLabel('{label}')"
            
            print(f"  Deleting version {label} (ID: {version_id})...", end="", flush=True)
            
            result = delete_file_version(delete_api_url, label)
            result.update({
                "version_label": label,
                "version_id": version_id,
                "skipped": False
            })
            
            if result["success"]:
                print(" ✓")
                deleted_count += 1
            else:
                print(f" ✗ (Error: {result['message'][:50]}...)")
                failed_count += 1
        
        deletion_results.append(result)
        
        # Small delay to avoid rate limiting
        time.sleep(0.5)
    
    return {
        "success": True,
        "site_name": site_name,
        "file_path": file_path,
        "is_onedrive": is_onedrive,
        "available_versions": available_versions,
        "deletion_results": deletion_results,
        "summary": {
            "total_available": len(available_versions),
            "total_requested": len(labels_to_delete),
            "deleted": deleted_count,
            "failed": failed_count,
            "skipped": skipped_count
        }
    }

def save_versions_report(results, filename):
    """Save available versions report to CSV"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'timestamp',
                'site_name',
                'file_path',
                'site_type',
                'version_id',
                'version_label',
                'created_date',
                'is_current',
                'size_bytes',
                'length',
                'checkin_comment',
                'url'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for result in results:
                if "available_versions" in result:
                    for version in result["available_versions"]:
                        writer.writerow({
                            'timestamp': datetime.now().isoformat(),
                            'site_name': result.get('site_name', ''),
                            'file_path': result.get('file_path', ''),
                            'site_type': 'OneDrive' if result.get('is_onedrive') else 'SharePoint',
                            'version_id': version.get('ID', 0),
                            'version_label': version.get('VersionLabel', ''),
                            'created_date': version.get('Created', ''),
                            'is_current': version.get('IsCurrentVersion', False),
                            'size_bytes': version.get('Size', 0),
                            'length': version.get('Length', '0'),
                            'checkin_comment': version.get('CheckInComment', ''),
                            'url': version.get('Url', '')
                        })
        
        print(f"\nVersions report saved to: {filename}")
        
    except Exception as e:
        print(f"Error saving versions CSV: {str(e)}")

def save_deletion_report(results, filename):
    """Save deletion results to CSV"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'timestamp',
                'site_name',
                'file_path',
                'site_type',
                'version_id',
                'version_label',
                'operation',
                'status',
                'status_code',
                'message',
                'skipped'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for result in results:
                if "deletion_results" in result:
                    for deletion in result["deletion_results"]:
                        writer.writerow({
                            'timestamp': datetime.now().isoformat(),
                            'site_name': result.get('site_name', ''),
                            'file_path': result.get('file_path', ''),
                            'site_type': 'OneDrive' if result.get('is_onedrive') else 'SharePoint',
                            'version_id': deletion.get('version_id', 0),
                            'version_label': deletion.get('version_label', ''),
                            'operation': 'DELETE',
                            'status': 'SUCCESS' if deletion.get('success') else 'FAILED',
                            'status_code': deletion.get('status_code', 0),
                            'message': deletion.get('message', '')[:500],
                            'skipped': deletion.get('skipped', False)
                        })
        
        print(f"Deletion report saved to: {filename}")
        
    except Exception as e:
        print(f"Error saving deletion CSV: {str(e)}")

def print_summary(results):
    """Print summary of operations"""
    total_files = len(results)
    total_available = 0
    total_deleted = 0
    total_failed = 0
    total_skipped = 0
    
    for result in results:
        if "summary" in result:
            summary = result["summary"]
            total_available += summary.get("total_available", 0)
            total_deleted += summary.get("deleted", 0)
            total_failed += summary.get("failed", 0)
            total_skipped += summary.get("skipped", 0)
    
    print("\n" + "="*80)
    print("OPERATIONS SUMMARY")
    print("="*80)
    print(f"Total files processed: {total_files}")
    print(f"Total versions available: {total_available}")
    print(f"Total versions deleted: {total_deleted}")
    print(f"Total versions failed: {total_failed}")
    print(f"Total versions skipped: {total_skipped}")
    print("="*80)
    
    # Print per-file summary
    print("\nPer-file Summary:")
    for i, result in enumerate(results, 1):
        if "summary" in result:
            summary = result["summary"]
            print(f"\n  File {i}: {result.get('file_path')}")
            print(f"    Site: {result.get('site_name')} ({'OneDrive' if result.get('is_onedrive') else 'SharePoint'})")
            print(f"    Available versions: {summary.get('total_available', 0)}")
            print(f"    Requested deletions: {summary.get('total_requested', 0)}")
            print(f"    Deleted: {summary.get('deleted', 0)}")
            print(f"    Failed: {summary.get('failed', 0)}")
            print(f"    Skipped: {summary.get('skipped', 0)}")

def main():
    """Main function - CONFIGURE YOUR FILES HERE"""
    try:
        print("="*80)
        print("FILE VERSION MANAGEMENT TOOL")
        print("="*80)
        print("Using Certificate-based Authentication")
        print("Operations: GET versions (POST endpoint) + DELETE versions")
        print("="*80)
        
        # Test authentication first
        print("\nTesting authentication...", end="", flush=True)
        token = get_cached_token()
        if token:
            print(" ✓ Authentication successful")
        else:
            print(" ✗ Authentication failed")
            print("Please check your certificate files and configuration.")
            return
        
        # ============================================
        # CONFIGURE YOUR FILES HERE
        # ============================================
        
        # Define the files to process
        # Format: (site_name, file_path, start_version, end_version, is_onedrive)
        
        files_to_process = [
            # SharePoint Site Example:
            # ("New365", "Shared Documents/Book.xlsx", "1.0", "10.0", False),
            
            # OneDrive Personal Example:
            # ("jdoe_geekbyteonline_onmicrosoft_com", "Documents/PersonalFile.docx", "1.0", "5.0", True),
            
            # Add more files as needed...
        ]
        
        # If no files configured, show examples
        if not files_to_process:
            print("\nNo files configured. Please add files to the 'files_to_process' list.")
            print("\nExamples:")
            print("  SharePoint Site:")
            print('    ("New365", "Shared Documents/Book.xlsx", "1.0", "10.0", False)')
            print("\n  OneDrive Personal:")
            print('    ("jdoe_geekbyteonline_onmicrosoft_com", "Documents/PersonalFile.docx", "1.0", "5.0", True)')
            print("\nNote: For OneDrive, site_name should be the user part from the URL")
            print('      Example: jdoe_geekbyteonline_onmicrosoft_com')
            return
        
        print(f"\nStarting processing of {len(files_to_process)} files...")
        
        all_results = []
        
        for i, (site_name, file_path, start_ver, end_ver, is_onedrive) in enumerate(files_to_process, 1):
            print(f"\n{'='*80}")
            print(f"Processing File {i}/{len(files_to_process)}")
            print(f"{'='*80}")
            
            result = process_file_operations(site_name, file_path, start_ver, end_ver, is_onedrive)
            all_results.append(result)
        
        # Generate report filenames with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Save versions report
        versions_report_file = f"available_versions_report_{timestamp}.csv"
        save_versions_report(all_results, versions_report_file)
        
        # Save deletion report
        deletion_report_file = f"deletion_results_report_{timestamp}.csv"
        save_deletion_report(all_results, deletion_report_file)
        
        # Print summary
        print_summary(all_results)
        
        print(f"\nReports generated:")
        print(f"  1. Available versions: {versions_report_file}")
        print(f"  2. Deletion results: {deletion_report_file}")
        
        print("\nAvailable Versions Report Columns:")
        print("  - timestamp: When the operation occurred")
        print("  - site_name: Site name or user folder")
        print("  - file_path: Relative file path")
        print("  - site_type: SharePoint or OneDrive")
        print("  - version_id: Version ID number")
        print("  - version_label: Version label (e.g., 1.0, 2.0)")
        print("  - created_date: When version was created")
        print("  - is_current: Whether this is the current version")
        print("  - size_bytes: File size in bytes")
        print("  - length: File length")
        print("  - checkin_comment: Check-in comment")
        print("  - url: Version URL")
        
        print("\nDeletion Report Columns:")
        print("  - timestamp: When the deletion occurred")
        print("  - site_name: Site name or user folder")
        print("  - file_path: Relative file path")
        print("  - site_type: SharePoint or OneDrive")
        print("  - version_id: Version ID that was attempted")
        print("  - version_label: Version label attempted")
        print("  - operation: DELETE")
        print("  - status: SUCCESS or FAILED")
        print("  - status_code: HTTP status code")
        print("  - message: Error/success message")
        print("  - skipped: True if version didn't exist")
        
    except Exception as e:
        print(f"\nScript failed: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
