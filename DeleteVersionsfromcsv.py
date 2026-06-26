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
import pandas as pd

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

def get_file_name_by_item_id(site_url, list_id, item_id):
    """Get file name using SharePoint REST API with item ID"""
    try:
        # Get SharePoint token
        sharepoint_token = get_cached_token()
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Construct API URL to get file name
        api_url = f"{site_url}/_api/web/lists(guid'{list_id}')/items({item_id})/file?$select=Name"
        
        headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        
        # Make the GET request
        response = requests.get(api_url, headers=headers)
        
        if response.status_code == 200:
            data = response.json()
            file_name = data.get("Name", f"Item_{item_id}")
            return {
                "success": True,
                "file_name": file_name
            }
        else:
            return {
                "success": False,
                "file_name": f"Item_{item_id}",
                "message": f"Failed to get file name: {response.text}"
            }
        
    except Exception as e:
        return {
            "success": False,
            "file_name": f"Item_{item_id}",
            "message": f"Exception occurred: {str(e)}"
        }

def get_file_versions_by_item_id(site_url, list_id, item_id):
    """Get all file versions using SharePoint REST API with item ID"""
    try:
        # Get SharePoint token
        sharepoint_token = get_cached_token()
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Construct API URL using item ID
        api_url = f"{site_url}/_api/web/lists(guid'{list_id}')/items({item_id})/file/versions"
        
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
            
            # Sort versions by label (ascending) - oldest first
            def parse_version(label):
                try:
                    return float(label)
                except:
                    return 0
            
            version_list.sort(key=lambda x: parse_version(x["VersionLabel"]))
            
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

def delete_version_by_item_id(site_url, list_id, item_id, version_label):
    """Delete a specific file version using item ID"""
    try:
        # Get SharePoint token
        sharepoint_token = get_cached_token()
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Construct API URL using item ID
        delete_url = f"{site_url}/_api/web/lists(guid'{list_id}')/items({item_id})/file/versions/deletebylabel('{version_label}')"
        
        headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        
        # Make the POST request (not DELETE)
        response = requests.post(delete_url, headers=headers)
        
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

def process_file_versions(site_url, list_id, item_id, keep_last_n=50):
    """
    Process file: keep LAST N versions (newest) and delete all older versions
    
    Example with 100 versions (1.0 to 100.0), keep_last_n=50:
    - Keep: 51.0 to 100.0 (last 50 versions - newest)
    - Delete: 1.0 to 50.0 (all older versions)
    - Current version (100.0) is automatically kept as it's in the newest 50
    """
    
    # Step 0: Get file name first
    print(f"\n📄 Fetching file information for Item ID: {item_id}...", end="", flush=True)
    file_name_result = get_file_name_by_item_id(site_url, list_id, item_id)
    file_name = file_name_result.get("file_name", f"Item_{item_id}")
    if file_name_result["success"]:
        print(f" ✅ Found: {file_name}")
    else:
        print(f" ⚠️  Using default name: {file_name}")
    
    print(f"\n{'='*80}")
    print(f"📄 Processing File: {file_name}")
    print(f"📍 Item ID: {item_id}")
    print(f"🔗 Site URL: {site_url}")
    print(f"📋 List ID: {list_id}")
    print("-" * 80)
    
    # Step 1: Get all available versions
    print("Step 1: Getting all file versions...", end="", flush=True)
    versions_result = get_file_versions_by_item_id(site_url, list_id, item_id)
    
    if not versions_result["success"]:
        print(f" ✗ (Error: {versions_result.get('message', 'Unknown error')})")
        return {
            "success": False,
            "file_name": file_name,
            "message": versions_result.get('message', 'Failed to get versions'),
            "available_versions": [],
            "deletion_results": [],
            "summary": {
                "total_available": 0,
                "to_delete": 0,
                "deleted": 0,
                "failed": 0,
                "kept": 0
            }
        }
    
    available_versions = versions_result["versions"]
    total_versions = len(available_versions)
    print(f" ✓ Found {total_versions} versions")
    
    # Display available versions
    if available_versions:
        print(f"\n📋 Available versions ({total_versions} total):")
        # Show first 5 and last 5 versions for brevity
        if total_versions <= 10:
            for i, v in enumerate(available_versions):
                current_marker = " (Current)" if v["IsCurrentVersion"] else ""
                print(f"  {i+1}. ID: {v['ID']}, Label: {v['VersionLabel']}{current_marker}")
        else:
            # Show first 5
            for i, v in enumerate(available_versions[:5]):
                current_marker = " (Current)" if v["IsCurrentVersion"] else ""
                print(f"  {i+1}. ID: {v['ID']}, Label: {v['VersionLabel']}{current_marker}")
            print(f"  ... ({total_versions - 10} more versions)")
            # Show last 5
            for i, v in enumerate(available_versions[-5:], total_versions - 4):
                current_marker = " (Current)" if v["IsCurrentVersion"] else ""
                print(f"  {i}. ID: {v['ID']}, Label: {v['VersionLabel']}{current_marker}")
    
    # Step 2: Determine which versions to keep and delete
    print(f"\nStep 2: Calculating versions to keep/delete...")
    print(f"  📌 Keeping: Last {keep_last_n} versions (newest)")
    
    # Sort versions by label (numeric) - oldest first
    def parse_version(label):
        try:
            return float(label)
        except:
            return 0
    
    sorted_versions = sorted(available_versions, key=lambda x: parse_version(x["VersionLabel"]))
    
    # Keep the last N versions (newest)
    if total_versions > keep_last_n:
        versions_to_keep = sorted_versions[-keep_last_n:]  # Last N versions
        versions_to_delete = sorted_versions[:-keep_last_n]  # All versions except last N
    else:
        versions_to_keep = sorted_versions
        versions_to_delete = []
    
    # Display what will be kept and deleted
    print(f"\n  📊 Total versions: {total_versions}")
    print(f"  ✅ Keeping: {len(versions_to_keep)} versions (newest {keep_last_n})")
    print(f"  🗑️  Deleting: {len(versions_to_delete)} versions (oldest)")
    
    if versions_to_delete:
        print(f"\n🗑️  Versions to DELETE (old versions):")
        for v in versions_to_delete[:10]:  # Show first 10
            print(f"  - Label: {v['VersionLabel']}, Created: {v['Created']}")
        if len(versions_to_delete) > 10:
            print(f"  ... and {len(versions_to_delete) - 10} more older versions")
    else:
        print(f"\n✅ No versions to delete (total {total_versions} <= {keep_last_n})")
    
    # Show kept versions (first 10 and last 10)
    if versions_to_keep:
        kept_labels = [v["VersionLabel"] for v in versions_to_keep]
        print(f"\n✅ Versions to KEEP (newest {len(kept_labels)}):")
        if len(kept_labels) <= 10:
            for label in kept_labels:
                is_current = " (Current)" if any(v["IsCurrentVersion"] and v["VersionLabel"] == label for v in versions_to_keep) else ""
                print(f"  - Label: {label}{is_current}")
        else:
            for label in kept_labels[:5]:
                is_current = " (Current)" if any(v["IsCurrentVersion"] and v["VersionLabel"] == label for v in versions_to_keep) else ""
                print(f"  - Label: {label}{is_current}")
            print(f"  ... ({len(kept_labels) - 10} more versions)")
            for label in kept_labels[-5:]:
                is_current = " (Current)" if any(v["IsCurrentVersion"] and v["VersionLabel"] == label for v in versions_to_keep) else ""
                print(f"  - Label: {label}{is_current}")
    
    # Step 3: Delete old versions
    if versions_to_delete:
        print(f"\nStep 3: Deleting {len(versions_to_delete)} old versions...")
        deletion_results = []
        deleted_count = 0
        failed_count = 0
        
        # Delete oldest first
        for version in versions_to_delete:
            label = version["VersionLabel"]
            version_id = version["ID"]
            
            print(f"  🗑️  Deleting version {label} (ID: {version_id})...", end="", flush=True)
            
            result = delete_version_by_item_id(site_url, list_id, item_id, label)
            result.update({
                "version_label": label,
                "version_id": version_id,
                "skipped": False
            })
            
            if result["success"]:
                print(" ✅")
                deleted_count += 1
            else:
                print(f" ❌ (Error: {result['message'][:50]}...)")
                failed_count += 1
            
            deletion_results.append(result)
            
            # Small delay to avoid rate limiting
            time.sleep(0.5)
    else:
        print(f"\nStep 3: No versions to delete.")
        deletion_results = []
        deleted_count = 0
        failed_count = 0
    
    print(f"\n✅ File '{file_name}' processed successfully!")
    print(f"  📊 Summary: {deleted_count} deleted, {failed_count} failed, {len(versions_to_keep)} kept")
    
    return {
        "success": True,
        "file_name": file_name,
        "site_url": site_url,
        "list_id": list_id,
        "item_id": item_id,
        "available_versions": available_versions,
        "deletion_results": deletion_results,
        "summary": {
            "total_available": total_versions,
            "to_delete": len(versions_to_delete),
            "kept": len(versions_to_keep),
            "deleted": deleted_count,
            "failed": failed_count
        }
    }

def read_files_from_csv(csv_file_path):
    """Read file details from CSV"""
    files = []
    
    try:
        # Try to read with pandas first (supports various CSV formats)
        df = pd.read_csv(csv_file_path)
        
        # Check required columns
        required_columns = ['site_url', 'list_id', 'item_id']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            print(f"Error: Missing required columns in CSV: {missing_columns}")
            print(f"Required columns: {required_columns}")
            print(f"Available columns: {df.columns.tolist()}")
            return []
        
        # Optional: keep_last_n column
        keep_last_n = 50  # Default
        if 'keep_last_n' in df.columns:
            keep_last_n = int(df['keep_last_n'].iloc[0])  # Use first value
        
        for _, row in df.iterrows():
            file_info = {
                'site_url': str(row['site_url']).strip(),
                'list_id': str(row['list_id']).strip(),
                'item_id': int(row['item_id']),
                'keep_last_n': int(row.get('keep_last_n', keep_last_n))
            }
            files.append(file_info)
            
    except Exception as e:
        print(f"Error reading CSV file: {str(e)}")
        print("Please ensure CSV has columns: site_url, list_id, item_id")
        print("Optional column: keep_last_n (defaults to 50)")
        return []
    
    return files

def save_versions_report(results, filename):
    """Save available versions report to CSV"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'timestamp',
                'file_name',
                'site_url',
                'list_id',
                'item_id',
                'version_id',
                'version_label',
                'created_date',
                'is_current',
                'size_bytes',
                'length',
                'checkin_comment',
                'url',
                'kept_or_deleted'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for result in results:
                if "available_versions" in result:
                    # Determine which versions were kept or deleted
                    deleted_labels = set()
                    
                    if "deletion_results" in result:
                        for deletion in result["deletion_results"]:
                            if deletion.get("success", False):
                                deleted_labels.add(deletion.get("version_label", ""))
                    
                    for version in result["available_versions"]:
                        label = version.get('VersionLabel', '')
                        status = "DELETED" if label in deleted_labels else "KEPT"
                        
                        writer.writerow({
                            'timestamp': datetime.now().isoformat(),
                            'file_name': result.get('file_name', 'Unknown'),
                            'site_url': result.get('site_url', ''),
                            'list_id': result.get('list_id', ''),
                            'item_id': result.get('item_id', ''),
                            'version_id': version.get('ID', 0),
                            'version_label': label,
                            'created_date': version.get('Created', ''),
                            'is_current': version.get('IsCurrentVersion', False),
                            'size_bytes': version.get('Size', 0),
                            'length': version.get('Length', '0'),
                            'checkin_comment': version.get('CheckInComment', ''),
                            'url': version.get('Url', ''),
                            'kept_or_deleted': status
                        })
        
        print(f"\n📊 Versions report saved to: {filename}")
        
    except Exception as e:
        print(f"Error saving versions CSV: {str(e)}")

def save_summary_report(results, filename):
    """Save summary report to CSV"""
    try:
        with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'timestamp',
                'file_name',
                'site_url',
                'list_id',
                'item_id',
                'total_versions',
                'versions_kept',
                'versions_deleted',
                'deleted_success',
                'deleted_failed'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            
            for result in results:
                if "summary" in result:
                    summary = result["summary"]
                    writer.writerow({
                        'timestamp': datetime.now().isoformat(),
                        'file_name': result.get('file_name', 'Unknown'),
                        'site_url': result.get('site_url', ''),
                        'list_id': result.get('list_id', ''),
                        'item_id': result.get('item_id', ''),
                        'total_versions': summary.get('total_available', 0),
                        'versions_kept': summary.get('kept', 0),
                        'versions_deleted': summary.get('to_delete', 0),
                        'deleted_success': summary.get('deleted', 0),
                        'deleted_failed': summary.get('failed', 0)
                    })
        
        print(f"📊 Summary report saved to: {filename}")
        
    except Exception as e:
        print(f"Error saving summary CSV: {str(e)}")

def print_summary(results):
    """Print summary of operations"""
    total_files = len(results)
    total_available = 0
    total_kept = 0
    total_deleted = 0
    total_failed = 0
    
    for result in results:
        if "summary" in result:
            summary = result["summary"]
            total_available += summary.get("total_available", 0)
            total_kept += summary.get("kept", 0)
            total_deleted += summary.get("deleted", 0)
            total_failed += summary.get("failed", 0)
    
    print("\n" + "="*80)
    print("📊 OPERATIONS SUMMARY")
    print("="*80)
    print(f"📄 Total files processed: {total_files}")
    print(f"📋 Total versions available: {total_available}")
    print(f"✅ Total versions kept (newest): {total_kept}")
    print(f"🗑️  Total versions deleted successfully: {total_deleted}")
    print(f"❌ Total versions failed to delete: {total_failed}")
    print("="*80)
    
    # Print per-file summary
    print("\n📄 Per-file Summary:")
    for i, result in enumerate(results, 1):
        if "summary" in result:
            summary = result["summary"]
            print(f"\n  {i}. File: {result.get('file_name', 'Unknown')}")
            print(f"     📍 Item ID: {result.get('item_id', 'N/A')}")
            print(f"     📋 Total versions: {summary.get('total_available', 0)}")
            print(f"     ✅ Kept (newest): {summary.get('kept', 0)}")
            print(f"     🗑️  Deleted: {summary.get('deleted', 0)}")
            print(f"     ❌ Failed: {summary.get('failed', 0)}")

def main():
    """Main function - Read files from CSV"""
    try:
        print("="*80)
        print("📁 FILE VERSION MANAGEMENT TOOL - KEEP LAST 50 VERSIONS")
        print("="*80)
        print("🔐 Using Certificate-based Authentication")
        print("📋 Operations: GET versions + DELETE older versions (keeping newest 50)")
        print("="*80)
        
        # Test authentication first
        print("\n🔐 Testing authentication...", end="", flush=True)
        token = get_cached_token()
        if token:
            print(" ✅ Authentication successful")
        else:
            print(" ❌ Authentication failed")
            print("Please check your certificate files and configuration.")
            return
        
        # ============================================
        # CONFIGURE CSV FILE PATH
        # ============================================
        csv_file_path = "files_to_process.csv"
        
        # Check if CSV file exists
        if not os.path.exists(csv_file_path):
            print(f"\n❌ CSV file '{csv_file_path}' not found.")
            print("\nPlease create a CSV file with the following format:")
            print("\nExample CSV content:")
            print("-" * 60)
            print("site_url,list_id,item_id,keep_last_n")
            print("https://geekbyteonline.sharepoint.com/sites/Team_Site2,e6a3eb36-59ce-44e2-a7b9-77dd61e2b67b,3,50")
            print("https://geekbyteonline.sharepoint.com/sites/Team_Site2,e6a3eb36-59ce-44e2-a7b9-77dd61e2b67b,8,50")
            print("-" * 60)
            print("\n📋 Columns:")
            print("  - site_url: Full SharePoint site URL")
            print("  - list_id: Document library GUID (without 'guid' prefix)")
            print("  - item_id: File item ID number")
            print("  - keep_last_n: (Optional) Number of newest versions to keep (default: 50)")
            print("\nNote: Filename will be automatically fetched from SharePoint")
            return
        
        # Read files from CSV
        print(f"\n📂 Reading file list from: {csv_file_path}")
        files_to_process = read_files_from_csv(csv_file_path)
        
        if not files_to_process:
            print("No files found in CSV or invalid format.")
            return
        
        print(f"\n📄 Found {len(files_to_process)} files to process:")
        for i, file_info in enumerate(files_to_process, 1):
            print(f"  {i}. Item ID: {file_info['item_id']} (Keep: {file_info.get('keep_last_n', 50)} newest versions)")
        
        # Confirm before proceeding
        print("\n" + "="*80)
        response = input("Proceed with version cleanup? (yes/no): ")
        if response.lower() not in ['yes', 'y']:
            print("Operation cancelled.")
            return
        
        print(f"\n🚀 Starting processing of {len(files_to_process)} files...")
        
        all_results = []
        
        for i, file_info in enumerate(files_to_process, 1):
            print(f"\n{'='*80}")
            print(f"📄 Processing File {i}/{len(files_to_process)}")
            print(f"{'='*80}")
            
            result = process_file_versions(
                file_info['site_url'],
                file_info['list_id'],
                file_info['item_id'],
                file_info.get('keep_last_n', 50)
            )
            all_results.append(result)
        
        # Generate report filenames with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Save versions report
        versions_report_file = f"versions_report_{timestamp}.csv"
        save_versions_report(all_results, versions_report_file)
        
        # Save summary report
        summary_report_file = f"summary_report_{timestamp}.csv"
        save_summary_report(all_results, summary_report_file)
        
        # Print summary
        print_summary(all_results)
        
        print(f"\n📊 Reports generated:")
        print(f"  1. 📋 Detailed versions report: {versions_report_file}")
        print(f"  2. 📊 Summary report: {summary_report_file}")
        
    except Exception as e:
        print(f"\n❌ Script failed: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
