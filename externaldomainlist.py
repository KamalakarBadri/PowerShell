#!/usr/bin/env python3
"""
SharePoint Site Sharing Domains Report Generator - Parallel Version
Uses existing JWT token and multi-threading for faster processing
"""

import json
import csv
import time
import os
import requests
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
import sys

# ============================================================
# ALL CONFIGURATION PARAMETERS - EDIT THESE
# ============================================================

# JWT Token Configuration
JWT_TOKEN = "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCIsIng1dCI6IjEyMzQ1Njc4OTBhYmNkZWYxMjM0NTY3ODkwYWJjZGVmMTIzNDU2Nzg5MGFiY2RlZiJ9.eyJhdWQiOiJodHRwczovL2xvZ2luLm1pY3Jvc29mdG9ubGluZS5jb20vY29udG9zby5vbm1pY3Jvc29mdC5jb20vb2F1dGgyL3YyLjAvdG9rZW4iLCJleHAiOjE3MDQwMDAwMDAsImlzcyI6IjEyMzQ1Njc4LTkwYWItY2RlZi0xMjM0LTU2Nzg5MGFiY2RlZiIsImp0aSI6IjEyMzQ1Njc4LTkwYWItY2RlZi0xMjM0LTU2Nzg5MGFiY2RlZiIsIm5iZiI6MTcwMzk5OTcwMCwic3ViIjoiMTIzNDU2NzgtOTBhYi1jZGVmLTEyMzQtNTY3ODkwYWJjZGVmIn0.abcdefghijklmnopqrstuvwxyz1234567890abcdefghijklmnopqrstuvwxyz1234567890abcdefghijklmnopqrstuvwxyz1234567890abcdefghijklmnopqrstuvwxyz"

# Azure AD / Microsoft Entra ID Configuration
TENANT_NAME = "yourtenant.onmicrosoft.com"  # Your tenant name
APP_ID = "12345678-90ab-cdef-1234-567890abcdef"  # Your App Registration ID

# Input/Output Configuration
INPUT_CSV_FILE = "sites.csv"  # CSV file containing site IDs
OUTPUT_CSV_FILE = "sites_updated.csv"  # Output CSV with sharing domains added
REPORT_CSV_FILE = "sharing_domains_report.csv"  # Separate report file

# CSV Column Configuration
SITE_ID_COLUMN = "SiteId"  # Column name containing site IDs (auto-detect if None)

# SharePoint Configuration
SHAREPOINT_SCOPE = "https://{tenant}-admin.sharepoint.com/.default"  # Will be auto-generated

# Performance Configuration
MAX_WORKERS = 20  # Maximum number of parallel requests (adjust based on your needs)
REQUEST_TIMEOUT = 30  # Timeout in seconds for API requests
RATE_LIMIT_DELAY = 0.1  # Delay between requests in seconds (to avoid rate limiting)
MAX_RETRIES = 3  # Maximum retries for failed requests

# Progress Display
SHOW_PROGRESS = True  # Show progress bar
PROGRESS_UPDATE_INTERVAL = 1  # Update progress every N seconds

# ============================================================
# TOKEN EXCHANGE FUNCTIONS
# ============================================================

def get_access_token_from_jwt(jwt_token, tenant_name, app_id, scope):
    """
    Exchange JWT for an access token
    """
    url = f"https://login.microsoftonline.com/{tenant_name}/oauth2/v2.0/token"
    
    headers = {
        "Content-Type": "application/x-www-form-urlencoded"
    }
    
    data = {
        "client_id": app_id,
        "client_assertion": jwt_token,
        "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        "scope": scope,
        "grant_type": "client_credentials"
    }
    
    try:
        response = requests.post(url, headers=headers, data=data, timeout=REQUEST_TIMEOUT)
        response.raise_for_status()
        result = response.json()
        return result["access_token"]
    except Exception as err:
        print(f"❌ Error getting access token: {err}")
        return None

# ============================================================
# SHAREPOINT API FUNCTIONS
# ============================================================

def get_site_sharing_domains(access_token, site_id, tenant_prefix, retry_count=0):
    """
    Get sharing allowed domain list for a SharePoint site with retry logic
    """
    url = f"https://{tenant_prefix}-admin.sharepoint.com/_api/spo.tenant/sites('{site_id}')"
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose"
    }
    
    try:
        response = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
        
        if response.status_code == 200:
            data = response.json()
            
            # Extract sharing allowed domain list
            sharing_domains = None
            
            # Try different possible locations of the property
            if 'd' in data:
                site_data = data['d']
            else:
                site_data = data
            
            # Look for sharingalloweddomainlist
            if 'SharingAllowedDomainList' in site_data:
                sharing_domains = site_data['SharingAllowedDomainList']
            elif 'sharingalloweddomainlist' in site_data:
                sharing_domains = site_data['sharingalloweddomainlist']
            
            # Also get additional info if available
            site_info = {
                'site_id': site_id,
                'sharing_allowed_domain_list': sharing_domains if sharing_domains else '',
                'sharing_capability': site_data.get('SharingCapability'),
                'site_url': site_data.get('Url', ''),
                'site_title': site_data.get('Title', ''),
                'status': 'Success'
            }
            
            return site_info
            
        elif response.status_code == 429:  # Rate limiting
            if retry_count < MAX_RETRIES:
                wait_time = (2 ** retry_count) * 2  # Exponential backoff
                time.sleep(wait_time)
                return get_site_sharing_domains(access_token, site_id, tenant_prefix, retry_count + 1)
            else:
                return {
                    'site_id': site_id,
                    'sharing_allowed_domain_list': '',
                    'sharing_capability': None,
                    'site_url': '',
                    'site_title': '',
                    'status': 'Failed - Rate Limited'
                }
        else:
            return {
                'site_id': site_id,
                'sharing_allowed_domain_list': '',
                'sharing_capability': None,
                'site_url': '',
                'site_title': '',
                'status': f'Failed - HTTP {response.status_code}'
            }
            
    except requests.exceptions.Timeout:
        if retry_count < MAX_RETRIES:
            time.sleep(2)
            return get_site_sharing_domains(access_token, site_id, tenant_prefix, retry_count + 1)
        else:
            return {
                'site_id': site_id,
                'sharing_allowed_domain_list': '',
                'sharing_capability': None,
                'site_url': '',
                'site_title': '',
                'status': 'Failed - Timeout'
            }
    except Exception as e:
        return {
            'site_id': site_id,
            'sharing_allowed_domain_list': '',
            'sharing_capability': None,
            'site_url': '',
            'site_title': '',
            'status': f'Failed - {str(e)}'
        }

def get_sharing_capability_text(value):
    """Convert numeric sharing capability to text"""
    mapping = {
        0: "Disabled",
        1: "External User Sharing Only",
        2: "External User and Guest Sharing",
        3: "Existing External User Sharing Only",
        4: "Guest Sharing Only"
    }
    return mapping.get(value, f"Unknown ({value})")

# ============================================================
# PROGRESS TRACKING
# ============================================================

class ProgressTracker:
    def __init__(self, total, show_progress=True):
        self.total = total
        self.completed = 0
        self.success = 0
        self.failed = 0
        self.lock = Lock()
        self.show_progress = show_progress
        self.start_time = time.time()
        self.last_update = 0
        
    def update(self, status='success'):
        with self.lock:
            self.completed += 1
            if status == 'success':
                self.success += 1
            else:
                self.failed += 1
            
            if self.show_progress:
                current_time = time.time()
                if current_time - self.last_update >= PROGRESS_UPDATE_INTERVAL or self.completed == self.total:
                    self.display_progress()
                    self.last_update = current_time
    
    def display_progress(self):
        if self.total == 0:
            return
        
        percentage = (self.completed / self.total) * 100
        elapsed = time.time() - self.start_time
        
        # Create progress bar
        bar_length = 30
        filled = int(bar_length * self.completed / self.total)
        bar = '█' * filled + '░' * (bar_length - filled)
        
        # Estimate time remaining
        if self.completed > 0:
            avg_time = elapsed / self.completed
            remaining = (self.total - self.completed) * avg_time
            eta = f"{int(remaining // 60)}m {int(remaining % 60)}s"
        else:
            eta = "Calculating..."
        
        print(f"\r⏳ Progress: [{bar}] {percentage:.1f}% | "
              f"✅ {self.success} | ❌ {self.failed} | "
              f"⏱️ {int(elapsed // 60)}m {int(elapsed % 60)}s | "
              f"⏳ ETA: {eta}", end='', flush=True)
    
    def get_stats(self):
        return {
            'total': self.total,
            'completed': self.completed,
            'success': self.success,
            'failed': self.failed,
            'elapsed': time.time() - self.start_time
        }

# ============================================================
# CSV HANDLING FUNCTIONS
# ============================================================

def read_site_ids_from_csv(csv_file_path, site_id_column=None):
    """
    Read site IDs from CSV file
    """
    site_ids = []
    rows = []
    fieldnames = []
    delimiter = ','
    
    try:
        with open(csv_file_path, 'r', encoding='utf-8-sig') as file:
            # Detect delimiter
            sample = file.read(1024)
            file.seek(0)
            
            try:
                sniffer = csv.Sniffer()
                delimiter = sniffer.sniff(sample).delimiter
            except:
                delimiter = ','  # Default to comma
            
            reader = csv.DictReader(file, delimiter=delimiter)
            fieldnames = reader.fieldnames
            rows = list(reader)
            
            # If column not specified, try to auto-detect
            if not site_id_column:
                for field in fieldnames:
                    if 'site' in field.lower() or 'id' in field.lower():
                        site_id_column = field
                        break
                else:
                    print(f"❌ Could not auto-detect site ID column. Available columns: {fieldnames}")
                    return None, None, None, None
            
            if site_id_column not in fieldnames:
                print(f"❌ Column '{site_id_column}' not found. Available columns: {fieldnames}")
                return None, None, None, None
            
            # Extract site IDs
            for row in rows:
                site_id = row.get(site_id_column, '').strip()
                if site_id:
                    site_ids.append(site_id)
        
        print(f"✅ Loaded {len(site_ids)} site IDs from {csv_file_path}")
        print(f"   Using column: '{site_id_column}'")
        return site_ids, rows, fieldnames, site_id_column
        
    except Exception as e:
        print(f"❌ Error reading CSV file: {str(e)}")
        return None, None, None, None

def update_csv_with_sharing_domains(input_csv_path, output_csv_path, rows, fieldnames, site_id_column, site_data_map):
    """
    Update CSV file with sharing allowed domain list
    """
    try:
        # Add new columns if they don't exist
        new_columns = [
            'sharing_allowed_domain_list',
            'sharing_capability',
            'sharing_capability_text',
            'site_url'
        ]
        
        for col in new_columns:
            if col not in fieldnames:
                fieldnames = list(fieldnames) + [col]
        
        # Update rows with sharing domain data
        for row in rows:
            site_id = row.get(site_id_column, '').strip()
            if site_id in site_data_map:
                data = site_data_map[site_id]
                row['sharing_allowed_domain_list'] = data.get('sharing_allowed_domain_list', '')
                row['sharing_capability'] = data.get('sharing_capability', '')
                row['sharing_capability_text'] = get_sharing_capability_text(data.get('sharing_capability'))
                row['site_url'] = data.get('site_url', '')
            else:
                row['sharing_allowed_domain_list'] = ''
                row['sharing_capability'] = ''
                row['sharing_capability_text'] = ''
                row['site_url'] = ''
        
        # Write updated CSV
        with open(output_csv_path, 'w', newline='', encoding='utf-8') as file:
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(rows)
        
        return True
        
    except Exception as e:
        print(f"❌ Error updating CSV: {str(e)}")
        return False

def generate_report(report_csv_path, site_data_map, site_ids):
    """
    Generate a separate report CSV with all site data
    """
    try:
        with open(report_csv_path, 'w', newline='', encoding='utf-8') as file:
            fieldnames = [
                'site_id',
                'sharing_allowed_domain_list',
                'sharing_capability',
                'sharing_capability_text',
                'site_url',
                'site_title',
                'status'
            ]
            writer = csv.DictWriter(file, fieldnames=fieldnames)
            writer.writeheader()
            
            for site_id in site_ids:
                data = site_data_map.get(site_id, {})
                writer.writerow({
                    'site_id': site_id,
                    'sharing_allowed_domain_list': data.get('sharing_allowed_domain_list', ''),
                    'sharing_capability': data.get('sharing_capability', ''),
                    'sharing_capability_text': get_sharing_capability_text(data.get('sharing_capability')),
                    'site_url': data.get('site_url', ''),
                    'site_title': data.get('site_title', ''),
                    'status': data.get('status', 'Not Processed')
                })
        
        return True
    except Exception as e:
        print(f"❌ Error generating report: {str(e)}")
        return False

# ============================================================
# VALIDATION FUNCTIONS
# ============================================================

def validate_config():
    """Validate all configuration parameters"""
    errors = []
    
    # Check JWT token
    if not JWT_TOKEN or JWT_TOKEN == "eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCIsIng1dCI6IjEyMzQ1Njc4OTBhYmNkZWYxMjM0NTY3ODkwYWJjZGVmMTIzNDU2Nzg5MGFiY2RlZiJ9.eyJhdWQiOiJodHRwczovL2xvZ2luLm1pY3Jvc29mdG9ubGluZS5jb20vY29udG9zby5vbm1pY3Jvc29mdC5jb20vb2F1dGgyL3YyLjAvdG9rZW4iLCJleHAiOjE3MDQwMDAwMDAsImlzcyI6IjEyMzQ1Njc4LTkwYWItY2RlZi0xMjM0LTU2Nzg5MGFiY2RlZiIsImp0aSI6IjEyMzQ1Njc4LTkwYWItY2RlZi0xMjM0LTU2Nzg5MGFiY2RlZiIsIm5iZiI6MTcwMzk5OTcwMCwic3ViIjoiMTIzNDU2NzgtOTBhYi1jZGVmLTEyMzQtNTY3ODkwYWJjZGVmIn0.abcdefghijklmnopqrstuvwxyz1234567890abcdefghijklmnopqrstuvwxyz1234567890abcdefghijklmnopqrstuvwxyz1234567890abcdefghijklmnopqrstuvwxyz":
        errors.append("Please update JWT_TOKEN with your actual JWT token")
    
    # Check tenant name
    if TENANT_NAME == "yourtenant.onmicrosoft.com":
        errors.append("Please update TENANT_NAME with your actual tenant name")
    
    # Check app ID
    if APP_ID == "12345678-90ab-cdef-1234-567890abcdef":
        errors.append("Please update APP_ID with your actual App Registration ID")
    
    # Check input CSV file
    if not os.path.exists(INPUT_CSV_FILE):
        errors.append(f"Input CSV file not found: {INPUT_CSV_FILE}")
    
    return errors

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    print("\n" + "="*70)
    print("🔐 SHAREPOINT SITE SHARING DOMAINS RETRIEVER - PARALLEL VERSION")
    print("="*70)
    print(f"⚡ Parallel Workers: {MAX_WORKERS}")
    print(f"🔄 Max Retries: {MAX_RETRIES}")
    
    # Validate configuration
    print("\n📋 Validating configuration...")
    errors = validate_config()
    if errors:
        print("❌ Configuration errors found:")
        for error in errors:
            print(f"   - {error}")
        print("\nPlease update the configuration parameters at the top of the script.")
        return
    
    print("✅ Configuration validated successfully")
    
    # Step 1: Get access token using JWT
    tenant_prefix = TENANT_NAME.split('.')[0]
    sharepoint_scope = SHAREPOINT_SCOPE.format(tenant=tenant_prefix)
    
    print(f"\n🔄 Getting access token for SharePoint...")
    print(f"   Tenant: {TENANT_NAME}")
    print(f"   App ID: {APP_ID}")
    print(f"   Scope: {sharepoint_scope}")
    
    access_token = get_access_token_from_jwt(
        JWT_TOKEN,
        TENANT_NAME,
        APP_ID,
        sharepoint_scope
    )
    
    if not access_token:
        print("❌ Failed to get access token")
        return
    
    print("✅ Access token obtained successfully")
    
    # Step 2: Read site IDs from CSV
    print(f"\n📖 Reading site IDs from {INPUT_CSV_FILE}...")
    site_ids, rows, fieldnames, detected_column = read_site_ids_from_csv(
        INPUT_CSV_FILE,
        SITE_ID_COLUMN if SITE_ID_COLUMN != "SiteId" else None
    )
    
    if not site_ids:
        print("❌ No site IDs found in the CSV file")
        return
    
    site_id_column = detected_column or SITE_ID_COLUMN
    total_sites = len(site_ids)
    
    # Step 3: Process sites in parallel
    print(f"\n🚀 Processing {total_sites} sites in parallel...")
    print("-" * 70)
    
    site_data_map = {}
    progress = ProgressTracker(total_sites, SHOW_PROGRESS)
    
    # Create a thread pool
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        # Submit all tasks
        future_to_site = {
            executor.submit(get_site_sharing_domains, access_token, site_id, tenant_prefix): site_id
            for site_id in site_ids
        }
        
        # Process completed tasks
        for future in as_completed(future_to_site):
            site_id = future_to_site[future]
            try:
                result = future.result(timeout=REQUEST_TIMEOUT + 10)
                site_data_map[site_id] = result
                
                # Update progress
                if result['status'] == 'Success':
                    progress.update('success')
                else:
                    progress.update('failed')
                    
            except Exception as e:
                site_data_map[site_id] = {
                    'site_id': site_id,
                    'sharing_allowed_domain_list': '',
                    'sharing_capability': None,
                    'site_url': '',
                    'site_title': '',
                    'status': f'Failed - Exception: {str(e)}'
                }
                progress.update('failed')
    
    # Final progress display
    print("\n")
    
    # Step 4: Summary
    stats = progress.get_stats()
    print("\n" + "="*70)
    print("📊 SUMMARY:")
    print(f"  ✅ Successful: {stats['success']}")
    print(f"  ❌ Failed: {stats['failed']}")
    print(f"  📊 Total: {stats['total']}")
    print(f"  ⏱️ Time taken: {int(stats['elapsed'] // 60)}m {int(stats['elapsed'] % 60)}s")
    
    # Step 5: Generate output files
    print("\n💾 Generating output files...")
    
    # Generate timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Update existing CSV
    output_csv = OUTPUT_CSV_FILE.replace('.csv', f'_{timestamp}.csv') if OUTPUT_CSV_FILE else f'sites_updated_{timestamp}.csv'
    if update_csv_with_sharing_domains(INPUT_CSV_FILE, output_csv, rows, fieldnames, site_id_column, site_data_map):
        print(f"  ✅ Updated CSV saved to: {output_csv}")
    
    # Generate separate report
    report_csv = REPORT_CSV_FILE.replace('.csv', f'_{timestamp}.csv') if REPORT_CSV_FILE else f'sharing_domains_report_{timestamp}.csv'
    if generate_report(report_csv, site_data_map, site_ids):
        print(f"  ✅ Report saved to: {report_csv}")
    
    # Step 6: Print sample results
    print("\n" + "="*70)
    print("📊 SAMPLE RESULTS (First 5 sites):")
    print("="*70)
    
    # Get successful results for sample
    successful_results = [site_data_map.get(sid) for sid in site_ids[:10] 
                          if site_data_map.get(sid, {}).get('status') == 'Success']
    
    for i, data in enumerate(successful_results[:5], 1):
        print(f"\n{i}. Site ID: {data['site_id'][:40]}...")
        sharing_domains = data.get('sharing_allowed_domain_list', '')
        if sharing_domains:
            print(f"   Sharing Allowed Domains: {sharing_domains}")
        else:
            print(f"   Sharing Allowed Domains: (None or not set)")
        print(f"   Sharing Capability: {get_sharing_capability_text(data.get('sharing_capability'))}")
        print(f"   Site URL: {data.get('site_url', 'N/A')}")
        print(f"   Status: {data.get('status', 'Unknown')}")
    
    if len(successful_results) > 5:
        print(f"\n... and {len(successful_results) - 5} more successful results (see output files)")
    
    # Show failed sites if any
    failed_results = [site_data_map.get(sid) for sid in site_ids 
                      if site_data_map.get(sid, {}).get('status') != 'Success']
    if failed_results:
        print(f"\n❌ Failed Sites ({len(failed_results)}):")
        for data in failed_results[:5]:
            print(f"   - {data['site_id'][:40]}... ({data.get('status', 'Unknown')})")
        if len(failed_results) > 5:
            print(f"   ... and {len(failed_results) - 5} more failures")
    
    print("\n" + "="*70)
    print("✅ PROCESS COMPLETE!")
    print("="*70)
    print(f"\n📁 Output files:")
    print(f"  📄 Updated CSV: {output_csv}")
    print(f"  📄 Report CSV: {report_csv}")
    print("\n💡 Tip: You can use these files for further analysis or reporting.")

# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⏹️ Process interrupted by user")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Unexpected error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
