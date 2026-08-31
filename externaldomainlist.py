#!/usr/bin/env python3
"""
SharePoint Site Sharing Domains Report Generator
Uses existing JWT token to get sharing allowed domain lists for SharePoint sites
"""

import json
import csv
import time
import os
import requests
from datetime import datetime

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

# Request Configuration
REQUEST_TIMEOUT = 30  # Timeout in seconds for API requests
DELAY_BETWEEN_REQUESTS = 0.5  # Delay in seconds between requests to avoid rate limiting

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

def get_site_sharing_domains(access_token, site_id, tenant_prefix):
    """
    Get sharing allowed domain list for a SharePoint site
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
                'sharing_allowed_domain_list': sharing_domains if sharing_domains else '',
                'sharing_capability': site_data.get('SharingCapability'),
                'site_url': site_data.get('Url', ''),
                'site_title': site_data.get('Title', ''),
                'status': 'Success'
            }
            
            return site_info
        else:
            return {
                'sharing_allowed_domain_list': '',
                'sharing_capability': None,
                'site_url': '',
                'site_title': '',
                'status': f'Failed - HTTP {response.status_code}'
            }
            
    except requests.exceptions.Timeout:
        return {
            'sharing_allowed_domain_list': '',
            'sharing_capability': None,
            'site_url': '',
            'site_title': '',
            'status': 'Failed - Timeout'
        }
    except Exception as e:
        return {
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
    print("🔐 SHAREPOINT SITE SHARING DOMAINS RETRIEVER")
    print("="*70)
    
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
    print(f"   JWT Token: {JWT_TOKEN[:50]}...")
    
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
    
    # Step 3: Get sharing domains for each site
    print(f"\n🔍 Fetching sharing allowed domain lists for {len(site_ids)} sites...")
    print("-" * 70)
    
    site_data_map = {}
    success_count = 0
    
    for idx, site_id in enumerate(site_ids, 1):
        print(f"  [{idx}/{len(site_ids)}] Processing site: {site_id[:30]}...")
        
        site_data = get_site_sharing_domains(access_token, site_id, tenant_prefix)
        site_data_map[site_id] = site_data
        
        if site_data['status'] == 'Success':
            success_count += 1
            sharing_domains = site_data.get('sharing_allowed_domain_list', '')
            if sharing_domains:
                print(f"    ✅ Domains: {sharing_domains[:50]}")
            else:
                print(f"    ℹ️ No domain restrictions")
        else:
            print(f"    ❌ {site_data['status']}")
        
        # Add delay between requests
        if idx < len(site_ids):
            time.sleep(DELAY_BETWEEN_REQUESTS)
    
    # Step 4: Summary
    print("\n" + "="*70)
    print("📊 SUMMARY:")
    print(f"  ✅ Successful: {success_count}")
    print(f"  ❌ Failed: {len(site_ids) - success_count}")
    print(f"  📊 Total: {len(site_ids)}")
    
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
    
    for i, site_id in enumerate(site_ids[:5], 1):
        data = site_data_map.get(site_id, {})
        print(f"\n{i}. Site ID: {site_id[:40]}...")
        sharing_domains = data.get('sharing_allowed_domain_list', '')
        if sharing_domains:
            print(f"   Sharing Allowed Domains: {sharing_domains}")
        else:
            print(f"   Sharing Allowed Domains: (None or not set)")
        print(f"   Sharing Capability: {get_sharing_capability_text(data.get('sharing_capability'))}")
        print(f"   Site URL: {data.get('site_url', 'N/A')}")
        print(f"   Status: {data.get('status', 'Unknown')}")
    
    if len(site_ids) > 5:
        print(f"\n... and {len(site_ids) - 5} more sites (see output files)")
    
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
    except Exception as e:
        print(f"\n❌ Unexpected error: {str(e)}")
