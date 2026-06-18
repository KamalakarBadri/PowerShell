import requests
import json
import os
import csv
import datetime
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Optional, Set

# ==================== CONFIGURATION ====================
TENANT_ID = "your-tenant-id"
CLIENT_ID = "your-client-id"
CLIENT_SECRET = "your-client-secret"

# File names
CSV_REPORT_FILE = "azure_ad_users_report.csv"
BACKUP_FILE = "azure_ad_users_report_backup.csv"

# Performance settings
MAX_WORKERS = 10
BATCH_SIZE = 100

# Microsoft Graph API endpoints
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
# =======================================================

# ==================== AUTHENTICATION ====================
def get_access_token():
    """Get OAuth2 access token using client credentials"""
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    
    data = {
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    
    response = requests.post(token_url, data=data)
    response.raise_for_status()
    return response.json()["access_token"]

# ==================== GRAPH API CALLS ====================
def make_graph_request(headers, url, params=None):
    """Make a Graph API request with error handling"""
    try:
        response = requests.get(url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.RequestException as e:
        return None

def get_all_users_filtered(headers):
    """
    Get all users excluding guests using BOTH $filter and code filtering
    This provides double protection against guests
    """
    all_users = []
    url = f"{GRAPH_API_BASE}/users"
    params = {
        "$select": "id,displayName,userPrincipalName,mail,createdDateTime,userType,accountEnabled,department,jobTitle",
        "$filter": "userType eq 'Member'",  # Primary filter - exclude guests at API level
        "$top": 100
    }
    
    page_count = 0
    guest_count = 0
    
    while url:
        page_count += 1
        print(f"  Fetching page {page_count}...")
        data = make_graph_request(headers, url, params)
        if not data:
            break
            
        if "value" in data:
            # Secondary filter in code (fallback)
            for user in data["value"]:
                user_type = user.get("userType", "")
                
                # Double-check: exclude guests explicitly
                if user_type.lower() != "guest":
                    all_users.append(user)
                else:
                    guest_count += 1
                    print(f"    Skipping guest user: {user.get('userPrincipalName', 'Unknown')}")
        
        url = data.get("@odata.nextLink")
        params = None
    
    if guest_count > 0:
        print(f"  Filtered out {guest_count} guest users")
    
    return all_users

def get_user_manager_batch(headers, user_ids):
    """Get managers for multiple users in batch"""
    managers = {}
    for user_id in user_ids:
        url = f"{GRAPH_API_BASE}/users/{user_id}/manager"
        data = make_graph_request(headers, url)
        
        if data and "id" in data:
            manager_url = f"{GRAPH_API_BASE}/users/{data['id']}"
            manager_data = make_graph_request(headers, manager_url)
            if manager_data:
                managers[user_id] = {
                    "id": manager_data.get("id"),
                    "displayName": manager_data.get("displayName"),
                    "mail": manager_data.get("mail"),
                    "userPrincipalName": manager_data.get("userPrincipalName")
                }
    return managers

def get_onedrive_batch(headers, user_ids):
    """Get OneDrive URLs for multiple users in batch"""
    onedrives = {}
    for user_id in user_ids:
        url = f"{GRAPH_API_BASE}/users/{user_id}/drive"
        data = make_graph_request(headers, url)
        
        if data and "webUrl" in data:
            onedrives[user_id] = data.get("webUrl")
    return onedrives

# ==================== PROCESSING ====================
def process_user_batch(user_batch, headers, existing_data):
    """Process a batch of users in parallel"""
    user_ids = [user["id"] for user in user_batch]
    
    # Fetch managers and OneDrive in parallel
    with ThreadPoolExecutor(max_workers=2) as executor:
        manager_future = executor.submit(get_user_manager_batch, headers, user_ids)
        onedrive_future = executor.submit(get_onedrive_batch, headers, user_ids)
        
        managers = manager_future.result()
        onedrives = onedrive_future.result()
    
    # Process each user in the batch
    results = []
    for user in user_batch:
        user_id = user["id"]
        user_type = user.get("userType", "")
        
        # Skip guests just in case (extra safety)
        if user_type.lower() == "guest":
            continue
        
        # Get manager info
        manager = managers.get(user_id)
        manager_name = manager.get("displayName") if manager else None
        manager_email = manager.get("mail") or manager.get("userPrincipalName") if manager else None
        
        # Get OneDrive URL
        onedrive_url = onedrives.get(user_id)
        
        # Determine status
        is_active = user.get("accountEnabled", True)
        status = "active" if is_active else "disabled"
        
        # Check if user exists in existing CSV
        existing_user = next((u for u in existing_data if u.get('id') == user_id), None)
        
        # Determine change type
        change_type = "NO_CHANGE"
        if existing_user is None:
            change_type = "NEW_USER"
        elif existing_user.get('status') == 'deleted' and status == 'active':
            change_type = "REACTIVATED"
        elif existing_user.get('status') != 'deleted':
            # Check for changes
            changes = []
            for field in ['displayName', 'userPrincipalName', 'mail', 'department', 'jobTitle']:
                if existing_user.get(field, '') != user.get(field, ''):
                    changes.append(field)
            
            if existing_user.get('manager', '') != (manager_name or ''):
                changes.append('manager')
            if existing_user.get('managerEmail', '') != (manager_email or ''):
                changes.append('managerEmail')
            if existing_user.get('onedrive_url', '') != (onedrive_url or ''):
                changes.append('onedrive_url')
            
            if changes:
                change_type = "UPDATED"
        
        # Build user data
        user_data = {
            "id": user_id,
            "displayName": user.get("displayName", ""),
            "userPrincipalName": user.get("userPrincipalName", ""),
            "mail": user.get("mail", ""),
            "createdDateTime": user.get("createdDateTime", ""),
            "department": user.get("department", ""),
            "jobTitle": user.get("jobTitle", ""),
            "manager": manager_name if manager_name else "",
            "managerEmail": manager_email if manager_email else "",
            "onedrive_url": onedrive_url if onedrive_url else "",
            "status": status,
            "lastUpdated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "changeType": change_type,
            "userType": user_type  # Added for visibility
        }
        
        results.append(user_data)
    
    return results

# ==================== CSV OPERATIONS ====================
def load_existing_csv():
    """Load existing CSV report if it exists"""
    if os.path.exists(CSV_REPORT_FILE):
        try:
            with open(CSV_REPORT_FILE, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                return list(reader)
        except Exception as e:
            print(f"Warning: Could not load existing CSV: {e}")
            return []
    return []

def backup_existing_csv():
    """Create a backup of the existing CSV"""
    if os.path.exists(CSV_REPORT_FILE):
        try:
            import shutil
            shutil.copy2(CSV_REPORT_FILE, BACKUP_FILE)
            print(f"Backup created: {BACKUP_FILE}")
        except Exception as e:
            print(f"Warning: Could not create backup: {e}")

def save_csv_report(users_data):
    """Save the report to CSV file"""
    if not users_data:
        print("No data to save")
        return
    
    fieldnames = [
        'id', 'displayName', 'userPrincipalName', 'mail', 'createdDateTime',
        'department', 'jobTitle', 'manager', 'managerEmail', 'onedrive_url',
        'status', 'lastUpdated', 'changeType', 'userType'
    ]
    
    try:
        backup_existing_csv()
        
        with open(CSV_REPORT_FILE, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
            writer.writeheader()
            
            for user in users_data:
                row = {field: user.get(field, '') for field in fieldnames}
                writer.writerow(row)
        
        print(f"\n✓ CSV report saved: {CSV_REPORT_FILE}")
        print(f"  Total rows: {len(users_data)}")
        
    except Exception as e:
        print(f"Error saving CSV: {e}")

# ==================== CORE LOGIC ====================
def identify_deleted_users(current_user_ids, existing_data):
    """Identify users that are in the CSV but not in Azure AD"""
    deleted_users = []
    for user in existing_data:
        user_id = user.get('id')
        if user_id and user_id not in current_user_ids:
            if user.get('status') != 'deleted':
                deleted_users.append(user)
    return deleted_users

def process_users_parallel(current_users, existing_data, headers):
    """Process all users in parallel batches"""
    results = []
    current_user_ids = set()
    new_count = 0
    updated_count = 0
    
    # Create batches
    batches = []
    for i in range(0, len(current_users), BATCH_SIZE):
        batch = current_users[i:i + BATCH_SIZE]
        batches.append(batch)
    
    print(f"\nProcessing {len(current_users)} users in {len(batches)} batches...")
    
    # Process batches in parallel
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = []
        for batch in batches:
            future = executor.submit(process_user_batch, batch, headers, existing_data)
            futures.append(future)
        
        # Collect results
        completed = 0
        for future in as_completed(futures):
            try:
                batch_results = future.result()
                results.extend(batch_results)
                completed += 1
                print(f"  Completed batch {completed}/{len(batches)}")
            except Exception as e:
                print(f"  Error processing batch: {e}")
    
    # Update counts and collect user IDs
    for user in results:
        current_user_ids.add(user["id"])
        if user["changeType"] == "NEW_USER":
            new_count += 1
        elif user["changeType"] == "UPDATED":
            updated_count += 1
    
    # Handle deleted users
    print("\nChecking for deleted users...")
    deleted_users = identify_deleted_users(current_user_ids, existing_data)
    
    for user in deleted_users:
        user["status"] = "deleted"
        user["lastUpdated"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        user["changeType"] = "DELETED"
        results.append(user)
        print(f"  DELETED: {user.get('userPrincipalName', 'Unknown')}")
    
    # Print summary
    print("\n" + "="*50)
    print("PROCESSING SUMMARY")
    print("="*50)
    print(f"Total Member users in Azure AD: {len(current_users)}")
    print(f"Total users in report: {len(results)}")
    print(f"New users added: {new_count}")
    print(f"Users updated: {updated_count}")
    print(f"Users deleted/marked: {len(deleted_users)}")
    print("="*50)
    
    return results

# ==================== MAIN FUNCTION ====================
def main():
    """Main function - run the full process"""
    print("="*60)
    print("AZURE AD USER REPORT GENERATOR")
    print("(Excludes Guest Users - Double Filtering Applied)")
    print("="*60)
    print(f"Started at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
    start_time = time.time()
    
    # Get access token
    print("1. Authenticating...")
    try:
        global headers
        access_token = get_access_token()
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        print("   ✓ Authentication successful")
    except Exception as e:
        print(f"   ✗ Authentication failed: {e}")
        return
    
    # Load existing CSV
    print("\n2. Loading existing CSV report...")
    existing_data = load_existing_csv()
    print(f"   Found {len(existing_data)} users in existing report")
    
    # Fetch users from Azure AD (excluding guests)
    print("\n3. Fetching users from Azure AD...")
    print("   Filtering: userType eq 'Member'")
    print("   Additional code-level filtering for guests")
    try:
        fetch_start = time.time()
        current_users = get_all_users_filtered(headers)
        fetch_time = time.time() - fetch_start
        print(f"   Found {len(current_users)} Member users in Azure AD")
        print(f"   Fetch time: {fetch_time:.2f} seconds")
    except Exception as e:
        print(f"   ✗ Error fetching users: {e}")
        return
    
    # Process users in parallel
    print("\n4. Processing user data in parallel...")
    process_start = time.time()
    processed_data = process_users_parallel(current_users, existing_data, headers)
    process_time = time.time() - process_start
    print(f"   Processing time: {process_time:.2f} seconds")
    
    # Save CSV report
    print("\n5. Saving CSV report...")
    save_csv_report(processed_data)
    
    total_time = time.time() - start_time
    
    print("\n" + "="*60)
    print("COMPLETED SUCCESSFULLY")
    print("="*60)
    print(f"Report saved to: {CSV_REPORT_FILE}")
    print(f"Total execution time: {total_time:.2f} seconds")
    print(f"Completed at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\nScript interrupted by user")
    except Exception as e:
        print(f"\n\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
