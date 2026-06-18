import requests
import json
import os
import csv
import datetime
from typing import Dict, List, Optional

# ==================== CONFIGURATION ====================
TENANT_ID = "your-tenant-id"
CLIENT_ID = "your-client-id"
CLIENT_SECRET = "your-client-secret"

# File names
CSV_REPORT_FILE = "azure_ad_users_report.csv"
BACKUP_FILE = "azure_ad_users_report_backup.csv"

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
        print(f"API Error: {e}")
        if hasattr(e, 'response') and e.response:
            print(f"Response: {e.response.text}")
        return None

def get_all_users(headers):
    """Get all users with pagination"""
    all_users = []
    url = f"{GRAPH_API_BASE}/users"
    params = {
        "$select": "id,displayName,userPrincipalName,mail,createdDateTime,userType,accountEnabled,department,jobTitle",
        "$top": 100
    }
    
    page_count = 0
    while url:
        page_count += 1
        print(f"  Fetching page {page_count}...")
        data = make_graph_request(headers, url, params)
        if not data:
            break
            
        if "value" in data:
            all_users.extend(data["value"])
        
        url = data.get("@odata.nextLink")
        params = None
    
    return all_users

def get_user_manager(headers, user_id):
    """Get manager details for a user"""
    url = f"{GRAPH_API_BASE}/users/{user_id}/manager"
    data = make_graph_request(headers, url)
    
    if data and "id" in data:
        # Get manager's full details
        manager_url = f"{GRAPH_API_BASE}/users/{data['id']}"
        manager_data = make_graph_request(headers, manager_url)
        if manager_data:
            return {
                "id": manager_data.get("id"),
                "displayName": manager_data.get("displayName"),
                "mail": manager_data.get("mail"),
                "userPrincipalName": manager_data.get("userPrincipalName")
            }
    return None

def get_user_onedrive(headers, user_id):
    """Get OneDrive URL for a user"""
    url = f"{GRAPH_API_BASE}/users/{user_id}/drive"
    data = make_graph_request(headers, url)
    
    if data and "webUrl" in data:
        return data.get("webUrl")
    return None

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
    
    # Define CSV columns
    fieldnames = [
        'id',
        'displayName',
        'userPrincipalName',
        'mail',
        'createdDateTime',
        'department',
        'jobTitle',
        'manager',
        'managerEmail',
        'onedrive_url',
        'status',
        'lastUpdated',
        'changeType'
    ]
    
    try:
        # Backup existing file before writing
        backup_existing_csv()
        
        # Write new CSV
        with open(CSV_REPORT_FILE, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
            writer.writeheader()
            
            for user in users_data:
                # Ensure all fields exist
                row = {field: user.get(field, '') for field in fieldnames}
                writer.writerow(row)
        
        print(f"CSV report saved: {CSV_REPORT_FILE}")
        print(f"Total rows: {len(users_data)}")
        
    except Exception as e:
        print(f"Error saving CSV: {e}")

# ==================== CORE LOGIC ====================
def identify_deleted_users(current_user_ids, existing_data):
    """Identify users that are in the CSV but not in Azure AD"""
    deleted_users = []
    for user in existing_data:
        user_id = user.get('id')
        if user_id and user_id not in current_user_ids:
            # Check if not already marked as deleted
            if user.get('status') != 'deleted':
                deleted_users.append(user)
    return deleted_users

def process_users(current_users, existing_data):
    """Process all users and prepare data for CSV"""
    results = []
    current_user_ids = set()
    new_count = 0
    updated_count = 0
    
    print(f"\nProcessing {len(current_users)} users...")
    
    for idx, user in enumerate(current_users, 1):
        user_id = user["id"]
        current_user_ids.add(user_id)
        user_principal = user.get("userPrincipalName", "N/A")
        
        # Progress indicator
        if idx % 50 == 0:
            print(f"  Processed {idx}/{len(current_users)} users...")
        
        # Get additional info
        manager = get_user_manager(headers, user_id)
        manager_name = manager.get("displayName") if manager else None
        manager_email = manager.get("mail") or manager.get("userPrincipalName") if manager else None
        
        onedrive_url = get_user_onedrive(headers, user_id)
        
        # Determine status
        is_active = user.get("accountEnabled", True)
        status = "active" if is_active else "disabled"
        
        # Check if user exists in existing CSV
        existing_user = next((u for u in existing_data if u.get('id') == user_id), None)
        
        # Determine change type
        change_type = "NO_CHANGE"
        if existing_user is None:
            change_type = "NEW_USER"
            new_count += 1
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
                updated_count += 1
        
        # Build user data for CSV
        user_data = {
            "id": user_id,
            "displayName": user.get("displayName", ""),
            "userPrincipalName": user_principal,
            "mail": user.get("mail", ""),
            "createdDateTime": user.get("createdDateTime", ""),
            "department": user.get("department", ""),
            "jobTitle": user.get("jobTitle", ""),
            "manager": manager_name if manager_name else "",
            "managerEmail": manager_email if manager_email else "",
            "onedrive_url": onedrive_url if onedrive_url else "",
            "status": status,
            "lastUpdated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "changeType": change_type
        }
        
        results.append(user_data)
    
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
    print(f"Total users in Azure AD: {len(current_users)}")
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
    print("="*60)
    print(f"Started at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print()
    
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
    
    # Fetch users from Azure AD
    print("\n3. Fetching users from Azure AD...")
    try:
        current_users = get_all_users(headers)
        print(f"   Found {len(current_users)} users in Azure AD")
    except Exception as e:
        print(f"   ✗ Error fetching users: {e}")
        return
    
    # Process users
    print("\n4. Processing user data...")
    processed_data = process_users(current_users, existing_data)
    
    # Save CSV report
    print("\n5. Saving CSV report...")
    save_csv_report(processed_data)
    
    print("\n" + "="*60)
    print("COMPLETED SUCCESSFULLY")
    print("="*60)
    print(f"Report saved to: {CSV_REPORT_FILE}")
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
