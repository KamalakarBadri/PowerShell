import requests
import json
import time
from datetime import datetime
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
import csv

# Configuration
TENANT_ID = "0e439a1f-a497-462b-9e6b-4e582e203607"
CLIENT_ID = "73efa35d-6188-42d4-b258-838a977eb149"
CLIENT_SECRET = "CyG8Q~FYHuCMSyVm4a.t"

# Date range (UTC)
START_DATE = "2025-01-10T00:00:00Z"
END_DATE = "2025-04-02T23:59:59Z"

# Extract dates for filename
start_day = START_DATE.split('T')[0].replace('-', '')
end_day = END_DATE.split('T')[0].replace('-', '')

# Sites and operations
SITES = [
    "https://geekbyteonline.sharepoint.com/sites/New365",
    "https://geekbyteonline.sharepoint.com/sites/2DayRetention",
    "https://geekbyteonline.sharepoint.com/sites/geekbyte",
    "https://geekbyteonline.sharepoint.com/sites/geetkteam",
    "https://geekbyteonline.sharepoint.com/sites/New365Site5"
]

OPERATIONS = ["PageViewed", "FileAccessed", "FileDownloaded"]
SERVICE_FILTER = "SharePoint"

# Constants
MAX_CONCURRENT_SEARCHES = 10  # Microsoft Graph limit
WAIT_TIME_SECONDS = 300       # 5 minutes wait time when limit reached
RETRY_DELAY_SECONDS = 30      # Delay between status checks

# Global variables
access_token = None
active_search_count = 0
existing_searches = {}
completed_searches = {}
summary_data = []

def get_current_time():
    return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def get_access_token():
    global access_token
    auth_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    body = {
        "client_id": CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials"
    }

    try:
        response = requests.post(auth_url, data=body)
        response.raise_for_status()
        token_data = response.json()
        print(f"[{get_current_time()}] New access token acquired")
        access_token = token_data.get("access_token")
        return access_token
    except requests.exceptions.RequestException as e:
        print(f"[{get_current_time()}] Failed to get access token: {e}")
        exit(1)

def new_audit_search(site, operation):
    global active_search_count, access_token
    
    search_params = {
        "displayName": f"Audit_{site.split('/')[-1]}_{operation}_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
        "filterStartDateTime": START_DATE,
        "filterEndDateTime": END_DATE,
        "operationFilters": [operation],
        "serviceFilters": [SERVICE_FILTER],
        "objectIdFilters": [f"{site}*"]
    }

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    try:
        # Check if we've reached the concurrent search limit
        if active_search_count >= MAX_CONCURRENT_SEARCHES:
            print(f"[{get_current_time()}] Waiting for available search slot (current: {active_search_count}/{MAX_CONCURRENT_SEARCHES})...")
            return None

        print(f"[{get_current_time()}] Creating search for {site} - {operation}")
        response = requests.post(
            "https://graph.microsoft.com/beta/security/auditLog/queries",
            headers=headers,
            json=search_params
        )
        response.raise_for_status()
        
        search_data = response.json()
        active_search_count += 1
        print(f"[{get_current_time()}] Search created (current active: {active_search_count})")
        return search_data.get("id")
    except requests.exceptions.RequestException as e:
        if hasattr(e, 'response') and e.response.status_code in [429, 503]:
            print(f"[{get_current_time()}] Rate limit hit when creating search. Waiting {WAIT_TIME_SECONDS} seconds...")
            time.sleep(WAIT_TIME_SECONDS)
            return None
        else:
            print(f"[{get_current_time()}] Failed to create audit search for {site} - {operation}: {e}")
            return None

def get_search_status(search_id):
    global active_search_count, access_token
    
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    try:
        response = requests.get(
            f"https://graph.microsoft.com/beta/security/auditLog/queries/{search_id}",
            headers=headers
        )
        response.raise_for_status()
        
        search_status = response.json()
        
        if search_status.get("status") in ["succeeded", "failed"]:
            active_search_count -= 1
            print(f"[{get_current_time()}] Search {search_id} completed with status: {search_status.get('status')}")
        
        return search_status.get("status")
    except requests.exceptions.RequestException as e:
        try:
            error_details = e.response.json()
            if error_details.get("error", "").lower() in ["expired", "invalidauthenticationtoken"]:
                print(f"[{get_current_time()}] Token expired, refreshing...")
                access_token = get_access_token()
                return get_search_status(search_id)
        except:
            pass
        
        # For any other error, wait 2 seconds and retry
        print(f"[{get_current_time()}] Error checking status, will retry in 2 seconds: {e}")
        time.sleep(2)
        return get_search_status(search_id)

def get_audit_records(search_id):
    global access_token
    
    all_records = []
    url = f"https://graph.microsoft.com/beta/security/auditLog/queries/{search_id}/records?$top=1000"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    while url:
        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            data = response.json()
            all_records.extend(data.get("value", []))
            
            if "@odata.nextLink" in data:
                print(f"[{get_current_time()}] Retrieved {len(all_records)} records...")
                url = data["@odata.nextLink"]
                time.sleep(0.5)
            else:
                url = None
        except requests.exceptions.RequestException as e:
            try:
                error_details = e.response.json()
                if error_details.get("error", "").lower() in ["expired", "invalidauthenticationtoken"]:
                    print(f"[{get_current_time()}] Token expired, refreshing...")
                    access_token = get_access_token()
                    headers["Authorization"] = f"Bearer {access_token}"
                    continue
            except:
                pass
            
            # For any other error, wait 2 seconds and retry
            print(f"[{get_current_time()}] Error retrieving records, will retry in 2 seconds: {e}")
            time.sleep(2)
            continue

    return all_records

def save_audit_to_csv(records, site, operation):
    if not records:
        print(f"[{get_current_time()}] No records found for {operation} in {site}")
        return 0

    # Extract site name for filename
    site_name = site.split('/')[-1]
    current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f"{site_name}_{start_day}_{end_day}_{operation}_{current_time}.csv"

    report = []
    for record in records:
        report.append({
            "id": record.get("id"),
            "createdDateTime": record.get("createdDateTime"),
            "userPrincipalName": record.get("userPrincipalName"),
            "operation": record.get("operation"),
            "auditData": json.dumps(record.get("auditData", {}), indent=None)
        })

    try:
        mode = "a" if os.path.exists(filename) else "w"
        with open(filename, mode, newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=report[0].keys())
            if mode == "w":
                writer.writeheader()
            writer.writerows(report)
        
        print(f"[{get_current_time()}] Saved {len(records)} records to {filename}")
        return len(records)
    except Exception as e:
        print(f"[{get_current_time()}] Failed to save CSV: {e}")
        return 0

def generate_summary_file():
    if not summary_data:
        print(f"[{get_current_time()}] No data available for summary")
        return

    summary_filename = f"AuditSummary_{start_day}_{end_day}.csv"
    try:
        with open(summary_filename, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=["Site", "Operation", "RecordCount"])
            writer.writeheader()
            writer.writerows(summary_data)
        print(f"[{get_current_time()}] Summary file generated: {summary_filename}")
    except Exception as e:
        print(f"[{get_current_time()}] Failed to generate summary file: {e}")

def load_searches_from_file(filename):
    if os.path.exists(filename):
        with open(filename, "r") as f:
            return json.load(f)
    return {}

def save_searches_to_file(filename, data):
    with open(filename, "w") as f:
        json.dump(data, f)

def process_search(key, search_data):
    global completed_searches, access_token, summary_data
    
    search_id = search_data["SearchId"]
    status = None
    
    while status not in ["succeeded", "failed"]:
        status = get_search_status(search_id)
        
        if status == "retry":
            time.sleep(RETRY_DELAY_SECONDS)
            continue

        if status == "succeeded":
            # Extract site and operation from key
            parts = key.split("_", 1)
            site = parts[0]
            operation = parts[1]

            # Retrieve records
            print(f"[{get_current_time()}] Retrieving records for {site} - {operation}")
            records = get_audit_records(search_id)

            # Save to CSV and get record count
            record_count = save_audit_to_csv(records, site, operation)

            # Add to summary data
            site_name = site.split('/')[-1]
            summary_data.append({
                "Site": site_name,
                "Operation": operation,
                "RecordCount": record_count
            })

            # Mark as completed
            completed_searches[key] = {
                "CompletedTime": datetime.now().isoformat() + "Z",
                "RecordCount": record_count
            }
            save_searches_to_file("completed_searches.json", completed_searches)
            
        elif status == "failed":
            print(f"[{get_current_time()}] Search {search_id} failed")
            completed_searches[key] = {
                "CompletedTime": datetime.now().isoformat() + "Z",
                "Status": "failed"
            }
            save_searches_to_file("completed_searches.json", completed_searches)
        else:
            # Still running, wait before checking again
            print(f"[{get_current_time()}] Search {search_id} status: {status}. Waiting {RETRY_DELAY_SECONDS} seconds...")
            time.sleep(RETRY_DELAY_SECONDS)

def main():
    global access_token, existing_searches, completed_searches, active_search_count, summary_data
    
    print(f"[{get_current_time()}] Starting audit log collection process...")

    # Step 1: Authenticate
    print(f"[{get_current_time()}] Authenticating...")
    access_token = get_access_token()

    # Step 2: Load existing searches from tracking files
    existing_searches = load_searches_from_file("search_ids.json")
    print(f"[{get_current_time()}] Loaded {len(existing_searches)} existing searches from file.")

    completed_searches = load_searches_from_file("completed_searches.json")

    # Step 3: Create searches with proper throttling
    total_searches = len(SITES) * len(OPERATIONS)
    created_searches = 0

    for site in SITES:
        for operation in OPERATIONS:
            key = f"{site}_{operation}"
            
            # Skip if already completed
            if key in completed_searches:
                created_searches += 1
                continue

            # Create new search if needed
            if key not in existing_searches:
                search_id = None
                while not search_id:
                    search_id = new_audit_search(site, operation)
                    
                    if not search_id:
                        # Wait if we hit the limit or got throttled
                        time.sleep(RETRY_DELAY_SECONDS)

                if search_id:
                    existing_searches[key] = {
                        "SearchId": search_id,
                        "CreatedTime": datetime.now().isoformat() + "Z"
                    }
                    save_searches_to_file("search_ids.json", existing_searches)
                    created_searches += 1
            else:
                created_searches += 1

    # Step 4: Process searches with proper concurrency management
    search_keys = [k for k in existing_searches.keys() if k not in completed_searches]
    
    # Using ThreadPoolExecutor to manage concurrent searches
    with ThreadPoolExecutor(max_workers=MAX_CONCURRENT_SEARCHES) as executor:
        futures = []
        for key in search_keys:
            futures.append(executor.submit(process_search, key, existing_searches[key]))
        
        for future in as_completed(futures):
            try:
                future.result()
            except Exception as e:
                print(f"[{get_current_time()}] Error processing search: {e}")

    # Generate summary file
    generate_summary_file()

    # Final cleanup
    if len(completed_searches) == total_searches:
        print(f"[{get_current_time()}] All operations completed successfully!")
        if os.path.exists("search_ids.json"):
            os.remove("search_ids.json")
        if os.path.exists("completed_searches.json"):
            os.remove("completed_searches.json")
    else:
        remaining = total_searches - len(completed_searches)
        print(f"[{get_current_time()}] {remaining} searches remaining. Run the script again to continue.")

if __name__ == "__main__":
    main()
