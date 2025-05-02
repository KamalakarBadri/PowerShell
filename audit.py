
import requests
import json
import time
from datetime import datetime
import os
import csv
from colorama import init, Fore, Style

# Initialize colorama
init(autoreset=True)

# Configuration
TENANT_ID = "0e439a1f-a497-462b-9e6b-4e582e203607"
CLIENT_ID = "73efa35d-6188-42d4-b258-838a977eb149"
CLIENT_SECRET = "CyGuCMSyVmt4sNxt5IejrMc2c24Ziz4a.t"

# Date range (UTC)
START_DATE = "2025-01-10T00:00:00Z"
END_DATE = "2025-05-02T23:59:59Z"

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
MAX_CONCURRENT_SEARCHES = 10
WAIT_TIME_SECONDS = 300
RETRY_DELAY_SECONDS = 30

class AuditLogCollector:
    def __init__(self):
        self.access_token = None
        self.active_search_count = 0
        self.existing_searches = {}
        self.completed_searches = {}
        self.summary_data = []
        
    def log(self, message, color=Fore.WHITE):
        print(f"{color}[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}{Style.RESET_ALL}")
    
    def get_access_token(self):
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
            self.access_token = response.json().get("access_token")
            self.log("New access token acquired", Fore.GREEN)
        except Exception as e:
            self.log(f"Failed to get access token: {e}", Fore.RED)
            raise

    def update_active_search_count(self):
        """Check status of active searches and update count"""
        completed = 0
        for key, search_data in self.existing_searches.items():
            if key not in self.completed_searches:
                status = self.get_search_status(search_data["SearchId"])
                if status in ["succeeded", "failed"]:
                    completed += 1
        self.active_search_count = max(0, self.active_search_count - completed)

    def new_audit_search(self, site, operation):
        # Wait until a slot becomes available
        while self.active_search_count >= MAX_CONCURRENT_SEARCHES:
            self.log(f"Concurrency limit reached ({self.active_search_count}/{MAX_CONCURRENT_SEARCHES}). Waiting {WAIT_TIME_SECONDS} seconds...", Fore.RED)
            time.sleep(WAIT_TIME_SECONDS)
            self.update_active_search_count()

        search_params = {
            "displayName": f"Audit_{site.split('/')[-1]}_{operation}_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
            "filterStartDateTime": START_DATE,
            "filterEndDateTime": END_DATE,
            "operationFilters": [operation],
            "serviceFilters": [SERVICE_FILTER],
            "objectIdFilters": [f"{site}*"]
        }

        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        try:
            self.log(f"Creating search for {site} - {operation}", Fore.CYAN)
            response = requests.post(
                "https://graph.microsoft.com/beta/security/auditLog/queries",
                headers=headers,
                json=search_params
            )
            response.raise_for_status()
            
            search_data = response.json()
            self.active_search_count += 1
            self.log(f"Search created (current active: {self.active_search_count})", Fore.GREEN)
            return search_data.get("id")
        except requests.exceptions.HTTPError as e:
            if e.response.status_code in [429, 503]:
                self.log(f"API rate limit hit. Waiting {WAIT_TIME_SECONDS} seconds...", Fore.RED)
                time.sleep(WAIT_TIME_SECONDS)
                return None
            else:
                self.log(f"Failed to create audit search: {e}", Fore.RED)
                return None

    def get_search_status(self, search_id):
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        try:
            response = requests.get(
                f"https://graph.microsoft.com/beta/security/auditLog/queries/{search_id}",
                headers=headers
            )
            response.raise_for_status()
            
            search_status = response.json()
            status = search_status.get("status")
            
            if status in ["succeeded", "failed"]:
                color = Fore.GREEN if status == "succeeded" else Fore.RED
                self.log(f"Search {search_id} completed with status: {status}", color)
            
            return status
        except requests.exceptions.RequestException as e:
            try:
                error_details = e.response.json()
                if "expired" in error_details.get("error", "").lower():
                    self.log("Token expired, refreshing...", Fore.YELLOW)
                    self.get_access_token()
                    return self.get_search_status(search_id)
            except:
                pass
            
            self.log(f"Error checking status, will retry: {e}", Fore.YELLOW)
            time.sleep(2)
            return "retry"

    def get_audit_records(self, search_id):
        all_records = []
        url = f"https://graph.microsoft.com/beta/security/auditLog/queries/{search_id}/records?$top=1000"
        
        while url:
            try:
                headers = {
                    "Authorization": f"Bearer {self.access_token}",
                    "Content-Type": "application/json"
                }
                response = requests.get(url, headers=headers)
                response.raise_for_status()
                
                data = response.json()
                all_records.extend(data.get("value", []))
                
                if "@odata.nextLink" in data:
                    self.log(f"Retrieved {len(all_records)} records so far...", Fore.YELLOW)
                    url = data["@odata.nextLink"]
                    time.sleep(0.5)
                else:
                    url = None
                    
            except requests.exceptions.RequestException as e:
                try:
                    error_details = e.response.json()
                    if "expired" in error_details.get("error", "").lower():
                        self.log("Token expired, refreshing...", Fore.YELLOW)
                        self.get_access_token()
                        continue
                except:
                    pass
                
                self.log(f"Error retrieving records, will retry: {e}", Fore.YELLOW)
                time.sleep(2)
                continue

        return all_records

    def save_audit_to_csv(self, records, site, operation):
        if not records:
            self.log(f"No records found for {operation} in {site}", Fore.YELLOW)
            return 0

        site_name = site.split('/')[-1]
        start_day = START_DATE.split('T')[0].replace('-', '')
        end_day = END_DATE.split('T')[0].replace('-', '')
        current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{site_name}_{start_day}_{end_day}_{operation}_{current_time}.csv"

        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=["id", "createdDateTime", "userPrincipalName", "operation", "auditData"])
                writer.writeheader()
                for record in records:
                    writer.writerow({
                        "id": record.get("id"),
                        "createdDateTime": record.get("createdDateTime"),
                        "userPrincipalName": record.get("userPrincipalName"),
                        "operation": record.get("operation"),
                        "auditData": json.dumps(record.get("auditData", {}))
                    })
            
            self.log(f"Saved {len(records)} records to {filename}", Fore.GREEN)
            return len(records)
        except Exception as e:
            self.log(f"Failed to save CSV: {e}", Fore.RED)
            return 0

    def generate_summary(self):
        if not self.summary_data:
            self.log("No data available for summary", Fore.YELLOW)
            return

        summary_file = "AuditSummary.csv"
        try:
            with open(summary_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=["Site", "Operation", "RecordCount"])
                writer.writeheader()
                writer.writerows(self.summary_data)
            self.log(f"Summary file generated: {summary_file}", Fore.GREEN)
        except Exception as e:
            self.log(f"Failed to generate summary: {e}", Fore.RED)

    def load_state(self, filename):
        if os.path.exists(filename):
            with open(filename, 'r') as f:
                return json.load(f)
        return {}

    def save_state(self, filename, data):
        with open(filename, 'w') as f:
            json.dump(data, f)

    def run(self):
        try:
            self.log("Starting audit log collection process...", Fore.CYAN)
            
            # Authentication
            self.log("Authenticating...", Fore.YELLOW)
            self.get_access_token()
            
            # Load existing state
            self.existing_searches = self.load_state("search_ids.json")
            self.completed_searches = self.load_state("completed_searches.json")
            self.log(f"Loaded {len(self.existing_searches)} existing searches", Fore.YELLOW)
            
            # Initialize active search count
            self.active_search_count = sum(
                1 for key in self.existing_searches 
                if key not in self.completed_searches
            )
            
            # Create searches
            total_searches = len(SITES) * len(OPERATIONS)
            created_searches = 0
            
            for site in SITES:
                for operation in OPERATIONS:
                    key = f"{site}_{operation}"
                    
                    if key in self.completed_searches:
                        created_searches += 1
                        continue
                        
                    if key not in self.existing_searches:
                        search_id = None
                        while not search_id:
                            search_id = self.new_audit_search(site, operation)
                            if not search_id:
                                # This handles API rate limits (429/503)
                                time.sleep(RETRY_DELAY_SECONDS)
                        
                        if search_id:
                            self.existing_searches[key] = {
                                "SearchId": search_id,
                                "CreatedTime": datetime.now().isoformat() + "Z"
                            }
                            self.save_state("search_ids.json", self.existing_searches)
                            created_searches += 1
                    else:
                        created_searches += 1
            
            # Process searches
            for key, search_data in self.existing_searches.items():
                if key in self.completed_searches:
                    continue
                    
                status = None
                while status not in ["succeeded", "failed"]:
                    status = self.get_search_status(search_data["SearchId"])
                    if status not in ["succeeded", "failed"]:
                        time.sleep(RETRY_DELAY_SECONDS)
                
                if status == "succeeded":
                    site, operation = key.split("_", 1)
                    self.log(f"Retrieving records for {site} - {operation}", Fore.CYAN)
                    records = self.get_audit_records(search_data["SearchId"])
                    
                    record_count = self.save_audit_to_csv(records, site, operation)
                    self.summary_data.append({
                        "Site": site.split('/')[-1],
                        "Operation": operation,
                        "RecordCount": record_count
                    })
                    
                    self.completed_searches[key] = {
                        "CompletedTime": datetime.now().isoformat() + "Z",
                        "RecordCount": record_count
                    }
                    self.save_state("completed_searches.json", self.completed_searches)
            
            # Generate summary
            self.generate_summary()
            
            # Cleanup
            if len(self.completed_searches) == total_searches:
                self.log("All operations completed successfully!", Fore.CYAN)
                for f in ["search_ids.json", "completed_searches.json"]:
                    if os.path.exists(f):
                        os.remove(f)
            else:
                remaining = total_searches - len(self.completed_searches)
                self.log(f"{remaining} searches remaining. Run again to continue.", Fore.YELLOW)
                
        except Exception as e:
            self.log(f"Fatal error: {e}", Fore.RED)
            raise

if __name__ == "__main__":
    collector = AuditLogCollector()
    collector.run()