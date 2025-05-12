import requests
import json
import time
from datetime import datetime, timezone, timedelta
import os
import csv
from colorama import init, Fore, Style
from dateutil.relativedelta import relativedelta
from collections import OrderedDict
import pandas as pd
import glob

# Initialize colorama
init(autoreset=True)

# Configuration
TENANT_ID = "0e439a1f-a497-462b-9e6b-4e582e203607"
CLIENT_ID = "73efa35d-6188-42d4-b258-838a977eb149"
CLIENT_SECRET = "CyG8Q"
# Debug mode - SET TO FALSE FOR PRODUCTION!
DEBUG_MODE = False  # Shows full tokens when True

# Calculate automatic date range (runs on 2nd of each month)
current_date = datetime.now(timezone.utc)
if current_date.day != 2:
    current_date = current_date.replace(day=2)  # For testing purposes

# Set start date to previous month 1st at 04:00:00 UTC
start_date = (current_date - relativedelta(months=1)).replace(
    day=1, hour=4, minute=0, second=0, microsecond=0
)

# Set end date to current month 1st at 03:59:59 UTC
end_date = current_date.replace(
    day=1, hour=3, minute=59, second=59, microsecond=0
)

# Format as ISO strings
START_DATE = start_date.isoformat().replace('+00:00', 'Z')
END_DATE = end_date.isoformat().replace('+00:00', 'Z')

# Sites and operations
SITES = [
    "https://geekbyteonline.sharepoint.com/sites/New365",
    "https://geekbyteonline.sharepoint.com/sites/2DayRetention",
    "https://geekbyteonline.sharepoint.com/sites/geekbyte",
    "https://geekbyteonline.sharepoint.com/sites/geetkteam",
    "https://geekbyteonline.sharepoint.com/sites/New365Site5"
]

SITE_NAMES = [site.split('/')[-1] for site in SITES]  # For Excel report generation
OPERATIONS = ["PageViewed", "FileAccessed", "FileDownloaded"]
SERVICE_FILTER = "SharePoint"

# Constants
MAX_RETRIES = 5
RETRY_DELAY_SECONDS = 180  # Default wait time for rate limits

class AuditLogCollector:
    def __init__(self):
        self.access_token = None
        self.token_generation_time = None
        self.existing_searches = {}
        self.completed_searches = {}
        self.summary_data = []
        
    def log(self, message, color=Fore.WHITE):
        # Get local time with timezone
        local_time = datetime.now().astimezone()
        print(f"{color}[{local_time.strftime('%Y-%m-%d %H:%M:%S')}] {message}{Style.RESET_ALL}")
    
    def display_token_info(self):
        if self.access_token:
            if DEBUG_MODE:
                # SECURITY WARNING: Only for debugging
                self.log("=== SECURITY WARNING: FULL TOKEN DISPLAYED ===", Fore.RED)
                self.log(f"FULL TOKEN: {self.access_token}", Fore.RED)
                self.log("=== NEVER SHARE THIS TOKEN OR COMMIT TO VCS ===", Fore.RED)
            else:
                masked_token = f"{self.access_token[:10]}...{self.access_token[-10:]}"
                self.log(f"Current token (masked): {masked_token}", Fore.CYAN)
            
            if self.token_generation_time:
                now = datetime.now(timezone.utc)
                token_time = datetime.fromisoformat(
                    self.token_generation_time.replace('Z', '+00:00'))
                age = now - token_time
                self.log(f"Token generated at: {self.token_generation_time} (age: {age.total_seconds():.0f} seconds)", Fore.CYAN)
        else:
            self.log("No token available", Fore.YELLOW)
    
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
            self.token_generation_time = datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')
            self.log("New access token acquired", Fore.GREEN)
            self.display_token_info()
            return True
        except requests.exceptions.RequestException as e:
            self.log(f"Failed to get access token: {e}", Fore.RED)
            return False

    def make_api_request(self, method, url, params=None, json_data=None, max_retries=MAX_RETRIES):
        retry_count = 0
        while retry_count < max_retries:
            try:
                headers = {
                    "Authorization": f"Bearer {self.access_token}",
                    "Content-Type": "application/json"
                }
                
                response = requests.request(
                    method,
                    url,
                    headers=headers,
                    params=params,
                    json=json_data
                )
                
                if response.status_code == 401:
                    self.log("Token expired or invalid, refreshing...", Fore.YELLOW)
                    if not self.get_access_token():
                        retry_count += 1
                        time.sleep(RETRY_DELAY_SECONDS)
                        continue
                    # Retry with new token
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.request(
                        method,
                        url,
                        headers=headers,
                        params=params,
                        json=json_data
                    )
                
                # Handle rate limiting (429) and server errors (5xx)
                if response.status_code == 429 or response.status_code >= 500:
                    retry_after = int(response.headers.get('Retry-After', RETRY_DELAY_SECONDS))
                    self.log(f"Rate limited or server error (HTTP {response.status_code}), waiting {retry_after} seconds...", Fore.YELLOW)
                    time.sleep(retry_after)
                    retry_count += 1
                    continue
                
                response.raise_for_status()
                return response
            
            except requests.exceptions.ConnectionError:
                self.log(f"Connection error occurred, retrying... (attempt {retry_count + 1}/{max_retries})", Fore.YELLOW)
                retry_count += 1
                time.sleep(RETRY_DELAY_SECONDS)
                continue
                
            except requests.exceptions.Timeout:
                self.log(f"Request timed out, retrying... (attempt {retry_count + 1}/{max_retries})", Fore.YELLOW)
                retry_count += 1
                time.sleep(RETRY_DELAY_SECONDS)
                continue
                
            except requests.exceptions.RequestException as e:
                self.log(f"Request failed: {str(e)} - retrying... (attempt {retry_count + 1}/{max_retries})", Fore.YELLOW)
                retry_count += 1
                time.sleep(RETRY_DELAY_SECONDS)
                continue
        
        self.log(f"Max retries ({max_retries}) exceeded for this request", Fore.RED)
        return None

    def new_audit_search(self, site, operation):
        search_params = {
            "displayName": f"Audit_{site.split('/')[-1]}_{operation}_{datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')}",
            "filterStartDateTime": START_DATE,
            "filterEndDateTime": END_DATE,
            "operationFilters": [operation],
            "serviceFilters": [SERVICE_FILTER],
            "objectIdFilters": [f"{site}*"]
        }

        response = self.make_api_request(
            "POST",
            "https://graph.microsoft.com/beta/security/auditLog/queries",
            json_data=search_params
        )
        
        if not response:
            return None
            
        search_data = response.json()
        self.log(f"Search created for {site} - {operation}", Fore.GREEN)
        return search_data.get("id")

    def get_search_status(self, search_id):
        response = self.make_api_request(
            "GET",
            f"https://graph.microsoft.com/beta/security/auditLog/queries/{search_id}"
        )
        
        if not response:
            # If request failed, try refreshing token once before giving up
            self.log("Initial status check failed, attempting token refresh...", Fore.YELLOW)
            if self.get_access_token():
                response = self.make_api_request(
                    "GET",
                    f"https://graph.microsoft.com/beta/security/auditLog/queries/{search_id}"
                )
                if not response:
                    return "failed"
            else:
                return "failed"
            
        search_status = response.json()
        status = search_status.get("status")
        
        if status in ["succeeded", "failed"]:
            color = Fore.GREEN if status == "succeeded" else Fore.RED
            self.log(f"Search {search_id} completed with status: {status}", color)
        
        return status

    def get_audit_records(self, search_id):
        all_records = []
        url = f"https://graph.microsoft.com/beta/security/auditLog/queries/{search_id}/records?$top=1000"
        
        while url:
            response = self.make_api_request("GET", url)
            
            if not response:
                break
                
            data = response.json()
            all_records.extend(data.get("value", []))
            
            if "@odata.nextLink" in data:
                self.log(f"Retrieved {len(all_records)} records so far...", Fore.YELLOW)
                url = data["@odata.nextLink"]
                time.sleep(0.5)  # Small delay between pages
            else:
                url = None

        return all_records

    def clean_audit_data(self, audit_data):
        """Clean and format audit data with AppAccessContext first"""
        if isinstance(audit_data, str):
            try:
                audit_data = json.loads(audit_data)
            except json.JSONDecodeError:
                return OrderedDict()
        
        # Remove all @odata.type fields
        audit_data = {k: v for k, v in audit_data.items() if not k.endswith('@odata.type')}
        
        # Handle nested AppAccessContext
        app_access_context = OrderedDict()
        if "AppAccessContext" in audit_data and isinstance(audit_data["AppAccessContext"], dict):
            app_access_context = OrderedDict(
                (k, v) for k, v in audit_data["AppAccessContext"].items() 
                if not k.endswith('@odata.type')
            )
            del audit_data["AppAccessContext"]
        
        # Create new ordered dictionary with AppAccessContext first
        formatted_data = OrderedDict()
        
        # Add AppAccessContext first if it exists
        if app_access_context:
            formatted_data["AppAccessContext"] = app_access_context
        
        # Standard field order for remaining fields
        standard_fields = [
            "CreationTime", "Id", "Operation", "OrganizationId", "RecordType",
            "UserKey", "UserType", "Version", "Workload", "ClientIP", "UserId",
            "ApplicationId", "AuthenticationType", "BrowserName", "BrowserVersion",
            "CorrelationId", "EventSource", "GeoLocation", "IsManagedDevice",
            "ItemType", "ListId", "ListItemUniqueId", "Platform", "Site", "UserAgent",
            "WebId", "DeviceDisplayName", "HighPriorityMediaProcessing", "ListBaseType",
            "ListServerTemplate", "SiteUrl", "SourceRelativeUrl", "SourceFileName",
            "SourceFileExtension", "ApplicationDisplayName", "ObjectId"
        ]
        
        # Add standard fields in order
        for field in standard_fields:
            if field in audit_data:
                formatted_data[field] = audit_data[field]
        
        # Add any remaining fields not in our standard list
        for field, value in audit_data.items():
            if field not in formatted_data:
                formatted_data[field] = value
        
        return formatted_data

    def save_audit_to_csv(self, records, site, operation):
        if not records:
            self.log(f"No records found for {operation} in {site}", Fore.YELLOW)
            return 0

        site_name = site.split('/')[-1]
        start_day = START_DATE.split('T')[0].replace('-', '')
        end_day = END_DATE.split('T')[0].replace('-', '')
        current_time = datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')
        filename = f"{site_name}_{start_day}_{end_day}_{operation}_{current_time}.csv"

        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=["id", "createdDateTime", "userPrincipalName", "operation", "auditData"])
                writer.writeheader()
                for record in records:
                    # Clean and format the audit data
                    audit_data = record.get("auditData", {})
                    cleaned_data = self.clean_audit_data(audit_data)
                    
                    writer.writerow({
                        "id": record.get("id"),
                        "createdDateTime": record.get("createdDateTime"),
                        "userPrincipalName": record.get("userPrincipalName"),
                        "operation": record.get("operation"),
                        "auditData": json.dumps(cleaned_data, ensure_ascii=False, indent=None)
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

        current_time = datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S')
        summary_file = f"AuditSummary_{current_time}.csv"
        
        try:
            with open(summary_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=["Site", "Operation", "RecordCount", "StartDate", "EndDate"])
                writer.writeheader()
                for row in self.summary_data:
                    writer.writerow({
                        "Site": row["Site"],
                        "Operation": row["Operation"],
                        "RecordCount": row["RecordCount"],
                        "StartDate": START_DATE,
                        "EndDate": END_DATE
                    })
            self.log(f"Summary file generated: {summary_file}", Fore.GREEN)
        except Exception as e:
            self.log(f"Failed to generate summary: {e}", Fore.RED)

    def generate_excel_reports(self):
        """Generate Excel reports from the CSV files for each site"""
        self.log("Starting Excel report generation...", Fore.CYAN)
        
        # Get previous month and year for filename
        previous_month = datetime.now().replace(day=1) - pd.Timedelta(days=1)
        report_month_year = previous_month.strftime("%B-%Y")  # e.g. "Mar2025"

        # Define file mappings based on OPERATIONS
        file_mappings = {
            'PageViewed': ['*PageViewed*'],
            'FileAccessed': ['*FileAccessed*'], 
            'FileDownloaded': ['*FileDownloaded*']
        }

        for site_name in SITE_NAMES:
            # Create Excel filename
            excel_file = f"{site_name} {report_month_year}.xlsx"
            
            # Create Excel file with separate sheets
            with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                sheets_created = 0
                
                for sheet_name, patterns in file_mappings.items():
                    # Find matching CSV file
                    csv_file = None
                    for pattern in patterns:
                        full_pattern = f"{site_name}{pattern}.csv"
                        matches = glob.glob(full_pattern)
                        if matches:
                            csv_file = matches[0]
                            break
                    
                    if not csv_file:
                        self.log(f"Warning: No file found for {site_name} {sheet_name}", Fore.YELLOW)
                        continue
                    
                    try:
                        # Read CSV file
                        df = pd.read_csv(csv_file)
                        
                        # Write to Excel with exact sheet name
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        sheets_created += 1
                        
                        self.log(f"Added sheet: {sheet_name} (from {os.path.basename(csv_file)})", Fore.GREEN)
                    except Exception as e:
                        self.log(f"Error processing {csv_file}: {str(e)}", Fore.RED)
            
                if sheets_created > 0:
                    self.log(f"\n✅ Created {excel_file} with {sheets_created} sheets\n", Fore.GREEN)
                else:
                    self.log(f"\n⚠️ No sheets created for {site_name}\n", Fore.YELLOW)
                    try:
                        os.remove(excel_file)  # Remove empty file
                    except Exception as e:
                        self.log(f"Failed to remove empty Excel file: {e}", Fore.YELLOW)

        self.log("Excel report generation complete!", Fore.GREEN)

    def load_state(self, filename):
        if os.path.exists(filename):
            try:
                with open(filename, 'r') as f:
                    return json.load(f)
            except Exception as e:
                self.log(f"Failed to load state file {filename}: {e}", Fore.YELLOW)
                return {}
        return {}

    def save_state(self, filename, data):
        try:
            with open(filename, 'w') as f:
                json.dump(data, f, indent=4)
            self.log(f"Saved state to {filename}", Fore.CYAN)
        except Exception as e:
            self.log(f"Failed to save state file {filename}: {e}", Fore.RED)

    def archive_search_ids(self):
        if os.path.exists("search_ids.json"):
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            new_name = f"searchIds_{timestamp}.json"
            try:
                os.rename("search_ids.json", new_name)
                self.log(f"Archived search_ids.json as {new_name}", Fore.GREEN)
                return True
            except Exception as e:
                self.log(f"Failed to archive search_ids.json: {e}", Fore.RED)
                return False
        return True

    def process_pending_searches(self):
        total_searches = len(SITES) * len(OPERATIONS)
        processed_searches = 0

        # First create any missing searches
        for site in SITES:
            for operation in OPERATIONS:
                key = f"{site}_{operation}"
                
                if key in self.completed_searches:
                    processed_searches += 1
                    continue
                    
                if key not in self.existing_searches:
                    search_id = None
                    attempts = 0
                    while not search_id and attempts < MAX_RETRIES:
                        search_id = self.new_audit_search(site, operation)
                        if not search_id:
                            time.sleep(RETRY_DELAY_SECONDS)
                            attempts += 1
                    
                    if search_id:
                        self.existing_searches[key] = {
                            "SearchId": search_id,
                            "CreatedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')
                        }
                        self.save_state("search_ids.json", self.existing_searches)
                        processed_searches += 1
                else:
                    processed_searches += 1

        # Then process all searches (existing and new)
        while len(self.completed_searches) < total_searches:
            self.log(f"Processing searches ({len(self.completed_searches)}/{total_searches} completed)...", Fore.YELLOW)
            
            for key, search_data in self.existing_searches.items():
                if key in self.completed_searches:
                    continue
                    
                status = self.get_search_status(search_data["SearchId"])
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
                        "CompletedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
                        "RecordCount": record_count
                    }
                    self.save_state("completed_searches.json", self.completed_searches)
                    
                elif status == "failed":
                    self.completed_searches[key] = {
                        "CompletedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
                        "RecordCount": 0,
                        "Status": "failed"
                    }
                    self.save_state("completed_searches.json", self.completed_searches)
            
            if len(self.completed_searches) < total_searches:
                time.sleep(RETRY_DELAY_SECONDS)  # Reduced wait time between checks

    def run(self):
        try:
            self.log("Starting audit log collection process...", Fore.CYAN)
            self.log(f"Date range: {START_DATE} to {END_DATE}", Fore.CYAN)
            
            # Authentication
            self.log("Authenticating...", Fore.YELLOW)
            if not self.get_access_token():
                self.log("Failed to authenticate, exiting...", Fore.RED)
                return
            
            # Load existing state
            self.existing_searches = self.load_state("search_ids.json")
            self.completed_searches = self.load_state("completed_searches.json")
            self.log(f"Loaded {len(self.existing_searches)} existing searches and {len(self.completed_searches)} completed searches", Fore.YELLOW)
            
            # Process all searches
            self.process_pending_searches()
            
            # Generate summary
            self.generate_summary()
            
            # Generate Excel reports
            self.generate_excel_reports()
            
            # Archive search IDs
            if len(self.completed_searches) == len(SITES) * len(OPERATIONS):
                if self.archive_search_ids():
                    # Only remove if archiving was successful
                    if os.path.exists("completed_searches.json"):
                        try:
                            os.remove("completed_searches.json")
                        except Exception as e:
                            self.log(f"Failed to remove completed_searches.json: {e}", Fore.YELLOW)
            
            self.log("All operations completed successfully!", Fore.GREEN)
                
        except Exception as e:
            self.log(f"Fatal error: {e}", Fore.RED)
            raise

if __name__ == "__main__":
    collector = AuditLogCollector()
    collector.run()
