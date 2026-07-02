import requests
import json
import time
from datetime import datetime, timezone
import os
import csv
from colorama import init, Fore, Style
from dateutil.relativedelta import relativedelta
from collections import OrderedDict
import glob
import math
import logging
from logging.handlers import RotatingFileHandler

# Initialize colorama
init(autoreset=True)

# Configuration
TENANT_ID = "0e439a1f-a497-462b-9e6b-4e582e203607"
CLIENT_ID = "73efa35d-6188-42d4-b258-838a977eb149"
CLIENT_SECRET = "CyG8Q~FYHuCMSyVmt4sNxt5IejrMc2c24Ziz4a.t"
ACCESS_TOKEN2 = "invalid"
# SharePoint Upload Configuration
DOCUMENT_LIBRARY_NAME = "Site Analytics"
CHUNK_SIZE = 3276800  # 3.125 MB
# Debug mode - SET TO FALSE FOR PRODUCTION!
DEBUG_MODE = False  # Shows full tokens when True

# Date range configuration
USE_MANUAL_DATE_RANGE = False
# Example format: "2026-02-01T04:00:00Z"
MANUAL_START_DATE = "2026-02-01T04:00:00Z"
MANUAL_END_DATE = "2026-03-01T03:59:59Z"

# CSV report mode options:
# "RAW"             -> id, createdDateTime, userPrincipalName, operation, auditData as JSON
# "REQUIRED_FIELDS" -> only selected important fields
# "ALL_FIELDS"      -> all fields inside auditData expanded as CSV columns
CSV_REPORT_MODE = "RAW"

# Sites - collect all operations (no operation filtering)
SITES = [
    "https://geekbyteonline.sharepoint.com/sites/2DayRetention",
    "https://geekbyteonline.sharepoint.com/sites/Geekbyteteam"
]

# Constants
MAX_RETRIES = 10
RETRY_DELAY_SECONDS = 10
MAX_SEARCH_STATUS_POLLS = 180

class AuditLogCollector:
    def __init__(self):
        # Fixed state directory (never changes between runs)
        self.state_dir = "audit_state"
        os.makedirs(self.state_dir, exist_ok=True)
        
        # Create new output directory for reports
        self.run_timestamp = datetime.now().astimezone().strftime('%Y%m%d_%H%M%S')
        self.output_dir = f"AuditLogs_{self.run_timestamp}"
        os.makedirs(self.output_dir, exist_ok=True)
        
        # Initialize logging
        self.setup_logging()
        
        # Initialize variables
        self.access_token = None
        self.token_generation_time = None
        self.existing_searches = {}
        self.completed_searches = {}
        self.summary_data = []
        self.process_start_time = None
        self.process_end_time = None
        self.uploaded_files = []
        self.site_library_cache = {}
        self.search_poll_counts = {}
        
        # Get date ranges
        self.set_date_ranges()
        
    def setup_logging(self):
        """Configure logging to both console and file"""
        self.logger = logging.getLogger('AuditLogCollector')
        self.logger.setLevel(logging.DEBUG)
        
        # Create formatter
        formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
        
        # Console handler
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(formatter)
        self.logger.addHandler(ch)
        
        # File handler with rotation (10MB max, 5 backups)
        log_file = os.path.join(self.output_dir, 'audit_log_collector.log')
        fh = RotatingFileHandler(
            log_file, 
            maxBytes=10*1024*1024, 
            backupCount=5,
            encoding='utf-8'
        )
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        self.logger.addHandler(fh)
        
        # Log startup message
        self.log("Initialized Audit Log Collector", Fore.CYAN)
        self.log(f"State directory: {os.path.abspath(self.state_dir)}", Fore.CYAN)
        self.log(f"Output directory: {os.path.abspath(self.output_dir)}", Fore.CYAN)
    
    def set_date_ranges(self):
        """Set the date ranges for the report"""
        if USE_MANUAL_DATE_RANGE:
            self.start_date = datetime.fromisoformat(
                MANUAL_START_DATE.replace('Z', '+00:00')
            ).astimezone(timezone.utc)
            self.end_date = datetime.fromisoformat(
                MANUAL_END_DATE.replace('Z', '+00:00')
            ).astimezone(timezone.utc)
            self.log("Using manual date range from configuration", Fore.CYAN)
        else:
            # Get current datetime in system timezone
            current_date = datetime.now().astimezone()

            # Set start date to previous month 1st at 04:00:00 UTC
            self.start_date = (current_date - relativedelta(months=1)).replace(
                day=1, hour=4, minute=0, second=0, microsecond=0, tzinfo=timezone.utc
            )

            # Set end date to current month 1st at 03:59:59 UTC
            self.end_date = current_date.replace(
                day=1, hour=3, minute=59, second=59, microsecond=0, tzinfo=timezone.utc
            )

        # Format as ISO strings
        self.START_DATE = self.start_date.isoformat().replace('+00:00', 'Z')
        self.END_DATE = self.end_date.isoformat().replace('+00:00', 'Z')

        # Get month and year for file properties
        self.REPORT_MONTH = self.start_date.strftime("%B")
        self.REPORT_YEAR = self.start_date.strftime("%Y")
        
        self.log(f"Date range set to: {self.START_DATE} to {self.END_DATE}", Fore.CYAN)
        self.log(f"Report Month/Year: {self.REPORT_MONTH} {self.REPORT_YEAR}", Fore.CYAN)
    
    def log(self, message, color=Fore.WHITE, level=logging.INFO):
        """Log message to both console (with color) and log file"""
        local_time = datetime.now().astimezone()
        console_message = f"{color}[{local_time.strftime('%Y-%m-%d %H:%M:%S')}] {message}{Style.RESET_ALL}"
        print(console_message)
        self.logger.log(level, message)
    
    def display_token_info(self):
        if self.access_token:
            if DEBUG_MODE:
                self.log("=== SECURITY WARNING: FULL TOKEN DISPLAYED ===", Fore.RED)
                self.log(f"FULL TOKEN: {self.access_token}", Fore.RED)
                self.log("=== NEVER SHARE THIS TOKEN OR COMMIT TO VCS ===", Fore.RED)
            else:
                masked_token = f"{self.access_token[:10]}...{self.access_token[-10:]}"
                self.log(f"Current token (masked): {masked_token}", Fore.CYAN)
            
            if self.token_generation_time:
                now = datetime.now().astimezone()
                token_time = datetime.fromisoformat(
                    self.token_generation_time.replace('Z', '+00:00')).astimezone()
                age = now - token_time
                self.log(f"Token generated at: {token_time.strftime('%Y-%m-%d %H:%M:%S')} (age: {age.total_seconds():.0f} seconds)", Fore.CYAN)
        else:
            self.log("No token available", Fore.YELLOW)
    
    def get_access_token(self):
        """Get access token using client credentials flow"""
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
            self.log(f"Failed to get access token: {e}", Fore.RED, logging.ERROR)
            return False

    def make_api_request(self, method, url, params=None, json_data=None, data=None, headers=None, max_retries=MAX_RETRIES):
        """Generic API request method with token refresh and retry logic"""
        retry_count = 0
        while retry_count < max_retries:
            try:
                if headers is None:
                    headers = {
                        "Authorization": f"Bearer {self.access_token}",
                        "Content-Type": "application/json"
                    }
                elif "Authorization" not in headers:
                    headers["Authorization"] = f"Bearer {self.access_token}"
                
                response = requests.request(
                    method,
                    url,
                    headers=headers,
                    params=params,
                    json=json_data,
                    data=data
                )
                
                if response.status_code == 401:
                    self.log("Token expired or invalid, refreshing...", Fore.RED)
                    if not self.get_access_token():
                        retry_count += 1
                        time.sleep(RETRY_DELAY_SECONDS)
                        continue
                    headers["Authorization"] = f"Bearer {self.access_token}"
                    response = requests.request(
                        method,
                        url,
                        headers=headers,
                        params=params,
                        json=json_data,
                        data=data
                    )
                
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
        
        self.log(f"Max retries ({max_retries}) exceeded for this request", Fore.RED, logging.ERROR)
        return None
    
    def check_existing_search(self, site):
        """Check if a search already exists for the site with the same date range"""
        response = self.make_api_request(
            "GET",
            "https://graph.microsoft.com/beta/security/auditLog/queries"
        )
        
        if not response:
            return None
            
        searches = response.json().get('value', [])
        
        for search in searches:
            if (search.get('filterStartDateTime') == self.START_DATE and
                search.get('filterEndDateTime') == self.END_DATE and
                f"{site}/*" in search.get('objectIdFilters', [])):
                
                self.log(f"Found existing search with matching criteria: {search['id']}", Fore.CYAN)
                return search['id']
                
        return None

    def new_audit_search(self, site):
        """Create a new audit log search for the site (all operations)"""
        existing_search_id = self.check_existing_search(site)
        if existing_search_id:
            return existing_search_id
        
        search_params = {
            "displayName": f"Audit_{site.split('/')[-1]}_{datetime.now().astimezone().strftime('%Y%m%d_%H%M%S')}",
            "filterStartDateTime": self.START_DATE,
            "filterEndDateTime": self.END_DATE,
            "objectIdFilters": [f"{site}/*"]
        }

        response = self.make_api_request(
            "POST",
            "https://graph.microsoft.com/beta/security/auditLog/queries",
            json_data=search_params
        )
        
        if not response:
            return None
            
        search_data = response.json()
        self.log(f"Search created for {site}", Fore.GREEN)
        return search_data.get("id")

    def verify_all_searches_created(self):
        """Verify that all required searches exist before processing"""
        self.log("Verifying all required searches exist...", Fore.CYAN)
        all_searches_created = True
        
        for site in SITES:
            # First, check if we have a search ID in our local state
            search_id = None
            
            # Try to find existing search in Microsoft Graph
            existing_id = self.check_existing_search(site)
            if existing_id:
                search_id = existing_id
                self.log(f"Found existing search for {site} in Microsoft Graph", Fore.GREEN)
            else:
                # Only create new if none exists
                self.log(f"No existing search found for {site}, creating new...", Fore.YELLOW)
                attempts = 0
                while not search_id and attempts < MAX_RETRIES:
                    search_id = self.new_audit_search(site)
                    if not search_id:
                        time.sleep(RETRY_DELAY_SECONDS)
                        attempts += 1
            
            if search_id:
                self.existing_searches[site] = {
                    "SearchId": search_id,
                    "CreatedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')
                }
                self.save_state("search_ids.json", self.existing_searches)
                self.log(f"Using search for {site}", Fore.GREEN)
            else:
                self.log(f"Failed to create/find search for {site}", Fore.RED)
                all_searches_created = False
        
        if all_searches_created:
            self.log("All required searches exist and are ready for processing", Fore.GREEN)
        else:
            self.log("Some searches could not be created - processing may be incomplete", Fore.RED)
        
        return all_searches_created

    def get_search_status(self, search_id):
        response = self.make_api_request(
            "GET",
            f"https://graph.microsoft.com/beta/security/auditLog/queries/{search_id}"
        )
        
        if not response:
            return "failed"
            
        search_status = response.json()
        status = (search_status.get("status") or "").strip().lower()
        
        if status in ["succeeded", "failed", "cancelled"]:
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
                time.sleep(0.5)
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
        
        audit_data = {k: v for k, v in audit_data.items() if not k.endswith('@odata.type')}
        
        app_access_context = OrderedDict()
        if "AppAccessContext" in audit_data and isinstance(audit_data["AppAccessContext"], dict):
            app_access_context = OrderedDict(
                (k, v) for k, v in audit_data["AppAccessContext"].items() 
                if not k.endswith('@odata.type')
            )
            del audit_data["AppAccessContext"]
        
        formatted_data = OrderedDict()
        
        if app_access_context:
            formatted_data["AppAccessContext"] = app_access_context
        
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
        
        for field in standard_fields:
            if field in audit_data:
                formatted_data[field] = audit_data[field]
        
        for field, value in audit_data.items():
            if field not in formatted_data:
                formatted_data[field] = value
        
        return formatted_data

    def save_audit_to_csv(self, records, site):
        """Save audit records to a single CSV file per site"""
        site_name = site.split('/')[-1]
        start_day = self.START_DATE.split('T')[0].replace('-', '')
        end_day = self.END_DATE.split('T')[0].replace('-', '')
        current_time = datetime.now().astimezone().strftime('%Y%m%d_%H%M%S')
        filename = os.path.join(self.output_dir, f"{site_name}_{start_day}_{end_day}_{current_time}.csv")

        try:
            mode = CSV_REPORT_MODE.strip().upper()

            if not records:
                with open(filename, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.DictWriter(f, fieldnames=["Message"])
                    writer.writeheader()
                    writer.writerow({"Message": "No records found"})

                self.log(f"No records found for {site}. Created placeholder file {filename}", Fore.YELLOW)
                return 0

            normalized_records = []
            all_audit_fields = []
            all_audit_field_set = set()

            for record in records:
                audit_data = record.get("auditData", {})
                cleaned_data = self.clean_audit_data(audit_data)
                normalized_records.append({
                    "record": record,
                    "audit_data": cleaned_data
                })

                if mode == "ALL_FIELDS":
                    for field in cleaned_data.keys():
                        if field not in all_audit_field_set:
                            all_audit_field_set.add(field)
                            all_audit_fields.append(field)

            if mode == "RAW":
                fieldnames = ["id", "createdDateTime", "userPrincipalName", "operation", "auditData"]
            elif mode == "REQUIRED_FIELDS":
                fieldnames = [
                    "CreationTime",
                    "Id",
                    "userPrincipalName",
                    "ObjectId",
                    "Operation",
                    "ClientIP",
                    "ItemType"
                ]
            elif mode == "ALL_FIELDS":
                fieldnames = all_audit_fields
            else:
                raise ValueError(f"Invalid CSV_REPORT_MODE '{CSV_REPORT_MODE}'. Use RAW, REQUIRED_FIELDS, or ALL_FIELDS.")

            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                
                for entry in normalized_records:
                    record = entry["record"]
                    audit_data = entry["audit_data"]

                    if mode == "RAW":
                        row = {
                            "id": record.get("id"),
                            "createdDateTime": record.get("createdDateTime"),
                            "userPrincipalName": record.get("userPrincipalName"),
                            "operation": record.get("operation"),
                            "auditData": json.dumps(audit_data, ensure_ascii=False, indent=None)
                        }
                    elif mode == "REQUIRED_FIELDS":
                        row = {
                            "CreationTime": audit_data.get("CreationTime", ""),
                            "Id": audit_data.get("Id", ""),
                            "userPrincipalName": record.get("userPrincipalName"),
                            "ObjectId": audit_data.get("ObjectId", ""),
                            "Operation": audit_data.get("Operation", ""),
                            "ClientIP": audit_data.get("ClientIP", ""),
                            "ItemType": audit_data.get("ItemType", "")
                        }
                    else:
                        row = {
                            field: json.dumps(audit_data[field], ensure_ascii=False, indent=None)
                            if isinstance(audit_data.get(field), (dict, list))
                            else audit_data.get(field, "")
                            for field in fieldnames
                        }

                    writer.writerow(row)
            
            self.log(f"Saved {len(records)} records to {filename} using mode {mode}", Fore.GREEN)
            return len(records)
        except Exception as e:
            self.log(f"Failed to save CSV: {e}", Fore.RED, logging.ERROR)
            return 0

    def generate_summary(self):
        if not self.summary_data:
            self.log("No data available for summary", Fore.YELLOW)
            return

        current_time = datetime.now().astimezone().strftime('%Y%m%d_%H%M%S')
        summary_file = os.path.join(self.output_dir, f"AuditSummary_{current_time}.csv")
        
        try:
            with open(summary_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=["Site", "RecordCount", "StartDate", "EndDate"])
                writer.writeheader()
                for row in self.summary_data:
                    writer.writerow({
                        "Site": row["Site"],
                        "RecordCount": row["RecordCount"],
                        "StartDate": self.START_DATE,
                        "EndDate": self.END_DATE
                    })
            self.log(f"Summary file generated: {summary_file}", Fore.GREEN)
        except Exception as e:
            self.log(f"Failed to generate summary: {e}", Fore.RED, logging.ERROR)

    def get_site_library_info(self, site_url):
        """Resolve and cache the Graph site ID and Site Analytics drive ID for a site"""
        if site_url in self.site_library_cache:
            return self.site_library_cache[site_url]

        parsed = requests.utils.urlparse(site_url)
        hostname = parsed.netloc
        site_path = parsed.path.rstrip('/')

        site_response = self.make_api_request(
            "GET",
            f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
        )
        if not site_response:
            self.log(f"Failed to resolve Graph site for {site_url}", Fore.RED, logging.ERROR)
            return None

        site_id = site_response.json().get("id")
        if not site_id:
            self.log(f"Graph site ID missing for {site_url}", Fore.RED, logging.ERROR)
            return None

        drives_response = self.make_api_request(
            "GET",
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        )
        if not drives_response:
            self.log(f"Failed to get drives for {site_url}", Fore.RED, logging.ERROR)
            return None

        drives = drives_response.json().get("value", [])
        matching_drive = next(
            (drive for drive in drives if drive.get("name") == DOCUMENT_LIBRARY_NAME),
            None
        )
        if not matching_drive:
            self.log(f"Library '{DOCUMENT_LIBRARY_NAME}' not found for {site_url}", Fore.RED, logging.ERROR)
            return None

        library_info = {
            "site_id": site_id,
            "drive_id": matching_drive["id"],
            "drive_name": matching_drive["name"]
        }
        self.site_library_cache[site_url] = library_info
        self.log(f"Resolved {site_url} -> {DOCUMENT_LIBRARY_NAME} ({matching_drive['id']})", Fore.CYAN)
        return library_info

    def upload_csv_files(self):
        """Upload generated CSV files to SharePoint"""
        csv_files = glob.glob(os.path.join(self.output_dir, "*.csv"))
        csv_files = [f for f in csv_files if not os.path.basename(f).startswith("AuditSummary")]
        
        if not csv_files:
            self.log("No CSV files to upload", Fore.YELLOW)
            return
        
        self.log("Starting CSV file upload to SharePoint...", Fore.CYAN)
        
        for csv_file in csv_files:
            filename = os.path.basename(csv_file)
            site_name = filename.split('_')[0]
            site_url = next((site for site in SITES if site.endswith(f"/{site_name}")), None)
            
            if not site_url:
                self.log(f"Could not map CSV file {filename} to a site URL, skipping upload", Fore.YELLOW)
                continue
                
            try:
                self.log(f"Uploading {filename} to '{DOCUMENT_LIBRARY_NAME}' in {site_url}...", Fore.YELLOW)
                result = self.upload_file_to_sharepoint(csv_file, site_url)
                
                if result:
                    self.uploaded_files.append(csv_file)
                    self.log(f"Successfully uploaded {filename}", Fore.GREEN)
                    self.log(f"File name: {result.get('name')}", Fore.CYAN)
                    self.log(f"File size: {result.get('size')} bytes", Fore.CYAN)
                    self.log(f"Web URL: {result.get('webUrl')}", Fore.CYAN)
                    
            except Exception as e:
                self.log(f"Failed to upload {filename}: {str(e)}", Fore.RED, logging.ERROR)
        
        self.log("CSV file upload complete!", Fore.GREEN)

    def upload_file_to_sharepoint(self, file_path, site_url, folder_path=''):
        """Upload file to SharePoint using make_api_request for all operations"""
        library_info = self.get_site_library_info(site_url)
        if not library_info:
            raise Exception(f"Could not find '{DOCUMENT_LIBRARY_NAME}' library for {site_url}")

        file_size = os.path.getsize(file_path)
        file_name = os.path.basename(file_path)
        
        if file_size <= 4194304:
            result = self.simple_upload(file_path, library_info["site_id"], library_info["drive_id"], folder_path)
        else:
            result = self.upload_large_file(file_path, library_info["site_id"], library_info["drive_id"], folder_path)
        
        if result and 'id' in result:
            self.update_file_properties(library_info["site_id"], library_info["drive_id"], result['id'], file_name)
        return result

    def simple_upload(self, file_path, site_id, drive_id, folder_path):
        """Upload small files (<4MB) directly using make_api_request"""
        file_name = os.path.basename(file_path)
        
        if folder_path:
            upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path}/{file_name}:/content"
        else:
            upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{file_name}:/content"
        
        with open(file_path, 'rb') as file:
            file_content = file.read()
        
        response = self.make_api_request(
            "PUT",
            upload_url,
            data=file_content,
            headers={
                "Authorization": f"Bearer {self.access_token}",
                "Content-Type": "application/octet-stream"
            }
        )
        
        if not response:
            raise Exception(f"Failed to upload {file_name}")
            
        if response.status_code in (200, 201):
            self.log(f"Successfully uploaded {file_name}", Fore.GREEN)
            return response.json()
        else:
            raise Exception(f"Failed to upload {file_name}. Status: {response.status_code}\n{response.text}")

    def upload_large_file(self, file_path, site_id, drive_id, folder_path):
        """Upload large files using upload session with chunking via make_api_request"""
        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)
        
        if folder_path:
            create_session_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{folder_path}/{file_name}:/createUploadSession"
        else:
            create_session_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{file_name}:/createUploadSession"
        
        session_body = {
            "item": {
                "@microsoft.graph.conflictBehavior": "replace",
                "name": file_name
            }
        }
        
        response = self.make_api_request("POST", create_session_url, json_data=session_body)
        
        if not response:
            raise Exception("Failed to create upload session")
            
        upload_url = response.json()['uploadUrl']
        
        chunk_size = CHUNK_SIZE
        total_chunks = math.ceil(file_size / chunk_size)
        
        with open(file_path, 'rb') as file:
            for chunk_num in range(total_chunks):
                start = chunk_num * chunk_size
                end = min(start + chunk_size, file_size)
                chunk_size_actual = end - start
                
                headers = {
                    'Content-Length': str(chunk_size_actual),
                    'Content-Range': f'bytes {start}-{end-1}/{file_size}',
                    'Authorization': f'Bearer {self.access_token}'
                }
                
                file.seek(start)
                chunk_data = file.read(chunk_size_actual)
                
                response = self.make_api_request("PUT", upload_url, data=chunk_data, headers=headers)
                
                if not response:
                    raise Exception(f"Upload failed at chunk {chunk_num+1}/{total_chunks}")
                
                if response.status_code not in (200, 201, 202):
                    raise Exception(f"Upload failed at chunk {chunk_num+1}/{total_chunks}: {response.text}")
                
                self.log(f"Uploaded chunk {chunk_num+1}/{total_chunks} ({end/file_size:.1%})", Fore.YELLOW)
        
        return response.json()

    def update_file_properties(self, site_id, drive_id, file_item_id, file_name):
        """Update file properties (metadata) in SharePoint using make_api_request"""
        list_item_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_item_id}/listItem"
        fields_url = f"{list_item_url}/fields"
        
        update_payload = {
            "Month": self.REPORT_MONTH,
            "Year": self.REPORT_YEAR
        }
        
        update_response = self.make_api_request("PATCH", fields_url, json_data=update_payload)
        
        if not update_response:
            self.log("\nFailed to update file properties", Fore.RED)
            return None
            
        verify_response = self.make_api_request("GET", fields_url)
        if not verify_response:
            self.log("\nFailed to verify property updates", Fore.YELLOW)
            return None
            
        verify_data = verify_response.json()
        
        self.log(f"\nProperty update results for file {file_name}:", Fore.CYAN)
        self.log(f"Month: {verify_data.get('Month', 'Not updated')}", Fore.CYAN)
        self.log(f"Year: {verify_data.get('Year', 'Not updated')}", Fore.CYAN)
        
        return update_response.json()

    def load_state(self, filename):
        """Load state from fixed state directory"""
        state_file = os.path.join(self.state_dir, filename)
        if os.path.exists(state_file):
            try:
                with open(state_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                self.log(f"Failed to load state file {filename}: {e}", Fore.YELLOW)
                return {}
        return {}

    def save_state(self, filename, data):
        """Save state to fixed state directory"""
        state_file = os.path.join(self.state_dir, filename)
        try:
            with open(state_file, 'w') as f:
                json.dump(data, f, indent=4)
            self.log(f"Saved state to {state_file}", Fore.CYAN)
        except Exception as e:
            self.log(f"Failed to save state file {filename}: {e}", Fore.RED, logging.ERROR)

    def archive_search_ids(self):
        """Archive search_ids.json after completion"""
        state_file = os.path.join(self.state_dir, "search_ids.json")
        if os.path.exists(state_file):
            timestamp = datetime.now().astimezone().strftime('%Y%m%d_%H%M%S')
            archive_dir = os.path.join(self.state_dir, "archive")
            os.makedirs(archive_dir, exist_ok=True)
            new_name = os.path.join(archive_dir, f"searchIds_{timestamp}.json")
            try:
                os.rename(state_file, new_name)
                self.log(f"Archived search_ids.json as {new_name}", Fore.GREEN)
                return True
            except Exception as e:
                self.log(f"Failed to archive search_ids.json: {e}", Fore.RED, logging.ERROR)
                return False
        return True

    def process_pending_searches(self):
        """Process all searches with verification step"""
        if not self.verify_all_searches_created():
            self.log("Cannot proceed with processing due to missing searches", Fore.RED)
            return

        total_searches = len(SITES)
        terminal_success_statuses = {"succeeded", "completed", "partiallysucceeded", "noresults", "nodata"}
        terminal_failure_statuses = {"failed", "cancelled", "canceled", "timeout"}

        while len(self.completed_searches) < total_searches:
            self.log(f"Processing searches ({len(self.completed_searches)}/{total_searches} completed)...", Fore.YELLOW)
            
            for site, search_data in self.existing_searches.items():
                if site in self.completed_searches:
                    continue
                    
                status = self.get_search_status(search_data["SearchId"])
                self.search_poll_counts[site] = self.search_poll_counts.get(site, 0) + 1

                if status in terminal_success_statuses:
                    self.log(f"Retrieving records for {site}", Fore.CYAN)
                    records = self.get_audit_records(search_data["SearchId"])
                    
                    record_count = self.save_audit_to_csv(records, site)
                    if record_count == 0:
                        self.log(f"No records returned for {site}. Marking search as completed.", Fore.YELLOW)

                    self.summary_data.append({
                        "Site": site.split('/')[-1],
                        "RecordCount": record_count
                    })
                    
                    self.completed_searches[site] = {
                        "CompletedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
                        "RecordCount": record_count
                    }
                    self.save_state("completed_searches.json", self.completed_searches)
                    
                elif status in terminal_failure_statuses:
                    self.completed_searches[site] = {
                        "CompletedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
                        "RecordCount": 0,
                        "Status": status
                    }
                    self.save_state("completed_searches.json", self.completed_searches)
                elif self.search_poll_counts[site] >= MAX_SEARCH_STATUS_POLLS:
                    self.log(f"Search {search_data['SearchId']} exceeded max polling attempts with status '{status or 'unknown'}'. Marking as timeout.", Fore.YELLOW)
                    self.completed_searches[site] = {
                        "CompletedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
                        "RecordCount": 0,
                        "Status": "timeout"
                    }
                    self.save_state("completed_searches.json", self.completed_searches)
            
            if len(self.completed_searches) < total_searches:
                time.sleep(RETRY_DELAY_SECONDS)

        self.log(f"Completed processing all {total_searches} searches", Fore.GREEN)

    def run(self):
        try:
            self.process_start_time = datetime.now().astimezone()
            
            self.log("Starting audit log collection process...", Fore.CYAN)
            self.log(f"Date range: {self.START_DATE} to {self.END_DATE}", Fore.CYAN)
            self.log(f"Report Month/Year: {self.REPORT_MONTH} {self.REPORT_YEAR}", Fore.CYAN)
            
            self.log("Authenticating...", Fore.YELLOW)
            if not self.get_access_token():
                self.log("Failed to authenticate, exiting...", Fore.RED)
                return
            
            # Load state from fixed directory
            self.existing_searches = self.load_state("search_ids.json")
            self.completed_searches = self.load_state("completed_searches.json")
            self.log(f"Loaded {len(self.existing_searches)} existing searches and {len(self.completed_searches)} completed searches", Fore.YELLOW)
            
            self.process_pending_searches()
            self.generate_summary()
            
            self.log("Authenticating for SharePoint upload...", Fore.YELLOW)
            if not self.get_access_token():
                self.log("Failed to authenticate, skipping upload...", Fore.RED)
            else:
                self.upload_csv_files()
            
            # Clean up state files if all searches completed
            if len(self.completed_searches) == len(SITES):
                if self.archive_search_ids():
                    completed_file = os.path.join(self.state_dir, "completed_searches.json")
                    if os.path.exists(completed_file):
                        try:
                            os.remove(completed_file)
                            self.log("Removed completed_searches.json", Fore.GREEN)
                        except Exception as e:
                            self.log(f"Failed to remove completed_searches.json: {e}", Fore.YELLOW)
            
            self.process_end_time = datetime.now().astimezone()
            self.log("All operations completed successfully!", Fore.GREEN)

        except Exception as e:
            self.log(f"Fatal error: {e}", Fore.RED)
            raise

if __name__ == "__main__":
    collector = AuditLogCollector()
    collector.run()
