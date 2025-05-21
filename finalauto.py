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
import math
from requests_toolbelt import MultipartEncoder
import logging
from logging.handlers import RotatingFileHandler

# Initialize colorama
init(autoreset=True)

# Configuration
TENANT_ID = "0e439a1f-a497-462b-9e6b-4e582e203607"
CLIENT_ID = "73efa35d-6188-42d4-b258-838a977eb149"
CLIENT_SECRET = "CyG8Q~FYHuCMSyVmt4sNxt5IejrMc2c24Ziz4a.t"

# SharePoint Upload Configuration
SITE_ID = 'cf2d90f2-6092-4b2b-a89e-1cf4981c633d'
DRIVE_ID = 'b!8pAtz5JgK0uonhz0mBxjPemrblfYP1VLrqWQcFv674DXWcxcOkJ8SpxymxPTeiNI'
CHUNK_SIZE = 3276800  # 3.125 MB

# Debug mode - SET TO FALSE FOR PRODUCTION!
DEBUG_MODE = False  # Shows full tokens when True

# Sites and operations
SITES = [
    "https://geekbyteonline.sharepoint.com/sites/2DayRetention",
    "https://geekbyteonline.sharepoint.com/sites/geekbyteteam",
    "https://geekbyteonline.sharepoint.com/sites/Radiation-Imaging",
    "https://geekbyteonline.sharepoint.com/sites/geetkteam",
    "https://geekbyteonline.sharepoint.com/sites/New365Site5"
]

# Mapping of site names to resource names
SITE_TO_RESOURCE = {
    "2DayRetention": "2-DayRetention",
    "Geekbyteteam": "Pharmacy",
    "Radiation-Imaging": "Radiation Imaging",
    "geetkteam": "Geetk Team",
    "New365Site5": "New 365 Site 5"
}

SITE_NAMES = [site.split('/')[-1] for site in SITES]  # For Excel report generation
OPERATIONS = ["PageViewed", "FileAccessed", "FileDownloaded"]

# Constants
MAX_RETRIES = 10
RETRY_DELAY_SECONDS = 300  # Default wait time for rate limits

class AuditLogCollector:
    def __init__(self):
        # Create output directory for this run
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
        self.log(f"Output directory: {os.path.abspath(self.output_dir)}", Fore.CYAN)
    
    def set_date_ranges(self):
        """Set the date ranges for the report"""
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
        self.REPORT_MONTH = self.start_date.strftime("%B")  # Full month name e.g. "May"
        self.REPORT_YEAR = self.start_date.strftime("%Y")   # e.g. "2025"
        
        self.log(f"Date range set to: {self.START_DATE} to {self.END_DATE}", Fore.CYAN)
        self.log(f"Report Month/Year: {self.REPORT_MONTH} {self.REPORT_YEAR}", Fore.CYAN)
    
    def log(self, message, color=Fore.WHITE, level=logging.INFO):
        """Log message to both console (with color) and log file"""
        # Get local time with timezone
        local_time = datetime.now().astimezone()
        console_message = f"{color}[{local_time.strftime('%Y-%m-%d %H:%M:%S')}] {message}{Style.RESET_ALL}"
        print(console_message)
        
        # Log to file without color codes
        self.logger.log(level, message)
    
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
                now = datetime.now().astimezone()
                token_time = datetime.fromisoformat(
                    self.token_generation_time.replace('Z', '+00:00')).astimezone()
                age = now - token_time
                self.log(f"Token generated at: {token_time.strftime('%Y-%m-%d %H:%M:%S')} (age: {age.total_seconds():.0f} seconds)", Fore.CYAN)
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
            self.log(f"Failed to get access token: {e}", Fore.RED, logging.ERROR)
            return False

    def check_existing_search(self, site, operation):
        """Check if a search already exists with the same parameters"""
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }
        
        try:
            # Get all existing searches
            response = requests.get(
                "https://graph.microsoft.com/beta/security/auditLog/queries",
                headers=headers
            )
            response.raise_for_status()
            
            searches = response.json().get('value', [])
            
            # Check each search for matching criteria
            for search in searches:
                if (search.get('filterStartDateTime') == self.START_DATE and
                    search.get('filterEndDateTime') == self.END_DATE and
                    operation in search.get('operationFilters', []) and
                    f"{site}/*" in search.get('objectIdFilters', [])):
                    
                    self.log(f"Found existing search with matching criteria: {search['id']}", Fore.CYAN)
                    return search['id']
                    
            return None
            
        except requests.exceptions.RequestException as e:
            self.log(f"Error checking existing searches: {e}", Fore.YELLOW)
            return None

    def new_audit_search(self, site, operation):
        """Create a new audit log search after checking for existing ones"""
        # First check if a matching search already exists
        existing_search_id = self.check_existing_search(site, operation)
        if existing_search_id:
            return existing_search_id
        
        # If no existing search found, create a new one
        search_params = {
            "displayName": f"Audit_{site.split('/')[-1]}_{operation}_{datetime.now().astimezone().strftime('%Y%m%d_%H%M%S')}",
            "filterStartDateTime": self.START_DATE,
            "filterEndDateTime": self.END_DATE,
            "operationFilters": [operation],
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
        self.log(f"Search created for {site} - {operation}", Fore.GREEN)
        return search_data.get("id")

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
        
        self.log(f"Max retries ({max_retries}) exceeded for this request", Fore.RED, logging.ERROR)
        return None

    def verify_all_searches_created(self):
        """Verify that all required searches exist before processing"""
        self.log("Verifying all required searches exist...", Fore.CYAN)
        all_searches_created = True
        
        for site in SITES:
            for operation in OPERATIONS:
                key = f"{site}_{operation}"
                if key not in self.existing_searches:
                    self.log(f"Missing search for {site} - {operation}", Fore.YELLOW)
                    all_searches_created = False
                    # Attempt to create the missing search
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
                        self.log(f"Created/Found search for {site} - {operation}", Fore.GREEN)
                        all_searches_created = True
                    else:
                        self.log(f"Failed to create/find search for {site} - {operation}", Fore.RED)
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
        start_day = self.START_DATE.split('T')[0].replace('-', '')
        end_day = self.END_DATE.split('T')[0].replace('-', '')
        current_time = datetime.now().astimezone().strftime('%Y%m%d_%H%M%S')
        filename = os.path.join(self.output_dir, f"{site_name}_{start_day}_{end_day}_{operation}_{current_time}.csv")

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
                writer = csv.DictWriter(f, fieldnames=["Site", "Operation", "RecordCount", "StartDate", "EndDate"])
                writer.writeheader()
                for row in self.summary_data:
                    writer.writerow({
                        "Site": row["Site"],
                        "Operation": row["Operation"],
                        "RecordCount": row["RecordCount"],
                        "StartDate": self.START_DATE,
                        "EndDate": self.END_DATE
                    })
            self.log(f"Summary file generated: {summary_file}", Fore.GREEN)
        except Exception as e:
            self.log(f"Failed to generate summary: {e}", Fore.RED, logging.ERROR)

    def format_resource_name(self, site_name):
        """Convert site name to Resource format using predefined mapping"""
        return SITE_TO_RESOURCE.get(site_name, site_name.replace('-', ' ').title())

    def generate_excel_reports(self):
        """Generate Excel reports from the CSV files for each site"""
        self.log("Starting Excel report generation...", Fore.CYAN)
        
        excel_files = []  # Track generated Excel files for upload
        
        for site_name in SITE_NAMES:
            # Create Excel filename with full month name
            excel_file = os.path.join(self.output_dir, f"{site_name} {self.REPORT_MONTH} {self.REPORT_YEAR}.xlsx")
            
            # Create Excel file with separate sheets
            with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                sheets_created = 0
                
                for operation in OPERATIONS:
                    # Find matching CSV file
                    pattern = os.path.join(self.output_dir, f"{site_name}_*_{operation}_*.csv")
                    matches = glob.glob(pattern)
                    
                    if not matches:
                        self.log(f"Warning: No file found for {site_name} {operation}", Fore.YELLOW)
                        continue
                    
                    csv_file = matches[0]  # Take the first match
                    
                    try:
                        # Read CSV file
                        df = pd.read_csv(csv_file)
                        
                        # Write to Excel with operation as sheet name
                        df.to_excel(writer, sheet_name=operation, index=False)
                        sheets_created += 1
                        
                        self.log(f"Added sheet: {operation} (from {os.path.basename(csv_file)})", Fore.GREEN)
                    except Exception as e:
                        self.log(f"Error processing {csv_file}: {str(e)}", Fore.RED, logging.ERROR)
            
                if sheets_created > 0:
                    self.log(f"\n✅ Created {excel_file} with {sheets_created} sheets\n", Fore.GREEN)
                    excel_files.append(excel_file)
                else:
                    self.log(f"\n⚠️ No sheets created for {site_name}\n", Fore.YELLOW)
                    try:
                        os.remove(excel_file)  # Remove empty file
                    except Exception as e:
                        self.log(f"Failed to remove empty Excel file: {e}", Fore.YELLOW)

        self.log("Excel report generation complete!", Fore.GREEN)
        return excel_files

    def upload_file_to_sharepoint(self, file_path, folder_path=''):
        """Upload file to SharePoint with automatic handling of large files and token refresh"""
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                file_size = os.path.getsize(file_path)
                file_name = os.path.basename(file_path)
                
                # Simple upload for files < 4MB
                if file_size <= 4 * 1024 * 1024:
                    result = self.simple_upload(file_path, folder_path)
                else:
                    # Upload session for large files
                    result = self.upload_large_file(file_path, folder_path)
                
                # Update file properties after upload if successful
                if result and 'id' in result:
                    self.update_file_properties(result['id'], file_name)
                return result
                
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 401 and retry_count < max_retries:
                    self.log("Token expired during upload, refreshing and retrying...", Fore.YELLOW)
                    if self.get_access_token():
                        retry_count += 1
                        continue
                    else:
                        self.log("Failed to refresh token during upload", Fore.RED)
                        raise
                else:
                    raise
            except Exception as e:
                self.log(f"Error during upload: {str(e)}", Fore.RED)
                raise
        
        raise Exception(f"Failed to upload after {max_retries} retries")

    def simple_upload(self, file_path, folder_path):
        """Upload small files (<4MB) directly with token refresh handling"""
        file_name = os.path.basename(file_path)
        
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/octet-stream'
        }
        
        if folder_path:
            upload_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{folder_path}/{file_name}:/content"
        else:
            upload_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{file_name}:/content"
        
        with open(file_path, 'rb') as file:
            file_content = file.read()
        
        response = requests.put(upload_url, headers=headers, data=file_content)
        
        if response.status_code in (200, 201):
            self.log(f"Successfully uploaded {file_name}", Fore.GREEN)
            return response.json()
        else:
            raise Exception(f"Failed to upload {file_name}. Status: {response.status_code}\n{response.text}")

    def upload_large_file(self, file_path, folder_path):
        """Upload large files using upload session with chunking and token refresh handling"""
        file_name = os.path.basename(file_path)
        file_size = os.path.getsize(file_path)
        
        # 1. Create upload session
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
        
        if folder_path:
            create_session_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{folder_path}/{file_name}:/createUploadSession"
        else:
            create_session_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root:/{file_name}:/createUploadSession"
        
        session_body = {
            "item": {
                "@microsoft.graph.conflictBehavior": "replace",
                "name": file_name
            }
        }
        
        # Handle token refresh for session creation
        try:
            response = requests.post(create_session_url, headers=headers, json=session_body)
            if response.status_code == 401:
                self.log("Token expired during session creation, refreshing...", Fore.YELLOW)
                if self.get_access_token():
                    headers['Authorization'] = f'Bearer {self.access_token}'
                    response = requests.post(create_session_url, headers=headers, json=session_body)
            
            response.raise_for_status()
        except Exception as e:
            self.log(f"Failed to create upload session: {str(e)}", Fore.RED)
            raise
        
        upload_url = response.json()['uploadUrl']
        
        # 2. Upload file in chunks
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
                
                try:
                    response = requests.put(upload_url, headers=headers, data=chunk_data)
                    
                    if response.status_code == 401:
                        self.log("Token expired during chunk upload, refreshing...", Fore.YELLOW)
                        if self.get_access_token():
                            headers['Authorization'] = f'Bearer {self.access_token}'
                            response = requests.put(upload_url, headers=headers, data=chunk_data)
                    
                    if response.status_code not in (200, 201, 202):
                        raise Exception(f"Upload failed at chunk {chunk_num+1}/{total_chunks}: {response.text}")
                    
                    self.log(f"Uploaded chunk {chunk_num+1}/{total_chunks} ({end/file_size:.1%})", Fore.YELLOW)
                except Exception as e:
                    self.log(f"Error uploading chunk {chunk_num+1}: {str(e)}", Fore.RED)
                    raise
        
        return response.json()

    def update_file_properties(self, file_item_id, file_name):
        """Update file properties (metadata) in SharePoint with token refresh handling"""
        max_retries = 3
        retry_count = 0
        
        while retry_count < max_retries:
            try:
                # Extract site name from file name (format: "sitename Month Year.xlsx")
                site_name = file_name.split()[0]
                resource_name = self.format_resource_name(site_name)
                
                # Get the list item associated with the file
                list_item_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/items/{file_item_id}/listItem"
                fields_url = f"{list_item_url}/fields"
                
                headers = {
                    'Authorization': f'Bearer {self.access_token}',
                    'Content-Type': 'application/json'
                }
                
                # Update properties including new Resource field
                update_payload = {
                    "Month": self.REPORT_MONTH,
                    "Year": self.REPORT_YEAR,
                    "Resource": resource_name
                }
                
                # Make the PATCH request
                update_response = requests.patch(fields_url, headers=headers, json=update_payload)
                
                if update_response.status_code == 401:
                    raise requests.exceptions.HTTPError("Unauthorized", response=update_response)
                    
                update_response.raise_for_status()
                
                # Verify the update
                verify_response = requests.get(fields_url, headers=headers)
                verify_data = verify_response.json()
                
                self.log(f"\nProperty update results for file {file_name}:", Fore.CYAN)
                self.log(f"Month: {verify_data.get('Month', 'Not updated')}", Fore.CYAN)
                self.log(f"Year: {verify_data.get('Year', 'Not updated')}", Fore.CYAN)
                self.log(f"Resource: {verify_data.get('Resource', 'Not updated')}", Fore.CYAN)
                
                return update_response.json()
                
            except requests.exceptions.HTTPError as e:
                if e.response.status_code == 401 and retry_count < max_retries:
                    self.log("Token expired during property update, refreshing...", Fore.YELLOW)
                    if self.get_access_token():
                        retry_count += 1
                        continue
                    else:
                        self.log("Failed to refresh token during property update", Fore.RED)
                        raise
                else:
                    self.log(f"\nDetailed error updating properties:", Fore.RED)
                    self.log(f"Status code: {e.response.status_code}", Fore.RED)
                    self.log(f"Error response: {e.response.text}", Fore.RED)
                    
                    # Check if this is a column type issue
                    if "Invalid request" in e.response.text and "field types" in e.response.text:
                        self.log("\nPossible column type mismatch. Please ensure:", Fore.YELLOW)
                        self.log("1. The 'Month', 'Year', and 'Resource' columns exist in the document library", Fore.YELLOW)
                        self.log("2. The columns are of type 'Choice' or 'Text'", Fore.YELLOW)
                        self.log("3. The values match the column's expected format", Fore.YELLOW)
                    raise
            except Exception as e:
                self.log(f"\nError updating properties: {str(e)}", Fore.RED)
                raise
        
        self.log("\nThe file was uploaded successfully, but properties couldn't be updated.", Fore.YELLOW)
        return None

    def upload_excel_reports(self, excel_files):
        """Upload generated Excel reports to SharePoint"""
        if not excel_files:
            self.log("No Excel files to upload", Fore.YELLOW)
            return
        
        self.log("Starting Excel report upload to SharePoint...", Fore.CYAN)
        
        for excel_file in excel_files:
            try:
                self.log(f"Uploading {excel_file} to SharePoint...", Fore.YELLOW)
                result = self.upload_file_to_sharepoint(excel_file)
                
                if result:
                    self.log(f"Successfully uploaded {excel_file}", Fore.GREEN)
                    self.log(f"File name: {result.get('name')}", Fore.CYAN)
                    self.log(f"File size: {result.get('size')} bytes", Fore.CYAN)
                    self.log(f"Web URL: {result.get('webUrl')}", Fore.CYAN)

                      # TESTING: Add 60-minute delay after each file upload
                    self.log("TEST MODE: Sleeping for 60 minutes after upload...", Fore.RED)
                    time.sleep(4000)  # 60 minutes = 3600 seconds
                    
                    # # Optionally remove the local file after successful upload
                    # try:
                        # os.remove(excel_file)
                        # self.log(f"Removed local file: {excel_file}", Fore.YELLOW)
                    # except Exception as e:
                        # self.log(f"Failed to remove local file: {e}", Fore.YELLOW)
            except Exception as e:
                self.log(f"Failed to upload {excel_file}: {str(e)}", Fore.RED, logging.ERROR)
        
        self.log("Excel report upload complete!", Fore.GREEN)

    def load_state(self, filename):
        state_file = os.path.join(self.output_dir, filename)
        if os.path.exists(state_file):
            try:
                with open(state_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                self.log(f"Failed to load state file {filename}: {e}", Fore.YELLOW)
                return {}
        return {}

    def save_state(self, filename, data):
        state_file = os.path.join(self.output_dir, filename)
        try:
            with open(state_file, 'w') as f:
                json.dump(data, f, indent=4)
            self.log(f"Saved state to {filename}", Fore.CYAN)
        except Exception as e:
            self.log(f"Failed to save state file {filename}: {e}", Fore.RED, logging.ERROR)

    def archive_search_ids(self):
        state_file = os.path.join(self.output_dir, "search_ids.json")
        if os.path.exists(state_file):
            timestamp = datetime.now().astimezone().strftime('%Y%m%d_%H%M%S')
            new_name = os.path.join(self.output_dir, f"searchIds_{timestamp}.json")
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
        # First verify and create any missing searches
        if not self.verify_all_searches_created():
            self.log("Cannot proceed with processing due to missing searches", Fore.RED)
            return

        total_searches = len(SITES) * len(OPERATIONS)
        processed_searches = 0

        # Then process all searches (existing and new)
        while len(self.completed_searches) < total_searches:
            self.log(f"Processing searches ({len(self.completed_searches)}/{total_searches} completed)...", Fore.YELLOW)
            
            for key, search_data in self.existing_searches.items():
                if key in self.completed_searches:
                    processed_searches += 1
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
                    processed_searches += 1
                    
                elif status == "failed":
                    self.completed_searches[key] = {
                        "CompletedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
                        "RecordCount": 0,
                        "Status": "failed"
                    }
                    self.save_state("completed_searches.json", self.completed_searches)
                    processed_searches += 1
            
            if len(self.completed_searches) < total_searches:
                time.sleep(RETRY_DELAY_SECONDS)  # Reduced wait time between checks

        self.log(f"Completed processing all {processed_searches} searches", Fore.GREEN)

    def run(self):
        try:
            self.log("Starting audit log collection process...", Fore.CYAN)
            self.log(f"Date range: {self.START_DATE} to {self.END_DATE}", Fore.CYAN)
            self.log(f"Report Month/Year: {self.REPORT_MONTH} {self.REPORT_YEAR}", Fore.CYAN)
            
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
            excel_files = self.generate_excel_reports()
            
            # Upload Excel reports to SharePoint
            self.upload_excel_reports(excel_files)
            
            # Archive search IDs
            if len(self.completed_searches) == len(SITES) * len(OPERATIONS):
                if self.archive_search_ids():
                    # Only remove if archiving was successful
                    completed_file = os.path.join(self.output_dir, "completed_searches.json")
                    if os.path.exists(completed_file):
                        try:
                            os.remove(completed_file)
                        except Exception as e:
                            self.log(f"Failed to remove completed_searches.json: {e}", Fore.YELLOW)
            
            self.log("All operations completed successfully!", Fore.GREEN)
                
        except Exception as e:
            self.log(f"Fatal error: {e}", Fore.RED)
            raise

if __name__ == "__main__":
    collector = AuditLogCollector()
    collector.run()