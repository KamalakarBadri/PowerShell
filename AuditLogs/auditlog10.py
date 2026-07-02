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
DEBUG_MODE = False

# Date range configuration
USE_MANUAL_DATE_RANGE = False
MANUAL_START_DATE = "2026-02-01T04:00:00Z"
MANUAL_END_DATE = "2026-03-01T03:59:59Z"

# CSV report mode
CSV_REPORT_MODE = "RAW"

# Sites
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
        # SINGLE output directory (reuse same folder)
        self.base_dir = "AuditLogs"
        os.makedirs(self.base_dir, exist_ok=True)
        
        # Use the same directory always
        self.output_dir = self.base_dir
        
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
        
        # Load state immediately
        self.load_all_state()
        
        # Get date ranges
        self.set_date_ranges()
        
    def setup_logging(self):
        """Configure logging to both console and file"""
        self.logger = logging.getLogger('AuditLogCollector')
        self.logger.setLevel(logging.DEBUG)
        
        formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
        
        # Console handler
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(formatter)
        self.logger.addHandler(ch)
        
        # File handler
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
        
        self.log("Initialized Audit Log Collector", Fore.CYAN)
        self.log(f"Output directory: {os.path.abspath(self.output_dir)}", Fore.CYAN)
    
    def load_all_state(self):
        """Load all state files from the output directory"""
        # Load search IDs
        search_file = os.path.join(self.output_dir, "search_ids.json")
        if os.path.exists(search_file):
            try:
                with open(search_file, 'r') as f:
                    self.existing_searches = json.load(f)
                self.log(f"Loaded {len(self.existing_searches)} existing searches from state", Fore.CYAN)
            except Exception as e:
                self.log(f"Failed to load search_ids.json: {e}", Fore.YELLOW)
                self.existing_searches = {}
        
        # Load completed searches
        completed_file = os.path.join(self.output_dir, "completed_searches.json")
        if os.path.exists(completed_file):
            try:
                with open(completed_file, 'r') as f:
                    self.completed_searches = json.load(f)
                self.log(f"Loaded {len(self.completed_searches)} completed searches from state", Fore.CYAN)
            except Exception as e:
                self.log(f"Failed to load completed_searches.json: {e}", Fore.YELLOW)
                self.completed_searches = {}
    
    def save_search_state(self):
        """Save search IDs to state file"""
        search_file = os.path.join(self.output_dir, "search_ids.json")
        try:
            with open(search_file, 'w') as f:
                json.dump(self.existing_searches, f, indent=4)
            self.log(f"Saved search state to {search_file}", Fore.CYAN)
        except Exception as e:
            self.log(f"Failed to save search state: {e}", Fore.RED, logging.ERROR)
    
    def save_completed_state(self):
        """Save completed searches to state file"""
        completed_file = os.path.join(self.output_dir, "completed_searches.json")
        try:
            with open(completed_file, 'w') as f:
                json.dump(self.completed_searches, f, indent=4)
            self.log(f"Saved completed state to {completed_file}", Fore.CYAN)
        except Exception as e:
            self.log(f"Failed to save completed state: {e}", Fore.RED, logging.ERROR)
    
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
            current_date = datetime.now().astimezone()
            self.start_date = (current_date - relativedelta(months=1)).replace(
                day=1, hour=4, minute=0, second=0, microsecond=0, tzinfo=timezone.utc
            )
            self.end_date = current_date.replace(
                day=1, hour=3, minute=59, second=59, microsecond=0, tzinfo=timezone.utc
            )

        self.START_DATE = self.start_date.isoformat().replace('+00:00', 'Z')
        self.END_DATE = self.end_date.isoformat().replace('+00:00', 'Z')
        self.REPORT_MONTH = self.start_date.strftime("%B")
        self.REPORT_YEAR = self.start_date.strftime("%Y")
        
        self.log(f"Date range: {self.START_DATE} to {self.END_DATE}", Fore.CYAN)
        self.log(f"Report Month/Year: {self.REPORT_MONTH} {self.REPORT_YEAR}", Fore.CYAN)
    
    def log(self, message, color=Fore.WHITE, level=logging.INFO):
        """Log message to both console and file"""
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
                self.log(f"Token age: {age.total_seconds():.0f} seconds", Fore.CYAN)
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
                    self.log(f"Rate limited (HTTP {response.status_code}), waiting {retry_after}s...", Fore.YELLOW)
                    time.sleep(retry_after)
                    retry_count += 1
                    continue
                
                response.raise_for_status()
                return response
            
            except requests.exceptions.ConnectionError:
                self.log(f"Connection error, retrying... ({retry_count + 1}/{max_retries})", Fore.YELLOW)
                retry_count += 1
                time.sleep(RETRY_DELAY_SECONDS)
                continue
                
            except requests.exceptions.Timeout:
                self.log(f"Request timeout, retrying... ({retry_count + 1}/{max_retries})", Fore.YELLOW)
                retry_count += 1
                time.sleep(RETRY_DELAY_SECONDS)
                continue
                
            except requests.exceptions.RequestException as e:
                self.log(f"Request failed: {str(e)} - retrying... ({retry_count + 1}/{max_retries})", Fore.YELLOW)
                retry_count += 1
                time.sleep(RETRY_DELAY_SECONDS)
                continue
        
        self.log(f"Max retries ({max_retries}) exceeded", Fore.RED, logging.ERROR)
        return None
    
    def find_existing_search_in_graph(self, site):
        """Find the BEST existing search in Microsoft Graph - prioritize completed ones"""
        self.log(f"🔍 Checking for existing searches for {site}...", Fore.YELLOW)
        
        response = self.make_api_request(
            "GET",
            "https://graph.microsoft.com/beta/security/auditLog/queries"
        )
        
        if not response:
            self.log("Failed to get list of searches from Graph", Fore.RED)
            return None
            
        all_searches = response.json().get('value', [])
        self.log(f"Found {len(all_searches)} total searches in Graph", Fore.CYAN)
        
        # Filter searches that match our criteria
        matching_searches = []
        for search in all_searches:
            search_start = search.get('filterStartDateTime', '')
            search_end = search.get('filterEndDateTime', '')
            object_filters = search.get('objectIdFilters', [])
            
            # Skip if date range doesn't match
            if search_start != self.START_DATE or search_end != self.END_DATE:
                continue
            
            # Check if this search matches our site
            site_match = False
            clean_site = site.rstrip('/')
            
            for filter_item in object_filters:
                # Remove trailing /* for comparison
                clean_filter = filter_item.rstrip('/*')
                
                # Check various matching patterns
                if filter_item == "*":
                    site_match = True  # Wildcard matches everything
                    break
                elif clean_filter == clean_site:
                    site_match = True  # Exact site match without /*
                    break
                elif filter_item == f"{clean_site}/*":
                    site_match = True  # Exact match with /*
                    break
                elif clean_site in filter_item:
                    site_match = True  # Partial match (site is contained)
                    break
            
            if site_match:
                matching_searches.append({
                    'id': search.get('id'),
                    'status': search.get('status', 'unknown'),
                    'displayName': search.get('displayName', ''),
                    'createdDateTime': search.get('createdDateTime', ''),
                    'search': search,
                    'object_filters': object_filters,
                    'search_start': search_start,
                    'search_end': search_end
                })
        
        if not matching_searches:
            self.log(f"No existing search found for {site} with this date range", Fore.YELLOW)
            return None
        
        # Log all matching searches with details
        self.log(f"Found {len(matching_searches)} matching searches for {site}:", Fore.CYAN)
        for idx, s in enumerate(matching_searches, 1):
            status_color = Fore.GREEN if s['status'] in ['succeeded', 'completed'] else Fore.YELLOW
            self.log(f"  {idx}. ID: {s['id'][:20]}... Status: {status_color}{s['status']}{Style.RESET_ALL}", Fore.CYAN)
            self.log(f"     Created: {s['createdDateTime']}", Fore.CYAN)
            self.log(f"     Filters: {s['object_filters']}", Fore.CYAN)
        
        # PRIORITY ORDER for selecting search:
        # 1. Succeeded/Completed searches (best) - MOST RECENT FIRST
        # 2. In-progress searches (still processing) - MOST RECENT FIRST
        # 3. Any other status - MOST RECENT FIRST
        
        succeeded = []
        in_progress = []
        others = []
        
        for s in matching_searches:
            if s['status'] in ['succeeded', 'completed', 'partiallysucceeded']:
                succeeded.append(s)
            elif s['status'] not in ['failed', 'cancelled', 'canceled', 'timeout']:
                in_progress.append(s)
            else:
                others.append(s)
        
        # Sort by createdDateTime (most recent first)
        succeeded.sort(key=lambda x: x['createdDateTime'], reverse=True)
        in_progress.sort(key=lambda x: x['createdDateTime'], reverse=True)
        others.sort(key=lambda x: x['createdDateTime'], reverse=True)
        
        # Select based on priority
        if succeeded:
            selected = succeeded[0]
            self.log(f"✅ Found COMPLETED search: {selected['id'][:20]}... (status: {selected['status']}, created: {selected['createdDateTime']})", Fore.GREEN)
            self.log(f"   Using this completed search instead of waiting for in-progress ones", Fore.GREEN)
            return selected['id']
        
        if in_progress:
            selected = in_progress[0]
            self.log(f"🔄 Found IN-PROGRESS search: {selected['id'][:20]}... (status: {selected['status']}, created: {selected['createdDateTime']})", Fore.YELLOW)
            return selected['id']
        
        if others:
            selected = others[0]
            self.log(f"⚠️ Using last resort search: {selected['id'][:20]}... (status: {selected['status']}, created: {selected['createdDateTime']})", Fore.YELLOW)
            return selected['id']
        
        return None

    def get_or_create_search(self, site):
        """Get existing search (prioritizing completed) or create a new one"""
        
        # FIRST: Check if we already have this site in our completed state
        if site in self.completed_searches:
            search_id = self.completed_searches[site].get("SearchId")
            if search_id:
                # Verify it's still valid
                status = self.get_search_status(search_id)
                if status in ['succeeded', 'completed']:
                    self.log(f"✅ Using completed search from state: {search_id}", Fore.GREEN)
                    return search_id
        
        # SECOND: Check if we have it in our existing searches state
        if site in self.existing_searches:
            search_id = self.existing_searches[site]["SearchId"]
            self.log(f"Found search ID in local state: {search_id}", Fore.CYAN)
            
            # Verify it still exists in Graph and get its status
            status = self.get_search_status(search_id)
            if status in ['succeeded', 'completed']:
                self.log(f"✅ Search already completed: {search_id}", Fore.GREEN)
                # Mark as completed
                self.completed_searches[site] = {
                    "CompletedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
                    "RecordCount": 0,  # Will be updated if we download
                    "SearchId": search_id
                }
                self.save_completed_state()
                return search_id
            elif status and status not in ['failed', 'cancelled', 'timeout']:
                return search_id
            else:
                self.log(f"Search {search_id} invalid (status: {status}), will find better one", Fore.YELLOW)
        
        # THIRD: Check Graph for existing searches (prioritizing completed)
        existing_id = self.find_existing_search_in_graph(site)
        if existing_id:
            self.log(f"✅ Using existing search from Graph: {existing_id}", Fore.GREEN)
            # Store in state
            self.existing_searches[site] = {
                "SearchId": existing_id,
                "CreatedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')
            }
            self.save_search_state()
            return existing_id
        
        # FOURTH: Create new search if none exists
        self.log(f"🆕 Creating new search for {site}...", Fore.YELLOW)
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
            self.log(f"❌ Failed to create search for {site}", Fore.RED)
            return None
            
        search_data = response.json()
        search_id = search_data.get("id")
        
        # Store in state
        self.existing_searches[site] = {
            "SearchId": search_id,
            "CreatedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z')
        }
        self.save_search_state()
        
        self.log(f"✅ Created new search: {search_id}", Fore.GREEN)
        return search_id

    def get_search_status(self, search_id):
        """Get the status of a search"""
        response = self.make_api_request(
            "GET",
            f"https://graph.microsoft.com/beta/security/auditLog/queries/{search_id}"
        )
        
        if not response:
            return "failed"
            
        search_status = response.json()
        status = (search_status.get("status") or "").strip().lower()
        
        # Log status changes
        if status in ["succeeded", "completed", "partiallysucceeded"]:
            color = Fore.GREEN
            self.log(f"✅ Search {search_id[:20]}... status: {status}", color)
        elif status in ["failed", "cancelled"]:
            color = Fore.RED
            self.log(f"❌ Search {search_id[:20]}... status: {status}", color)
        elif status:
            color = Fore.YELLOW
            self.log(f"⏳ Search {search_id[:20]}... status: {status} (in progress)", color)
        
        return status

    def get_audit_records(self, search_id):
        """Get all audit records for a search"""
        all_records = []
        url = f"https://graph.microsoft.com/beta/security/auditLog/queries/{search_id}/records?$top=1000"
        
        page_count = 0
        while url:
            response = self.make_api_request("GET", url)
            
            if not response:
                break
                
            data = response.json()
            records = data.get("value", [])
            all_records.extend(records)
            page_count += 1
            
            self.log(f"Retrieved page {page_count}: {len(records)} records (total: {len(all_records)})", Fore.YELLOW)
            
            if "@odata.nextLink" in data:
                url = data["@odata.nextLink"]
                time.sleep(0.5)
            else:
                url = None

        return all_records

    def clean_audit_data(self, audit_data):
        """Clean and format audit data"""
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
        """Save audit records to CSV"""
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
                self.log(f"No records for {site}. Created placeholder.", Fore.YELLOW)
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
                    "CreationTime", "Id", "userPrincipalName", "ObjectId",
                    "Operation", "ClientIP", "ItemType"
                ]
            elif mode == "ALL_FIELDS":
                fieldnames = all_audit_fields
            else:
                raise ValueError(f"Invalid CSV_REPORT_MODE '{CSV_REPORT_MODE}'")

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
            
            self.log(f"✅ Saved {len(records)} records to {filename}", Fore.GREEN)
            return len(records)
        except Exception as e:
            self.log(f"❌ Failed to save CSV: {e}", Fore.RED, logging.ERROR)
            return 0

    def generate_summary(self):
        """Generate summary CSV"""
        if not self.summary_data:
            self.log("No data for summary", Fore.YELLOW)
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
            self.log(f"✅ Summary generated: {summary_file}", Fore.GREEN)
        except Exception as e:
            self.log(f"Failed to generate summary: {e}", Fore.RED, logging.ERROR)

    def get_site_library_info(self, site_url):
        """Resolve site library info"""
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
            self.log(f"Failed to resolve Graph site for {site_url}", Fore.RED)
            return None

        site_id = site_response.json().get("id")
        if not site_id:
            return None

        drives_response = self.make_api_request(
            "GET",
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
        )
        if not drives_response:
            return None

        drives = drives_response.json().get("value", [])
        matching_drive = next(
            (drive for drive in drives if drive.get("name") == DOCUMENT_LIBRARY_NAME),
            None
        )
        if not matching_drive:
            self.log(f"Library '{DOCUMENT_LIBRARY_NAME}' not found", Fore.RED)
            return None

        library_info = {
            "site_id": site_id,
            "drive_id": matching_drive["id"],
            "drive_name": matching_drive["name"]
        }
        self.site_library_cache[site_url] = library_info
        return library_info

    def upload_csv_files(self):
        """Upload CSV files to SharePoint"""
        csv_files = glob.glob(os.path.join(self.output_dir, "*.csv"))
        csv_files = [f for f in csv_files if not os.path.basename(f).startswith("AuditSummary")]
        
        if not csv_files:
            self.log("No CSV files to upload", Fore.YELLOW)
            return
        
        self.log(f"Starting upload of {len(csv_files)} CSV files...", Fore.CYAN)
        
        for csv_file in csv_files:
            filename = os.path.basename(csv_file)
            site_name = filename.split('_')[0]
            site_url = next((site for site in SITES if site.endswith(f"/{site_name}")), None)
            
            if not site_url:
                self.log(f"Could not map {filename} to site, skipping", Fore.YELLOW)
                continue
                
            try:
                self.log(f"Uploading {filename} to {site_url}...", Fore.YELLOW)
                result = self.upload_file_to_sharepoint(csv_file, site_url)
                
                if result:
                    self.uploaded_files.append(csv_file)
                    self.log(f"✅ Uploaded {filename}", Fore.GREEN)
                    
            except Exception as e:
                self.log(f"❌ Failed to upload {filename}: {str(e)}", Fore.RED)
        
        self.log("Upload complete!", Fore.GREEN)

    def upload_file_to_sharepoint(self, file_path, site_url, folder_path=''):
        """Upload file to SharePoint"""
        library_info = self.get_site_library_info(site_url)
        if not library_info:
            raise Exception(f"Could not find library for {site_url}")

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
        """Upload small files"""
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
        
        if not response or response.status_code not in (200, 201):
            raise Exception(f"Upload failed")
        
        return response.json()

    def upload_large_file(self, file_path, site_id, drive_id, folder_path):
        """Upload large files with chunking"""
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
                
                if not response or response.status_code not in (200, 201, 202):
                    raise Exception(f"Upload failed at chunk {chunk_num+1}")
                
                self.log(f"Uploaded chunk {chunk_num+1}/{total_chunks} ({end/file_size:.1%})", Fore.YELLOW)
        
        return response.json()

    def update_file_properties(self, site_id, drive_id, file_item_id, file_name):
        """Update file properties"""
        fields_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_item_id}/listItem/fields"
        
        update_payload = {
            "Month": self.REPORT_MONTH,
            "Year": self.REPORT_YEAR
        }
        
        update_response = self.make_api_request("PATCH", fields_url, json_data=update_payload)
        
        if update_response:
            self.log(f"✅ Updated properties for {file_name}: Month={self.REPORT_MONTH}, Year={self.REPORT_YEAR}", Fore.CYAN)
        
        return update_response

    def process_pending_searches(self):
        """Process all searches"""
        total_searches = len(SITES)
        
        # First, get or create searches for all sites
        self.log("Getting or creating searches for all sites...", Fore.CYAN)
        for site in SITES:
            if site in self.completed_searches:
                self.log(f"✅ Site {site} already completed, skipping", Fore.GREEN)
                continue
                
            search_id = self.get_or_create_search(site)
            if not search_id:
                self.log(f"❌ Failed to get/create search for {site}", Fore.RED)
                continue
                
            self.log(f"✅ Using search {search_id} for {site}", Fore.GREEN)

        # Now process each search
        terminal_success_statuses = {"succeeded", "completed", "partiallysucceeded", "noresults", "nodata"}
        terminal_failure_statuses = {"failed", "cancelled", "canceled", "timeout"}

        while len(self.completed_searches) < total_searches:
            self.log(f"Processing searches ({len(self.completed_searches)}/{total_searches} completed)...", Fore.YELLOW)
            
            for site in SITES:
                if site in self.completed_searches:
                    continue
                    
                if site not in self.existing_searches:
                    self.log(f"No search found for {site}, skipping", Fore.RED)
                    continue
                    
                search_id = self.existing_searches[site]["SearchId"]
                status = self.get_search_status(search_id)
                self.search_poll_counts[site] = self.search_poll_counts.get(site, 0) + 1

                if status in terminal_success_statuses:
                    self.log(f"✅ Search completed for {site}, retrieving records...", Fore.GREEN)
                    records = self.get_audit_records(search_id)
                    
                    record_count = self.save_audit_to_csv(records, site)
                    
                    self.summary_data.append({
                        "Site": site.split('/')[-1],
                        "RecordCount": record_count
                    })
                    
                    self.completed_searches[site] = {
                        "CompletedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
                        "RecordCount": record_count,
                        "SearchId": search_id
                    }
                    self.save_completed_state()
                    
                elif status in terminal_failure_statuses:
                    self.log(f"❌ Search failed for {site} with status: {status}", Fore.RED)
                    self.completed_searches[site] = {
                        "CompletedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
                        "RecordCount": 0,
                        "Status": status,
                        "SearchId": search_id
                    }
                    self.save_completed_state()
                    
                elif self.search_poll_counts[site] >= MAX_SEARCH_STATUS_POLLS:
                    self.log(f"⏰ Search for {site} timed out after {MAX_SEARCH_STATUS_POLLS} polls", Fore.YELLOW)
                    self.completed_searches[site] = {
                        "CompletedTime": datetime.now(timezone.utc).isoformat().replace('+00:00', 'Z'),
                        "RecordCount": 0,
                        "Status": "timeout",
                        "SearchId": search_id
                    }
                    self.save_completed_state()
            
            if len(self.completed_searches) < total_searches:
                self.log(f"Waiting {RETRY_DELAY_SECONDS} seconds before next poll...", Fore.YELLOW)
                time.sleep(RETRY_DELAY_SECONDS)

        self.log(f"✅ Completed processing all {total_searches} searches", Fore.GREEN)

    def run(self):
        try:
            self.process_start_time = datetime.now().astimezone()
            
            self.log("=" * 60, Fore.CYAN)
            self.log("AUDIT LOG COLLECTOR STARTED", Fore.CYAN)
            self.log("=" * 60, Fore.CYAN)
            self.log(f"Date range: {self.START_DATE} to {self.END_DATE}", Fore.CYAN)
            self.log(f"Report: {self.REPORT_MONTH} {self.REPORT_YEAR}", Fore.CYAN)
            
            self.log("Authenticating...", Fore.YELLOW)
            if not self.get_access_token():
                self.log("❌ Failed to authenticate", Fore.RED)
                return
            
            self.process_pending_searches()
            self.generate_summary()
            
            self.log("Authenticating for SharePoint upload...", Fore.YELLOW)
            if not self.get_access_token():
                self.log("❌ Failed to authenticate for upload", Fore.RED)
            else:
                self.upload_csv_files()
            
            self.process_end_time = datetime.now().astimezone()
            
            self.log("=" * 60, Fore.GREEN)
            self.log("✅ ALL OPERATIONS COMPLETED SUCCESSFULLY!", Fore.GREEN)
            self.log("=" * 60, Fore.GREEN)
            
            # Summary
            total_records = sum(item.get("RecordCount", 0) for item in self.completed_searches.values())
            self.log(f"Total records collected: {total_records}", Fore.CYAN)
            self.log(f"Output directory: {os.path.abspath(self.output_dir)}", Fore.CYAN)

        except Exception as e:
            self.log(f"❌ Fatal error: {e}", Fore.RED)
            raise

if __name__ == "__main__":
    collector = AuditLogCollector()
    collector.run()
