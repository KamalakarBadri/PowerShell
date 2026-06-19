import requests
import os
import csv
import datetime
import time
import json
import logging
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, List, Set
import shutil
import re

# ==================== CONFIGURATION ====================
TENANT_ID = "your-tenant-id"
CLIENT_ID = "your-client-id"
CLIENT_SECRET = "your-client-secret"

# File names
CSV_REPORT_FILE = "azure_ad_users_master_report.csv"
CHANGE_LOG_FILE = "user_change_history.csv"
DELETED_USERS_FILE = "deleted_users_audit.csv"
BACKUP_FILE = "azure_ad_users_report_backup.csv"

# Performance settings
MAX_WORKERS = 20
BATCH_SIZE = 100
API_BATCH_SIZE = 20

# Rate limiting
REQUESTS_PER_SECOND = 25
MIN_REQUEST_INTERVAL = 1.0 / REQUESTS_PER_SECOND

# Timeouts
CONNECTION_TIMEOUT = 60
READ_TIMEOUT = 120

# Set up logging - ONLY to console, no file
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Microsoft Graph API endpoints
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
BATCH_API_URL = f"{GRAPH_API_BASE}/$batch"
DELETED_USERS_URL = f"{GRAPH_API_BASE}/directory/deletedItems/microsoft.graph.user"

# ==================== AUTHENTICATION ====================
def get_access_token():
    """Get OAuth2 access token with retry logic"""
    token_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    
    for attempt in range(5):
        try:
            data = {
                "client_id": CLIENT_ID,
                "client_secret": CLIENT_SECRET,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials"
            }
            
            response = requests.post(token_url, data=data, timeout=30)
            response.raise_for_status()
            token = response.json()["access_token"]
            logger.info("Authentication successful")
            return token
        except Exception as e:
            logger.error(f"Auth attempt {attempt + 1} failed: {e}")
            time.sleep(2 ** attempt)
    
    raise Exception("Authentication failed after 5 attempts")

# ==================== DELETED USERS FUNCTIONS ====================
def get_deleted_users(headers):
    """Get all soft-deleted users from Azure AD recycle bin"""
    all_deleted = []
    url = DELETED_USERS_URL
    
    while url:
        try:
            response = requests.get(
                url,
                headers=headers,
                timeout=(CONNECTION_TIMEOUT, READ_TIMEOUT)
            )
            
            if response.status_code == 429:
                logger.warning("Rate limited while fetching deleted users. Waiting...")
                time.sleep(10)
                continue
            
            response.raise_for_status()
            data = response.json()
            
            if "value" in data:
                for user in data["value"]:
                    deleted_time = user.get("deletedDateTime")
                    all_deleted.append({
                        "id": user.get("id"),
                        "displayName": user.get("displayName", ""),
                        "userPrincipalName": user.get("userPrincipalName", ""),
                        "deletedDateTime": deleted_time if deleted_time else datetime.datetime.now().isoformat(),
                        "deletionDetected": datetime.datetime.now().isoformat()
                    })
            
            url = data.get("@odata.nextLink")
            time.sleep(MIN_REQUEST_INTERVAL)
            
        except Exception as e:
            logger.error(f"Error fetching deleted users: {e}")
            break
    
    logger.info(f"Found {len(all_deleted)} soft-deleted users in recycle bin")
    return all_deleted

# ==================== GRAPH API CALLS ====================
def batch_graph_requests_with_retry(headers, requests_list, max_retries=3):
    """Send batch requests with retry logic"""
    for attempt in range(max_retries):
        try:
            batch_body = {"requests": requests_list}
            response = requests.post(
                BATCH_API_URL,
                headers=headers,
                json=batch_body,
                timeout=(CONNECTION_TIMEOUT, READ_TIMEOUT)
            )
            
            if response.status_code == 429:
                wait_time = 2 ** attempt * 5
                logger.warning(f"Rate limited. Waiting {wait_time}s...")
                time.sleep(wait_time)
                continue
            
            response.raise_for_status()
            return response.json().get("responses", [])
            
        except requests.exceptions.Timeout:
            logger.warning(f"Batch request timeout, attempt {attempt + 1}")
            time.sleep(2 ** attempt)
        except Exception as e:
            logger.warning(f"Batch request failed, attempt {attempt + 1}: {e}")
            time.sleep(2 ** attempt)
    
    return []

def get_all_users_filtered(headers):
    """Get all users including disabled, excluding guests"""
    all_users = []
    url = f"{GRAPH_API_BASE}/users"
    params = {
        "$select": "id,displayName,userPrincipalName,mail,createdDateTime,userType,accountEnabled,department,jobTitle",
        "$filter": "userType eq 'Member'",
        "$top": 100
    }
    
    page_count = 0
    
    while url:
        page_count += 1
        try:
            response = requests.get(
                url,
                headers=headers,
                params=params,
                timeout=(CONNECTION_TIMEOUT, READ_TIMEOUT)
            )
            
            if response.status_code == 429:
                logger.warning(f"Rate limited. Waiting...")
                time.sleep(10)
                continue
            
            if response.status_code != 200:
                logger.error(f"API Error: {response.status_code} - {response.text}")
                response.raise_for_status()
            
            data = response.json()
            
            if "value" in data:
                all_users.extend(data["value"])
                logger.info(f"  Page {page_count}: Fetched {len(data['value'])} users (Total: {len(all_users)})")
            
            url = data.get("@odata.nextLink")
            params = None
            time.sleep(MIN_REQUEST_INTERVAL)
            
        except Exception as e:
            logger.error(f"Error fetching page {page_count}: {e}")
            if page_count == 1:
                raise
            time.sleep(5)
            continue
    
    logger.info(f"Total active/disabled users fetched: {len(all_users)}")
    return all_users

def get_managers_batch_optimized(headers, user_ids):
    """Get managers using batch API"""
    if not user_ids:
        return {}
    
    all_managers = {}
    
    for i in range(0, len(user_ids), API_BATCH_SIZE):
        chunk = user_ids[i:i + API_BATCH_SIZE]
        
        requests_list = []
        for idx, user_id in enumerate(chunk):
            requests_list.append({
                "id": str(idx + 1),
                "method": "GET",
                "url": f"/users/{user_id}/manager"
            })
        
        responses = batch_graph_requests_with_retry(headers, requests_list)
        
        manager_ids = {}
        for response in responses:
            if response.get("status") == 200:
                manager_data = response.get("body", {})
                if manager_data and "id" in manager_data:
                    user_id = chunk[int(response["id"]) - 1]
                    manager_ids[user_id] = manager_data["id"]
        
        if manager_ids:
            manager_details = get_user_details_batch_optimized(
                headers, 
                list(manager_ids.values())
            )
            
            for user_id, manager_id in manager_ids.items():
                if manager_id in manager_details:
                    all_managers[user_id] = manager_details[manager_id]
        
        time.sleep(MIN_REQUEST_INTERVAL)
    
    return all_managers

def get_user_details_batch_optimized(headers, user_ids):
    """Get user details with proper batching"""
    if not user_ids:
        return {}
    
    all_users = {}
    
    for i in range(0, len(user_ids), API_BATCH_SIZE):
        chunk = user_ids[i:i + API_BATCH_SIZE]
        
        requests_list = []
        for idx, user_id in enumerate(chunk):
            requests_list.append({
                "id": str(idx + 1),
                "method": "GET",
                "url": f"/users/{user_id}?$select=id,displayName,mail,userPrincipalName"
            })
        
        responses = batch_graph_requests_with_retry(headers, requests_list)
        
        for response in responses:
            if response.get("status") == 200:
                user_data = response.get("body", {})
                user_id = user_data.get("id")
                if user_id:
                    all_users[user_id] = {
                        "displayName": user_data.get("displayName", ""),
                        "mail": user_data.get("mail", ""),
                        "userPrincipalName": user_data.get("userPrincipalName", "")
                    }
        
        time.sleep(MIN_REQUEST_INTERVAL)
    
    return all_users

def get_onedrives_batch_optimized(headers, user_ids):
    """Get OneDrive URLs with chunking"""
    if not user_ids:
        return {}
    
    all_onedrives = {}
    processed = 0
    
    for i in range(0, len(user_ids), API_BATCH_SIZE):
        chunk = user_ids[i:i + API_BATCH_SIZE]
        
        requests_list = []
        for idx, user_id in enumerate(chunk):
            requests_list.append({
                "id": str(idx + 1),
                "method": "GET",
                "url": f"/users/{user_id}/drive?$select=webUrl"
            })
        
        responses = batch_graph_requests_with_retry(headers, requests_list)
        
        for response in responses:
            if response.get("status") == 200:
                drive_data = response.get("body", {})
                if drive_data and "webUrl" in drive_data:
                    user_id = chunk[int(response["id"]) - 1]
                    all_onedrives[user_id] = drive_data["webUrl"]
            elif response.get("status") == 404:
                user_id = chunk[int(response["id"]) - 1]
                all_onedrives[user_id] = ""
        
        processed += len(chunk)
        if processed % 1000 == 0:
            logger.info(f"  OneDrive processed: {processed}/{len(user_ids)}")
        
        time.sleep(MIN_REQUEST_INTERVAL)
    
    return all_onedrives

# ==================== HELPER FUNCTIONS ====================
def extract_upn_from_onedrive_url(onedrive_url):
    """
    Extract UPN from OneDrive URL
    Example: https://company-my.sharepoint.com/personal/john_doe_company_com/Documents
    Returns: john.doe@company.com
    """
    if not onedrive_url:
        return None
    
    try:
        # Pattern: /personal/{upn_with_underscores}/
        match = re.search(r'/personal/([^/]+)/', onedrive_url)
        if match:
            upn_with_underscores = match.group(1)
            # Convert underscores back to dots and @
            # john_doe_company_com -> john.doe@company.com
            parts = upn_with_underscores.split('_')
            if len(parts) >= 2:
                # Last part is domain
                domain = parts[-1]
                # First part(s) is username
                username = '_'.join(parts[:-1])
                # If domain contains multiple parts (e.g., company_com)
                if '_' in domain:
                    domain = domain.replace('_', '.')
                return f"{username}@{domain}"
        return None
    except Exception:
        return None

def normalize_value(value):
    """Normalize values for comparison (handle None, empty strings, etc.)"""
    if value is None:
        return ""
    return str(value).strip()

def append_to_change_history(existing_history, new_change, timestamp):
    """Append a new change to the change history"""
    if existing_history:
        return f"{existing_history} | {timestamp}: {new_change}"
    else:
        return f"{timestamp}: {new_change}"

# ==================== CHANGE DETECTION WITH HISTORY ====================
def detect_changes_with_history(existing_user, current_user, manager_name, manager_email, onedrive_url):
    """
    Detect changes and maintain change history
    Returns: (change_type, changes_list, old_upn, updated_change_history)
    """
    changes = []
    old_upn = None
    
    # If no existing user, it's new
    if existing_user is None:
        return "NEW_USER", [], None, ""
    
    # If user was previously marked as deleted and now exists, it's reactivated
    if existing_user.get('status') == 'deleted':
        return "REACTIVATED", [], None, ""
    
    # Get existing change history
    existing_history = existing_user.get('changeHistory', '')
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Check for UPN change (MOST IMPORTANT)
    old_upn = normalize_value(existing_user.get('userPrincipalName', ''))
    new_upn = normalize_value(current_user.get('userPrincipalName', ''))
    
    upn_changed = False
    if old_upn and new_upn and old_upn != new_upn:
        upn_changed = True
        change_detail = f"UPN changed: '{old_upn}' -> '{new_upn}'"
        changes.append({
            'field': 'userPrincipalName',
            'old_value': old_upn,
            'new_value': new_upn,
            'change_type': 'UPN_CHANGED',
            'detail': change_detail
        })
        existing_history = append_to_change_history(existing_history, change_detail, current_time)
    
    # Check OneDrive URL change (ALWAYS track if changed)
    old_onedrive = normalize_value(existing_user.get('onedrive_url', ''))
    new_onedrive = normalize_value(onedrive_url)
    
    if old_onedrive != new_onedrive:
        # If UPN changed, this is expected - we can note it
        if upn_changed:
            change_detail = f"OneDrive URL updated due to UPN change: '{old_onedrive}' -> '{new_onedrive}'"
        else:
            change_detail = f"OneDrive URL changed: '{old_onedrive}' -> '{new_onedrive}'"
        
        changes.append({
            'field': 'onedrive_url',
            'old_value': old_onedrive,
            'new_value': new_onedrive,
            'change_type': 'ONEDRIVE_CHANGED',
            'detail': change_detail
        })
        existing_history = append_to_change_history(existing_history, change_detail, current_time)
    
    # Check displayName
    old_val = normalize_value(existing_user.get('displayName', ''))
    new_val = normalize_value(current_user.get('displayName', ''))
    if old_val != new_val:
        change_detail = f"DisplayName changed: '{old_val}' -> '{new_val}'"
        changes.append({
            'field': 'displayName',
            'old_value': old_val,
            'new_value': new_val,
            'change_type': 'FIELD_CHANGED',
            'detail': change_detail
        })
        existing_history = append_to_change_history(existing_history, change_detail, current_time)
    
    # Check mail
    old_val = normalize_value(existing_user.get('mail', ''))
    new_val = normalize_value(current_user.get('mail', ''))
    if old_val != new_val:
        change_detail = f"Email changed: '{old_val}' -> '{new_val}'"
        changes.append({
            'field': 'mail',
            'old_value': old_val,
            'new_value': new_val,
            'change_type': 'FIELD_CHANGED',
            'detail': change_detail
        })
        existing_history = append_to_change_history(existing_history, change_detail, current_time)
    
    # Check department
    old_val = normalize_value(existing_user.get('department', ''))
    new_val = normalize_value(current_user.get('department', ''))
    if old_val != new_val:
        change_detail = f"Department changed: '{old_val}' -> '{new_val}'"
        changes.append({
            'field': 'department',
            'old_value': old_val,
            'new_value': new_val,
            'change_type': 'FIELD_CHANGED',
            'detail': change_detail
        })
        existing_history = append_to_change_history(existing_history, change_detail, current_time)
    
    # Check jobTitle
    old_val = normalize_value(existing_user.get('jobTitle', ''))
    new_val = normalize_value(current_user.get('jobTitle', ''))
    if old_val != new_val:
        change_detail = f"Job Title changed: '{old_val}' -> '{new_val}'"
        changes.append({
            'field': 'jobTitle',
            'old_value': old_val,
            'new_value': new_val,
            'change_type': 'FIELD_CHANGED',
            'detail': change_detail
        })
        existing_history = append_to_change_history(existing_history, change_detail, current_time)
    
    # Check manager
    old_manager = normalize_value(existing_user.get('manager', ''))
    new_manager = normalize_value(manager_name)
    if old_manager != new_manager:
        change_detail = f"Manager changed: '{old_manager}' -> '{new_manager}'"
        changes.append({
            'field': 'manager',
            'old_value': old_manager,
            'new_value': new_manager,
            'change_type': 'MANAGER_CHANGED',
            'detail': change_detail
        })
        existing_history = append_to_change_history(existing_history, change_detail, current_time)
    
    # Check manager email
    old_manager_email = normalize_value(existing_user.get('managerEmail', ''))
    new_manager_email = normalize_value(manager_email)
    if old_manager_email != new_manager_email:
        change_detail = f"Manager Email changed: '{old_manager_email}' -> '{new_manager_email}'"
        changes.append({
            'field': 'managerEmail',
            'old_value': old_manager_email,
            'new_value': new_manager_email,
            'change_type': 'MANAGER_CHANGED',
            'detail': change_detail
        })
        existing_history = append_to_change_history(existing_history, change_detail, current_time)
    
    # Check status (active/disabled)
    old_status = normalize_value(existing_user.get('status', ''))
    new_status = 'active' if current_user.get('accountEnabled', True) else 'disabled'
    if old_status != new_status:
        change_detail = f"Status changed: '{old_status}' -> '{new_status}'"
        changes.append({
            'field': 'status',
            'old_value': old_status,
            'new_value': new_status,
            'change_type': 'STATUS_CHANGED',
            'detail': change_detail
        })
        existing_history = append_to_change_history(existing_history, change_detail, current_time)
    
    if changes:
        return "UPDATED", changes, old_upn if old_upn else None, existing_history
    
    return "NO_CHANGE", [], None, existing_history

# ==================== CSV OPERATIONS ====================
def load_existing_csv():
    """Load existing CSV report"""
    if not os.path.exists(CSV_REPORT_FILE):
        logger.info("No existing report found. Creating new one.")
        return []
    
    try:
        with open(CSV_REPORT_FILE, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            data = list(reader)
            logger.info(f"Loaded {len(data)} users from existing report")
            return data
    except Exception as e:
        logger.error(f"Error loading CSV: {e}")
        return []

def save_csv_report(users_data):
    """Save the complete report to CSV with change history"""
    if not users_data:
        logger.warning("No data to save")
        return
    
    fieldnames = [
        'id', 'displayName', 'userPrincipalName', 'old_userPrincipalName', 
        'mail', 'createdDateTime', 'department', 'jobTitle',
        'manager', 'managerEmail', 'onedrive_url', 'old_onedrive_url',
        'status', 'lastUpdated', 'changeType', 'changeDetails',
        'changeHistory',
        'deletedDateTime'
    ]
    
    try:
        # Create backup
        if os.path.exists(CSV_REPORT_FILE):
            shutil.copy2(CSV_REPORT_FILE, BACKUP_FILE)
            logger.info(f"Backup created: {BACKUP_FILE}")
        
        with open(CSV_REPORT_FILE, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
            writer.writeheader()
            
            for user in users_data:
                row = {field: user.get(field, '') for field in fieldnames}
                writer.writerow(row)
        
        logger.info(f"✓ CSV report saved: {CSV_REPORT_FILE}")
        logger.info(f"  Total rows: {len(users_data)}")
        
    except Exception as e:
        logger.error(f"Error saving CSV: {e}")

def save_deleted_users_report(deleted_users):
    """Save deleted users to a separate audit file"""
    if not deleted_users:
        logger.info("No deleted users to save")
        return
    
    fieldnames = ['id', 'displayName', 'userPrincipalName', 'deletedDateTime', 'deletionDetected']
    
    try:
        with open(DELETED_USERS_FILE, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            
            for user in deleted_users:
                row = {field: user.get(field, '') for field in fieldnames}
                writer.writerow(row)
        
        logger.info(f"✓ Deleted users report saved: {DELETED_USERS_FILE} ({len(deleted_users)} users)")
        
    except Exception as e:
        logger.error(f"Error saving deleted users report: {e}")

def save_change_log(changes):
    """Save detailed change log to CSV"""
    if not changes:
        return
    
    fieldnames = [
        'timestamp', 'userId', 'userPrincipalName', 'changeType',
        'field', 'old_value', 'new_value', 'details'
    ]
    
    file_exists = os.path.exists(CHANGE_LOG_FILE)
    
    try:
        with open(CHANGE_LOG_FILE, 'a', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            
            if not file_exists:
                writer.writeheader()
            
            for change in changes:
                writer.writerow(change)
        
        logger.info(f"✓ Change log saved: {CHANGE_LOG_FILE} ({len(changes)} changes)")
        
    except Exception as e:
        logger.error(f"Error saving change log: {e}")

# ==================== MAIN PROCESSING ====================
def process_users(current_users, existing_data, deleted_users_list, headers):
    """
    Process all users with complete change history
    """
    logger.info("\n" + "="*60)
    logger.info("STARTING USER PROCESSING")
    logger.info("="*60)
    logger.info(f"Current users from Azure AD: {len(current_users)}")
    logger.info(f"Existing users in report: {len(existing_data)}")
    logger.info(f"Deleted users in recycle bin: {len(deleted_users_list)}")
    
    # Create sets for quick lookups
    current_user_ids = {user["id"] for user in current_users}
    
    # Create dictionaries for existing data
    existing_dict = {user.get('id'): user for user in existing_data if user.get('id')}
    
    # Results will contain ALL users (active + disabled + deleted)
    final_results = []
    changes_log = []
    
    # Track which users we've processed
    processed_user_ids = set()
    
    # Statistics
    stats = {
        'active': 0,
        'disabled': 0,
        'deleted': 0,
        'new_users': 0,
        'updated_users': 0,
        'reactivated_users': 0,
        'no_change': 0,
        'deleted_from_azure': 0
    }
    
    # ============================================================
    # STEP 1: Process all active/disabled users from Azure AD
    # ============================================================
    logger.info("\n" + "-"*60)
    logger.info("PROCESSING ACTIVE/DISABLED USERS")
    logger.info("-"*60)
    
    # Process in chunks for performance
    chunks = []
    for i in range(0, len(current_users), BATCH_SIZE):
        chunk = current_users[i:i + BATCH_SIZE]
        chunks.append(chunk)
    
    logger.info(f"Processing {len(current_users)} users in {len(chunks)} chunks")
    
    processed_count = 0
    updated_count = 0
    no_change_count = 0
    
    for chunk_idx, chunk in enumerate(chunks, 1):
        logger.info(f"  Processing chunk {chunk_idx}/{len(chunks)}...")
        user_ids = [user["id"] for user in chunk]
        
        # Get managers and OneDrive for this chunk
        managers = get_managers_batch_optimized(headers, user_ids)
        onedrives = get_onedrives_batch_optimized(headers, user_ids)
        
        for user in chunk:
            user_id = user["id"]
            user_type = user.get("userType", "")
            
            # Skip guest users
            if user_type.lower() == "guest":
                continue
            
            # Mark as processed
            processed_user_ids.add(user_id)
            
            # Get enriched data
            manager_info = managers.get(user_id)
            manager_name = manager_info.get("displayName") if manager_info else ""
            manager_email = manager_info.get("mail") or manager_info.get("userPrincipalName") if manager_info else ""
            onedrive_url = onedrives.get(user_id, "")
            
            # Determine status
            is_active = user.get("accountEnabled", True)
            status = "active" if is_active else "disabled"
            
            if status == "active":
                stats['active'] += 1
            else:
                stats['disabled'] += 1
            
            # Check if user exists in existing report
            existing_user = existing_dict.get(user_id)
            
            # Get old OneDrive URL if it exists (for tracking)
            old_onedrive_url = existing_user.get('onedrive_url', '') if existing_user else ""
            
            # Detect changes with history
            change_type, changes, old_upn, change_history = detect_changes_with_history(
                existing_user, user, manager_name, manager_email, onedrive_url
            )
            
            # Update statistics
            if change_type == "NEW_USER":
                stats['new_users'] += 1
                logger.info(f"    NEW USER: {user.get('userPrincipalName', 'Unknown')}")
            elif change_type == "REACTIVATED":
                stats['reactivated_users'] += 1
                logger.info(f"    REACTIVATED: {user.get('userPrincipalName', 'Unknown')}")
            elif change_type == "UPDATED":
                stats['updated_users'] += 1
                updated_count += 1
                change_details = "; ".join([c['detail'] for c in changes])
                logger.info(f"    UPDATED: {user.get('userPrincipalName', 'Unknown')} - {change_details}")
            else:
                stats['no_change'] += 1
                no_change_count += 1
            
            # Build user data with change history
            user_data = {
                "id": user_id,
                "displayName": user.get("displayName", ""),
                "userPrincipalName": user.get("userPrincipalName", ""),
                "old_userPrincipalName": old_upn if old_upn else "",
                "mail": user.get("mail", ""),
                "createdDateTime": user.get("createdDateTime", ""),
                "department": user.get("department", ""),
                "jobTitle": user.get("jobTitle", ""),
                "manager": manager_name,
                "managerEmail": manager_email,
                "onedrive_url": onedrive_url,
                "old_onedrive_url": old_onedrive_url if old_onedrive_url else "",
                "status": status,
                "lastUpdated": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "changeType": change_type,
                "changeDetails": "; ".join([c['detail'] for c in changes]) if changes else "",
                "changeHistory": change_history,
                "deletedDateTime": ""
            }
            
            final_results.append(user_data)
            
            # Log changes to change log file
            if changes:
                timestamp = datetime.datetime.now().isoformat()
                for change in changes:
                    changes_log.append({
                        'timestamp': timestamp,
                        'userId': user_id,
                        'userPrincipalName': user.get('userPrincipalName', ''),
                        'changeType': change_type,
                        'field': change['field'],
                        'old_value': str(change['old_value'])[:500],
                        'new_value': str(change['new_value'])[:500],
                        'details': change['detail']
                    })
            
            processed_count += 1
            
            if processed_count % 100 == 0:
                logger.info(f"    Processed {processed_count}/{len(current_users)} users... (Updated: {updated_count}, No Change: {no_change_count})")
    
    # ============================================================
    # STEP 2: Handle users that are in the report but NOT in Azure AD (DELETED)
    # IMPORTANT: We NEVER remove these users from the report
    # ============================================================
    logger.info("\n" + "-"*60)
    logger.info("CHECKING FOR DELETED USERS (KEEPING IN REPORT)")
    logger.info("-"*60)
    
    current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Find users that exist in the report but were NOT processed
    for user_id, user_data in existing_dict.items():
        if user_id not in processed_user_ids:
            # This user was deleted from Azure AD
            if user_data.get('status') != 'deleted':
                # Check if this user is currently in the recycle bin
                deleted_info = next((u for u in deleted_users_list if u["id"] == user_id), None)
                deleted_time = deleted_info.get("deletedDateTime") if deleted_info else ""
                
                # Update change history for deletion
                existing_history = user_data.get('changeHistory', '')
                deletion_detail = f"User deleted from Azure AD at {deleted_time if deleted_time else 'unknown time'}"
                existing_history = append_to_change_history(existing_history, deletion_detail, current_time)
                
                # Mark as deleted - KEEP IN REPORT
                user_data["status"] = "deleted"
                user_data["lastUpdated"] = current_time
                user_data["changeType"] = "DELETED"
                user_data["changeDetails"] = deletion_detail
                user_data["deletedDateTime"] = deleted_time if deleted_time else datetime.datetime.now().isoformat()
                user_data["changeHistory"] = existing_history
                
                final_results.append(user_data)
                stats['deleted'] += 1
                stats['deleted_from_azure'] += 1
                
                # Log deletion
                changes_log.append({
                    'timestamp': datetime.datetime.now().isoformat(),
                    'userId': user_id,
                    'userPrincipalName': user_data.get('userPrincipalName', ''),
                    'changeType': 'DELETED',
                    'field': 'status',
                    'old_value': user_data.get('status', 'active'),
                    'new_value': 'deleted',
                    'details': deletion_detail
                })
                
                logger.info(f"  DELETED (keeping in report): {user_data.get('userPrincipalName', 'Unknown')} (Deleted at: {deleted_time if deleted_time else 'unknown'})")
            else:
                # User is already marked as deleted, just keep them in the report
                if 'changeHistory' not in user_data:
                    user_data['changeHistory'] = user_data.get('changeDetails', '')
                final_results.append(user_data)
                stats['deleted'] += 1
    
    # ============================================================
    # STEP 3: Generate reports
    # ============================================================
    
    # Save deleted users report
    if deleted_users_list:
        save_deleted_users_report(deleted_users_list)
    else:
        deleted_in_report = [u for u in final_results if u.get('status') == 'deleted']
        if deleted_in_report:
            deleted_report_data = []
            for user in deleted_in_report:
                deleted_report_data.append({
                    "id": user.get("id", ""),
                    "displayName": user.get("displayName", ""),
                    "userPrincipalName": user.get("userPrincipalName", ""),
                    "deletedDateTime": user.get("deletedDateTime", datetime.datetime.now().isoformat()),
                    "deletionDetected": user.get("lastUpdated", datetime.datetime.now().isoformat())
                })
            save_deleted_users_report(deleted_report_data)
        else:
            logger.info("No deleted users to report")
    
    # Save change log
    if changes_log:
        save_change_log(changes_log)
    
    # Summary
    logger.info("\n" + "="*60)
    logger.info("PROCESSING SUMMARY")
    logger.info("="*60)
    logger.info(f"Total users in Azure AD: {len(current_users)}")
    logger.info(f"  - Active users: {stats['active']}")
    logger.info(f"  - Disabled users: {stats['disabled']}")
    logger.info(f"Total users in report: {len(final_results)}")
    logger.info(f"  - Active: {len([u for u in final_results if u.get('status') == 'active'])}")
    logger.info(f"  - Disabled: {len([u for u in final_results if u.get('status') == 'disabled'])}")
    logger.info(f"  - Deleted: {len([u for u in final_results if u.get('status') == 'deleted'])}")
    logger.info(f"New users added: {stats['new_users']}")
    logger.info(f"Users updated: {stats['updated_users']}")
    logger.info(f"Users with no changes: {stats['no_change']}")
    logger.info(f"Users reactivated: {stats['reactivated_users']}")
    logger.info(f"Users marked as deleted (kept in report): {stats['deleted_from_azure']}")
    logger.info(f"Total changes logged: {len(changes_log)}")
    logger.info("="*60)
    
    return final_results

# ==================== MAIN FUNCTION ====================
def main():
    """Main function"""
    start_time = time.time()
    
    logger.info("="*60)
    logger.info("AZURE AD USER REPORT - ENTERPRISE EDITION")
    logger.info("With Complete Change History Tracking")
    logger.info("="*60)
    logger.info(f"Started at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("="*60)
    
    # Get access token
    logger.info("\n1. Authenticating...")
    try:
        global headers
        access_token = get_access_token()
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        logger.info("   ✓ Authentication successful")
    except Exception as e:
        logger.error(f"   ✗ Authentication failed: {e}")
        return
    
    # Load existing CSV
    logger.info("\n2. Loading existing CSV report...")
    existing_data = load_existing_csv()
    
    # Get deleted users from recycle bin
    logger.info("\n3. Fetching soft-deleted users from Azure AD...")
    deleted_users_list = []
    try:
        deleted_users_list = get_deleted_users(headers)
        logger.info(f"   Found {len(deleted_users_list)} soft-deleted users in recycle bin")
    except Exception as e:
        logger.error(f"   ✗ Error fetching deleted users: {e}")
    
    # Fetch active/disabled users
    logger.info("\n4. Fetching active/disabled users from Azure AD...")
    logger.info("   (Includes both active and disabled users)")
    fetch_start = time.time()
    
    try:
        current_users = get_all_users_filtered(headers)
        fetch_time = time.time() - fetch_start
        logger.info(f"   ✓ Found {len(current_users)} Member users")
        logger.info(f"   Fetch time: {fetch_time:.2f} seconds")
    except Exception as e:
        logger.error(f"   ✗ Error fetching users: {e}")
        return
    
    # Process users
    logger.info("\n5. Processing users with complete change history...")
    process_start = time.time()
    
    processed_data = process_users(
        current_users, existing_data, deleted_users_list, headers
    )
    process_time = time.time() - process_start
    logger.info(f"   Processing time: {process_time:.2f} seconds")
    
    # Save CSV
    logger.info("\n6. Saving CSV report...")
    save_csv_report(processed_data)
    
    total_time = time.time() - start_time
    
    logger.info("\n" + "="*60)
    logger.info("✅ COMPLETED SUCCESSFULLY")
    logger.info("="*60)
    logger.info(f"Report saved to: {CSV_REPORT_FILE}")
    logger.info(f"Change log saved to: {CHANGE_LOG_FILE}")
    logger.info(f"Deleted users report saved to: {DELETED_USERS_FILE}")
    logger.info(f"Total execution time: {total_time:.2f} seconds ({total_time/60:.1f} minutes)")
    logger.info(f"Completed at: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info("="*60)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.info("\n\nScript interrupted by user")
    except Exception as e:
        logger.error(f"\n\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
