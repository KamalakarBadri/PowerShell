import csv
import requests
import json
import uuid
import base64
import time
import logging
import argparse
import sys
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

# Configuration
CONFIG = {
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "tenant_name": "geekbyteonline.onmicrosoft.com",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    "repair_account": "edit@geekbyte.online",
    "new_id_site_url": "https://geekbyteonline.sharepoint.com/sites/2DayRetention",
    "scopes": {
        "graph": "https://graph.microsoft.com/.default",
        "sharepoint": "https://geekbyteonline.sharepoint.com/.default"
    }
}

# Token cache with expiration tracking
TOKEN_CACHE = {
    'graph': {'token': None, 'expires_at': 0},
    'sharepoint': {'token': None, 'expires_at': 0}
}

def setup_logging(mode, site_type):
    """Setup logging based on mode and site type"""
    log_file = f"nameid_cleanup_{site_type}_{mode}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)

def get_token_with_certificate(scope):
    """Get access token using certificate-based authentication"""
    try:
        if not os.path.exists(CONFIG['certificate_path']) or not os.path.exists(CONFIG['private_key_path']):
            raise Exception("Certificate files not found. Please ensure certificate.pem and private_key.pem exist.")
            
        with open(CONFIG['certificate_path'], "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
        with open(CONFIG['private_key_path'], "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())

        now = int(time.time())
        jwt_header = {
            "alg": "RS256",
            "typ": "JWT",
            "x5t": base64.urlsafe_b64encode(certificate.fingerprint(hashes.SHA1())).decode().rstrip('=')
        }
        jwt_payload = {
            "aud": f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            "exp": now + 300,
            "iss": CONFIG['app_id'],
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": CONFIG['app_id']
        }

        encoded_header = base64.urlsafe_b64encode(json.dumps(jwt_header).encode()).decode().rstrip('=')
        encoded_payload = base64.urlsafe_b64encode(json.dumps(jwt_payload).encode()).decode().rstrip('=')
        jwt_unsigned = f"{encoded_header}.{encoded_payload}"
        signature = private_key.sign(jwt_unsigned.encode(), padding.PKCS1v15(), hashes.SHA256())
        encoded_signature = base64.urlsafe_b64encode(signature).decode().rstrip('=')
        jwt = f"{jwt_unsigned}.{encoded_signature}"

        token_response = requests.post(
            f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            data={
                "client_id": CONFIG['app_id'],
                "client_assertion": jwt,
                "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                "scope": scope,
                "grant_type": "client_credentials"
            }
        )

        if token_response.status_code == 200:
            token_data = token_response.json()
            access_token = token_data["access_token"]
            expires_in = token_data.get("expires_in", 3600)  # Default to 1 hour if not provided
            
            logger.info(f"Successfully obtained token using certificate, expires in {expires_in} seconds")
            return access_token, now + expires_in - 300  # Subtract 5 minutes for safety buffer
        else:
            logger.error(f"Certificate token request failed: {token_response.text}")
            raise Exception(f"Token request failed: {token_response.text}")
            
    except Exception as e:
        logger.exception("Certificate authentication failed")
        raise Exception(f"Certificate authentication failed: {str(e)}")

def get_cached_token(token_type):
    """Get cached token or fetch new one if expired"""
    cache_key = 'graph' if 'graph' in token_type else 'sharepoint'
    cached_token = TOKEN_CACHE[cache_key]
    
    # Check if token is still valid (with 5 minute buffer)
    if cached_token['token'] and time.time() < cached_token['expires_at']:
        logger.debug(f"Using cached {cache_key.upper()} token")
        return cached_token['token']
    
    # Token expired or not available, get new one
    logger.info(f"Fetching new {cache_key.upper()} token")
    scope = CONFIG['scopes'][cache_key]
    token, expires_at = get_token_with_certificate(scope)
    
    # Update cache
    TOKEN_CACHE[cache_key]['token'] = token
    TOKEN_CACHE[cache_key]['expires_at'] = expires_at
    
    return token

def get_onedrive_url(onedrive_owner_upn):
    """Get OneDrive URL for a specific user"""
    try:
        # Get Graph API token from cache
        graph_token = get_cached_token('graph')
        
        if not graph_token:
            raise Exception("Failed to obtain Graph API access token")
        
        # Get OneDrive URL for the owner
        graph_headers = {
            "Authorization": f"Bearer {graph_token}",
            "Content-Type": "application/json"
        }
        
        logger.info(f"Getting OneDrive for owner: {onedrive_owner_upn}")
        onedrive_response = requests.get(
            f"https://graph.microsoft.com/v1.0/users/{onedrive_owner_upn}/drive?$select=webUrl",
            headers=graph_headers
        )
        
        if onedrive_response.status_code != 200:
            logger.error(f"OneDrive lookup failed: {onedrive_response.text}")
            raise Exception(f"OneDrive not found for owner: {onedrive_response.text}")
        
        onedrive_info = onedrive_response.json()
        onedrive_url = onedrive_info.get('webUrl', '')
        
        if not onedrive_url:
            raise Exception("OneDrive URL not found for owner")
        
        # Remove /Documents from the end if present
        if onedrive_url.endswith('/Documents'):
            site_url = onedrive_url[:-10]
        else:
            site_url = onedrive_url
        
        logger.info(f"OneDrive URL: {site_url}")
        return site_url
        
    except Exception as e:
        logger.exception(f"Failed to get OneDrive URL for {onedrive_owner_upn}")
        raise

def parse_site_users_json(json_content, target_upn):
    """Parse SharePoint site users JSON response and find specific user"""
    try:
        logger.debug(f"Parsing JSON for user: {target_upn}")
        
        # Handle both direct JSON object and response.json() structure
        if isinstance(json_content, dict):
            data = json_content
        else:
            data = json_content.json() if hasattr(json_content, 'json') else json_content
        
        # Handle different response formats
        if 'd' in data:
            # OData format with 'd' wrapper
            if 'results' in data['d']:
                users = data['d']['results']
            else:
                users = [data['d']]  # Single user
        elif 'value' in data:
            # Direct value array
            users = data['value']
        else:
            # Assume it's already a list or single user
            users = data if isinstance(data, list) else [data]
        
        logger.debug(f"Processing {len(users)} users")
        
        for user in users:
            # Extract user properties with safe access
            user_id = user.get('Id')
            title = user.get('Title')
            email = user.get('Email')
            login_name = user.get('LoginName')
            user_principal_name = user.get('UserPrincipalName')
            is_site_admin = user.get('IsSiteAdmin', False)
            
            # Extract NameId from UserId if available
            current_nameid = None
            user_id_obj = user.get('UserId')
            if user_id_obj and isinstance(user_id_obj, dict):
                current_nameid = user_id_obj.get('NameId')
            
            logger.debug(f"Checking user - ID: {user_id}, Email: {email}, Login: {login_name}, UPN: {user_principal_name}, NameId: {current_nameid}")
            
            # Check if this is the target user using multiple identifiers
            email_match = email and email.lower() == target_upn.lower()
            login_match = login_name and target_upn.lower() in login_name.lower()
            upn_match = user_principal_name and user_principal_name.lower() == target_upn.lower()
            
            if email_match or login_match or upn_match:
                user_info = {
                    'user_id': user_id,
                    'title': title,
                    'email': email,
                    'login_name': login_name,
                    'user_principal_name': user_principal_name,
                    'is_site_admin': is_site_admin,
                    'current_nameid': current_nameid
                }
                logger.info(f"Found target user: {user_info}")
                return user_info
        
        logger.warning(f"User {target_upn} not found in site users")
        return None
        
    except Exception as e:
        logger.exception("Failed to parse site users JSON")
        raise Exception(f"Failed to parse JSON response: {str(e)}")

def find_user_on_site(target_upn, site_url):
    """Find a specific user on a SharePoint site/OneDrive"""
    try:
        # Ensure site URL ends properly
        if not site_url.endswith('/'):
            site_url += '/'
        
        # Construct SharePoint API URL
        site_users_url = f"{site_url}_api/web/siteusers"
        logger.info(f"SharePoint API URL: {site_users_url}")
        
        # Get SharePoint token from cache
        sharepoint_token = get_cached_token('sharepoint')
        
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Call SharePoint API to get site users with detailed logging
        sharepoint_headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/json;odata=verbose"
        }
        
        logger.info("Calling SharePoint API to get site users")
        site_users_response = requests.get(site_users_url, headers=sharepoint_headers)
        
        logger.info(f"Response status: {site_users_response.status_code}")
        
        if site_users_response.status_code != 200:
            logger.error(f"SharePoint API failed: {site_users_response.status_code} - {site_users_response.text}")
            raise Exception(f"Failed to get site users: {site_users_response.text}")
        
        # Parse JSON and find the specific user
        user_info = parse_site_users_json(site_users_response, target_upn)
        
        if not user_info:
            # Try alternative approach - get user by login name directly
            logger.info(f"User not found in site users list, trying direct lookup...")
            user_info = get_user_by_login_name(site_url, sharepoint_token, target_upn)
        
        return user_info
        
    except Exception as e:
        logger.exception(f"Failed to find user {target_upn} on site {site_url}")
        raise

def get_user_by_login_name(site_url, token, target_upn):
    """Try to get user directly by login name"""
    try:
        # Try different login name formats
        login_formats = [
            target_upn,
            f"i:0#.f|membership|{target_upn}",
            f"i:0#.f|membership|{target_upn.lower()}",
            f"i:0%23.f|membership|{target_upn}",
            f"i:0%23.f|membership|{target_upn.lower()}"
        ]
        
        for login_format in login_formats:
            try:
                user_url = f"{site_url}_api/web/siteusers('{login_format}')"
                headers = {
                    "Authorization": f"Bearer {token}",
                    "Accept": "application/json;odata=verbose"
                }
                
                logger.info(f"Trying direct user lookup with: {login_format}")
                response = requests.get(user_url, headers=headers)
                
                if response.status_code == 200:
                    user_data = response.json()
                    user_info = parse_site_users_json(user_data, target_upn)
                    if user_info:
                        logger.info(f"Found user using direct lookup with format: {login_format}")
                        return user_info
                
            except Exception as e:
                logger.debug(f"Direct lookup failed for format {login_format}: {str(e)}")
                continue
        
        return None
        
    except Exception as e:
        logger.exception("Error in direct user lookup")
        return None

def remove_user_from_site(user_id, site_url):
    """Remove a user from a SharePoint site/OneDrive"""
    try:
        # Ensure site URL ends properly
        if not site_url.endswith('/'):
            site_url += '/'
        
        # Construct SharePoint API URL for user removal
        remove_user_url = f"{site_url}_api/web/siteusers/removebyid({user_id})"
        logger.info(f"Remove user URL: {remove_user_url}")
        
        # Get SharePoint token from cache
        sharepoint_token = get_cached_token('sharepoint')
        
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Call SharePoint API to remove user
        sharepoint_headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
        }
        
        logger.info("Calling SharePoint API to remove user")
        remove_response = requests.post(remove_user_url, headers=sharepoint_headers)
        
        logger.info(f"Remove response: {remove_response.status_code} - {remove_response.text}")
        
        if remove_response.status_code not in [200, 204]:
            raise Exception(f"Failed to remove user: {remove_response.text}")
        
        logger.info(f"Successfully removed user {user_id} from {site_url}")
        return True
        
    except Exception as e:
        logger.exception(f"Failed to remove user {user_id} from site {site_url}")
        raise

def get_new_site_nameid(target_upn):
    """Get NameId for user from new ID site by ensuring user exists there"""
    try:
        new_site_url = CONFIG['new_id_site_url']
        if not new_site_url.endswith('/'):
            new_site_url += '/'
        
        # Get SharePoint token from cache
        sharepoint_token = get_cached_token('sharepoint')
        
        if not sharepoint_token:
            logger.error("Failed to obtain SharePoint access token for new ID site")
            return None
        
        # Get request digest
        request_digest = get_request_digest(new_site_url, sharepoint_token)
        if not request_digest:
            logger.error("Failed to get request digest for new ID site")
            return None
        
        # Ensure user exists on new ID site
        ensure_url = f"{new_site_url}_api/web/ensureuser"
        headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": request_digest
        }
        
        body = {
            "logonName": target_upn
        }
        
        logger.info(f"Ensuring user {target_upn} on new ID site")
        response = requests.post(ensure_url, headers=headers, json=body)
        
        logger.info(f"Ensure user response status: {response.status_code}")
        
        if response.status_code == 200:
            user_data = response.json()
            
            # Extract NameId from UserId
            user_id_obj = user_data.get('d', {}).get('UserId')
            if user_id_obj and isinstance(user_id_obj, dict):
                nameid = user_id_obj.get('NameId')
                logger.info(f"Retrieved NameId from new ID site: {nameid}")
                return nameid
            else:
                logger.warning("UserId object not found or invalid in new ID site response")
                return None
        else:
            logger.error(f"Failed to ensure user on new ID site: {response.status_code} - {response.text}")
            return None
            
    except Exception as e:
        logger.exception(f"Error getting NameId from new ID site for user {target_upn}")
        return None

def get_request_digest(site_url, token):
    """Get request digest for SharePoint operations"""
    try:
        digest_url = f"{site_url}_api/contextinfo"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose"
        }
        
        response = requests.post(digest_url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            return data['d']['GetContextWebInformation']['FormDigestValue']
        else:
            logger.error(f"Failed to get request digest: {response.text}")
            return None
    except Exception as e:
        logger.exception("Error getting request digest")
        return None

def process_site(target_upn, site_identifier, site_type, remove_mismatch=False):
    """Process a single site (OneDrive or SharePoint) to check NameId mismatch with optional removal"""
    result = {
        'site_identifier': site_identifier,
        'site_type': site_type,
        'found': False,
        'nameid_mismatch': False,
        'nameid_status': 'NOT_FOUND',
        'removed': False,
        'error': None,
        'site_url': None,
        'user_id': None,
        'current_nameid': None,
        'new_nameid': None,
        'user_email': None,
        'user_login_name': None,
        'is_site_admin': False,
        'processed_at': datetime.now().isoformat()
    }
    
    try:
        # Get site URL based on type
        if site_type == 'onedrive':
            # site_identifier is OneDrive owner UPN
            site_url = get_onedrive_url(site_identifier)
            result['site_url'] = site_url
            result['onedrive_owner'] = site_identifier
        else:
            # site_identifier is direct SharePoint site URL
            site_url = site_identifier
            if not site_url.endswith('/'):
                site_url += '/'
            result['site_url'] = site_url
            result['sharepoint_site'] = site_identifier
        
        # Find user on this site
        user_info = find_user_on_site(target_upn, site_url)
        
        if user_info:
            result['found'] = True
            result['user_id'] = user_info['user_id']
            result['current_nameid'] = user_info.get('current_nameid')
            result['user_email'] = user_info.get('email')
            result['user_login_name'] = user_info.get('login_name')
            result['is_site_admin'] = user_info.get('is_site_admin', False)
            
            # Get NameId from new ID site for comparison
            new_nameid = get_new_site_nameid(target_upn)
            result['new_nameid'] = new_nameid
            
            # Check for NameId mismatch
            if result['current_nameid'] and new_nameid:
                nameid_match = result['current_nameid'] == new_nameid
                result['nameid_mismatch'] = not nameid_match
                result['nameid_status'] = 'MATCH' if nameid_match else 'MISMATCH'
                
                if result['nameid_mismatch']:
                    site_display = site_identifier if site_type == 'sharepoint' else f"{site_identifier}'s OneDrive"
                    logger.info(f"NameId MISMATCH detected for {target_upn} on {site_display}")
                    logger.info(f"   Current NameId: {result['current_nameid']}")
                    logger.info(f"   New NameId:     {result['new_nameid']}")
                    
                    # Remove user if requested
                    if remove_mismatch:
                        logger.info(f"Removing user due to mismatch (remove_mismatch={remove_mismatch})")
                        remove_user_from_site(result['user_id'], site_url)
                        result['removed'] = True
                        logger.info(f"Successfully removed {target_upn} from {site_display}")
                    else:
                        logger.info(f"Reporting only - user not removed (remove_mismatch={remove_mismatch})")
                else:
                    site_display = site_identifier if site_type == 'sharepoint' else f"{site_identifier}'s OneDrive"
                    logger.info(f"NameId MATCH for {target_upn} on {site_display}")
            else:
                result['nameid_status'] = 'INCOMPLETE_DATA'
                if not result['current_nameid']:
                    logger.warning(f"Current NameId not found for {target_upn}")
                if not new_nameid:
                    logger.warning(f"New NameId not found for {target_upn}")
        else:
            site_display = site_identifier if site_type == 'sharepoint' else f"{site_identifier}'s OneDrive"
            logger.info(f"User {target_upn} not found on {site_display}")
        
    except Exception as e:
        logger.exception(f"Error processing {site_type} for {site_identifier}")
        result['error'] = str(e)
    
    return result

def read_sites_from_csv(csv_file_path, site_type):
    """Read site list from CSV file"""
    sites = []
    try:
        with open(csv_file_path, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if row and row[0].strip():  # Skip empty rows
                    sites.append(row[0].strip())
        logger.info(f"Read {len(sites)} {site_type} sites from CSV")
        return sites
    except Exception as e:
        logger.exception(f"Error reading CSV file: {csv_file_path}")
        raise

def generate_detailed_report(results, target_upn, mode, site_type):
    """Generate a detailed report with statistics and recommendations"""
    
    # Calculate statistics
    total_processed = len(results)
    found_count = len([r for r in results if r['found']])
    mismatch_count = len([r for r in results if r['nameid_mismatch']])
    match_count = len([r for r in results if r['nameid_status'] == 'MATCH'])
    incomplete_count = len([r for r in results if r['nameid_status'] == 'INCOMPLETE_DATA'])
    error_count = len([r for r in results if r['error']])
    removed_count = len([r for r in results if r['removed']])
    
    # Generate report
    report = []
    report.append("=" * 80)
    if mode == "report":
        report.append(f"NAMEID MISMATCH REPORT - {site_type.upper()} - READ ONLY (NO USERS REMOVED)")
    else:
        report.append(f"NAMEID MISMATCH CLEANUP REPORT - {site_type.upper()} - USERS REMOVED")
    report.append("=" * 80)
    report.append(f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report.append(f"Target User: {target_upn}")
    report.append(f"Site Type: {site_type.upper()}")
    report.append(f"Mode: {mode.upper()}")
    report.append("")
    report.append("SUMMARY STATISTICS:")
    report.append("-" * 40)
    report.append(f"Total Sites Processed: {total_processed}")
    report.append(f"User Found On: {found_count} Sites")
    report.append(f"NameId Matches: {match_count}")
    report.append(f"NameId Mismatches: {mismatch_count}")
    if mode == "cleanup":
        report.append(f"Users Removed: {removed_count}")
    report.append(f"Incomplete Data: {incomplete_count}")
    report.append(f"Errors: {error_count}")
    report.append("")
    
    if mismatch_count > 0:
        if mode == "report":
            report.append("RECOMMENDED ACTIONS:")
            report.append("-" * 40)
            report.append(f"Found {mismatch_count} site(s) with NameId mismatches that need attention.")
            report.append("Run with '--mode cleanup' to automatically remove these users.")
            report.append("")
        else:
            report.append("ACTIONS TAKEN:")
            report.append("-" * 40)
            report.append(f"Successfully removed user from {removed_count} site(s) with NameId mismatches.")
            report.append("")
    
    report.append("DETAILED RESULTS:")
    report.append("-" * 40)
    
    # Group results by status
    mismatch_results = [r for r in results if r['nameid_mismatch']]
    match_results = [r for r in results if r['nameid_status'] == 'MATCH']
    incomplete_results = [r for r in results if r['nameid_status'] == 'INCOMPLETE_DATA']
    not_found_results = [r for r in results if not r['found'] and not r['error']]
    error_results = [r for r in results if r['error']]
    
    if mismatch_results:
        report.append("")
        if mode == "report":
            report.append("MISMATCHES (Require Action):")
        else:
            report.append("MISMATCHES (Removed):")
        for result in mismatch_results:
            identifier = result['site_identifier']
            report.append(f"  * {identifier}")
            report.append(f"    Current NameId: {result['current_nameid']}")
            report.append(f"    New NameId:     {result['new_nameid']}")
            if mode == "cleanup":
                report.append(f"    Status:         {'REMOVED' if result['removed'] else 'NOT REMOVED'}")
            report.append(f"    Site URL:       {result['site_url']}")
    
    if match_results:
        report.append("")
        report.append("MATCHES (No Action Needed):")
        for result in match_results[:10]:  # Show first 10 matches
            report.append(f"  * {result['site_identifier']}")
        if len(match_results) > 10:
            report.append(f"  ... and {len(match_results) - 10} more")
    
    if incomplete_results:
        report.append("")
        report.append("INCOMPLETE DATA (Manual Verification Needed):")
        for result in incomplete_results:
            report.append(f"  * {result['site_identifier']}")
            if not result['current_nameid']:
                report.append(f"    Missing Current NameId")
            if not result['new_nameid']:
                report.append(f"    Missing New NameId")
    
    if not_found_results:
        report.append("")
        report.append("USER NOT FOUND:")
        for result in not_found_results[:10]:  # Show first 10 not found
            report.append(f"  * {result['site_identifier']}")
        if len(not_found_results) > 10:
            report.append(f"  ... and {len(not_found_results) - 10} more")
    
    if error_results:
        report.append("")
        report.append("ERRORS:")
        for result in error_results:
            report.append(f"  * {result['site_identifier']}: {result['error']}")
    
    report.append("")
    report.append("=" * 80)
    
    return "\n".join(report)

def confirm_cleanup(mismatch_count, site_type):
    """Get confirmation before performing cleanup"""
    print(f"\nWARNING: You are about to remove the user from {mismatch_count} {site_type} sites.")
    print("This action cannot be undone!")
    print("\nPlease confirm by typing 'YES' to continue or anything else to cancel:")
    
    confirmation = input().strip().upper()
    return confirmation == 'YES'

def main():
    """Main function with both report and cleanup modes for OneDrive and SharePoint sites"""
    
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='NameId Mismatch Reporter and Cleanup Tool for OneDrive and SharePoint')
    parser.add_argument('--target', '-t', required=True, help='Target user UPN (e.g., user@domain.com)')
    parser.add_argument('--csv', '-c', required=True, help='Path to CSV file with site list')
    parser.add_argument('--type', choices=['onedrive', 'sharepoint'], required=True,
                       help='Site type: onedrive (list of UPNs) or sharepoint (list of site URLs)')
    parser.add_argument('--mode', '-m', choices=['report', 'cleanup'], default='report',
                       help='Operation mode: report (read-only) or cleanup (remove users)')
    parser.add_argument('--workers', '-w', type=int, default=5, help='Number of concurrent workers')
    
    args = parser.parse_args()
    
    # Setup logging
    global logger
    logger = setup_logging(args.mode, args.type)
    
    logger.info(f"Starting NameId mismatch tool")
    logger.info(f"Site Type: {args.type}")
    logger.info(f"Target User: {args.target}")
    logger.info(f"CSV File: {args.csv}")
    logger.info(f"Mode: {args.mode}")
    logger.info(f"Workers: {args.workers}")
    
    if args.mode == 'report':
        logger.info("REPORT MODE: No users will be removed")
    else:
        logger.info("CLEANUP MODE: Users will be removed from mismatched sites")
    
    try:
        # Read sites from CSV
        sites = read_sites_from_csv(args.csv, args.type)
        
        if not sites:
            logger.error("No sites found in CSV file")
            return
        
        # Pre-fetch tokens to warm up the cache
        logger.info("Pre-fetching access tokens...")
        get_cached_token('sharepoint')
        if args.type == 'onedrive':
            get_cached_token('graph')
        
        # First, run in report mode to get mismatch count
        logger.info(f"Scanning {len(sites)} {args.type} sites...")
        
        results = []
        
        with ThreadPoolExecutor(max_workers=args.workers) as executor:
            # Submit all tasks
            future_to_site = {
                executor.submit(process_site, args.target, site, args.type, remove_mismatch=False): site 
                for site in sites
            }
            
            # Process completed tasks with progress tracking
            completed = 0
            total = len(sites)
            
            for future in as_completed(future_to_site):
                result = future.result()
                results.append(result)
                completed += 1
                
                site_display = result['site_identifier']
                if args.type == 'onedrive':
                    site_display = f"{result['site_identifier']}'s OneDrive"
                
                logger.info(f"Progress: {completed}/{total} ({completed/total*100:.1f}%) - {site_display}: {result['nameid_status']}")
        
        # Count mismatches
        mismatch_count = len([r for r in results if r['nameid_mismatch']])
        
        # If in cleanup mode and mismatches found, get confirmation
        if args.mode == 'cleanup' and mismatch_count > 0:
            if not confirm_cleanup(mismatch_count, args.type):
                logger.info("Cleanup cancelled by user")
                print("Operation cancelled.")
                return
            
            # Re-process mismatched sites with removal
            logger.info("Processing mismatches with removal...")
            for result in results:
                if result['nameid_mismatch'] and not result['removed']:
                    try:
                        remove_user_from_site(result['user_id'], result['site_url'])
                        result['removed'] = True
                        site_display = result['site_identifier']
                        if args.type == 'onedrive':
                            site_display = f"{result['site_identifier']}'s OneDrive"
                        logger.info(f"Removed {args.target} from {site_display}")
                    except Exception as e:
                        logger.error(f"Failed to remove user from {result['site_identifier']}: {str(e)}")
                        result['error'] = f"Removal failed: {str(e)}"
        
        # Generate and save reports
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # 1. Save detailed CSV report
        csv_output_file = f"nameid_{args.type}_{args.mode}_{args.target.split('@')[0]}_{timestamp}.csv"
        with open(csv_output_file, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'site_identifier', 'site_type', 'found', 'nameid_status', 'nameid_mismatch', 'removed',
                'current_nameid', 'new_nameid', 'user_email', 'user_login_name', 'is_site_admin', 
                'site_url', 'user_id', 'error', 'processed_at'
            ]
            # Add type-specific fields
            if args.type == 'onedrive':
                fieldnames.insert(2, 'onedrive_owner')
            else:
                fieldnames.insert(2, 'sharepoint_site')
                
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for result in results:
                writer.writerow(result)
        
        logger.info(f"Detailed CSV report saved to: {csv_output_file}")
        
        # 2. Generate and save summary report
        summary_report = generate_detailed_report(results, args.target, args.mode, args.type)
        summary_output_file = f"nameid_{args.type}_{args.mode}_summary_{args.target.split('@')[0]}_{timestamp}.txt"
        
        with open(summary_output_file, 'w', encoding='utf-8') as f:
            f.write(summary_report)
        
        logger.info(f"Summary report saved to: {summary_output_file}")
        
        # 3. Print summary to console
        print("\n" + "=" * 80)
        print(f"OPERATION COMPLETE - {args.type.upper()} - {args.mode.upper()} MODE")
        print("=" * 80)
        print(summary_report)
        
        # 4. Save action items if in report mode
        if args.mode == 'report' and mismatch_count > 0:
            mismatch_output_file = f"nameid_{args.type}_mismatches_{args.target.split('@')[0]}_{timestamp}.csv"
            mismatch_results = [r for r in results if r['nameid_mismatch']]
            
            with open(mismatch_output_file, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['site_identifier', 'current_nameid', 'new_nameid', 'site_url', 'user_id']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                
                writer.writeheader()
                for result in mismatch_results:
                    writer.writerow({
                        'site_identifier': result['site_identifier'],
                        'current_nameid': result['current_nameid'],
                        'new_nameid': result['new_nameid'],
                        'site_url': result['site_url'],
                        'user_id': result['user_id']
                    })
            
            logger.info(f"Mismatch-only CSV saved to: {mismatch_output_file}")
            logger.info(f"Run with '--mode cleanup' to remove users from {mismatch_count} site(s)")
        
    except Exception as e:
        logger.exception("Fatal error during processing")
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()