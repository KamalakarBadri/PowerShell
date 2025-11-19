import csv
import requests
import json
import uuid
import base64
import time
import logging
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
    "client_secret": "CyG8Q~FYHuCMSyVmt4sNxt5IejrMc2c24Ziz4a.t",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    "repair_account": "edit@geekbyte.online",
    "new_id_site_url": "https://geekbyteonline.sharepoint.com/sites/2DayRetention",
    "scopes": {
        "graph": "https://graph.microsoft.com/.default",
        "sharepoint": "https://geekbyteonline.sharepoint.com/.default"
    }
}

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('nameid_report.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

def get_token_with_certificate(scope):
    """Get access token using certificate-based authentication"""
    try:
        if not os.path.exists(CONFIG['certificate_path']) or not os.path.exists(CONFIG['private_key_path']):
            logger.warning("Certificate files not found, falling back to client secret")
            return None
            
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
            logger.info("Successfully obtained token using certificate")
            return token_response.json()["access_token"]
        else:
            logger.error(f"Certificate token request failed: {token_response.text}")
            return None
            
    except Exception as e:
        logger.exception("Certificate authentication failed")
        return None

def get_token_with_secret(scope):
    """Get access token using client secret authentication"""
    try:
        token_url = f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token"
        
        token_data = {
            "client_id": CONFIG['app_id'],
            "client_secret": CONFIG['client_secret'],
            "scope": scope,
            "grant_type": "client_credentials"
        }
        
        token_response = requests.post(token_url, data=token_data)

        if token_response.status_code == 200:
            logger.info("Successfully obtained token using client secret")
            return token_response.json()["access_token"]
        else:
            logger.error(f"Client secret token request failed: {token_response.text}")
            return None
            
    except Exception as e:
        logger.exception("Client secret authentication failed")
        return None

def get_onedrive_url(onedrive_owner_upn):
    """Get OneDrive URL for a specific user"""
    try:
        # Get Graph API token
        graph_token = get_token_with_certificate(CONFIG['scopes']['graph'])
        if not graph_token:
            graph_token = get_token_with_secret(CONFIG['scopes']['graph'])
        
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
        
        # Get SharePoint token
        sharepoint_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sharepoint_token:
            sharepoint_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
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

def get_new_site_nameid(target_upn):
    """Get NameId for user from new ID site by ensuring user exists there"""
    try:
        new_site_url = CONFIG['new_id_site_url']
        if not new_site_url.endswith('/'):
            new_site_url += '/'
        
        # Get SharePoint token
        sharepoint_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sharepoint_token:
            sharepoint_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
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

def process_onedrive_owner(target_upn, owner):
    """Process a single OneDrive owner to check NameId mismatch (NO REMOVAL)"""
    result = {
        'onedrive_owner': owner,
        'found': False,
        'nameid_mismatch': False,
        'nameid_status': 'NOT_FOUND',
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
        # Get OneDrive URL
        site_url = get_onedrive_url(owner)
        result['site_url'] = site_url
        
        # Find user on this OneDrive
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
                    logger.info(f"🚨 NameId MISMATCH detected for {target_upn} on {owner}'s OneDrive")
                    logger.info(f"   Current NameId: {result['current_nameid']}")
                    logger.info(f"   New NameId:     {result['new_nameid']}")
                else:
                    logger.info(f"✅ NameId MATCH for {target_upn} on {owner}'s OneDrive")
            else:
                result['nameid_status'] = 'INCOMPLETE_DATA'
                if not result['current_nameid']:
                    logger.warning(f"⚠️  Current NameId not found for {target_upn} on {owner}'s OneDrive")
                if not new_nameid:
                    logger.warning(f"⚠️  New NameId not found for {target_upn}")
        else:
            logger.info(f"❌ User {target_upn} not found on {owner}'s OneDrive")
        
    except Exception as e:
        logger.exception(f"Error processing OneDrive for {owner}")
        result['error'] = str(e)
    
    return result

def read_onedrive_owners_from_csv(csv_file_path):
    """Read OneDrive owner list from CSV file"""
    owners = []
    try:
        with open(csv_file_path, 'r', newline='', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                if row and row[0].strip():  # Skip empty rows
                    owners.append(row[0].strip())
        logger.info(f"Read {len(owners)} OneDrive owners from CSV")
        return owners
    except Exception as e:
        logger.exception(f"Error reading CSV file: {csv_file_path}")
        raise

def generate_detailed_report(results, target_upn):
    """Generate a detailed report with statistics and recommendations"""
    
    # Calculate statistics
    total_processed = len(results)
    found_count = len([r for r in results if r['found']])
    mismatch_count = len([r for r in results if r['nameid_mismatch']])
    match_count = len([r for r in results if r['nameid_status'] == 'MATCH'])
    incomplete_count = len([r for r in results if r['nameid_status'] == 'INCOMPLETE_DATA'])
    error_count = len([r for r in results if r['error']])
    
    # Generate report
    report = []
    report.append("=" * 80)
    report.append("NAMEID MISMATCH REPORT - READ ONLY (NO USERS REMOVED)")
    report.append("=" * 80)
    report.append(f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    report.append(f"Target User: {target_upn}")
    report.append("")
    report.append("SUMMARY STATISTICS:")
    report.append("-" * 40)
    report.append(f"Total OneDrives Processed: {total_processed}")
    report.append(f"User Found On: {found_count} OneDrives")
    report.append(f"NameId Matches: {match_count}")
    report.append(f"NameId Mismatches: {mismatch_count}")
    report.append(f"Incomplete Data: {incomplete_count}")
    report.append(f"Errors: {error_count}")
    report.append("")
    
    if mismatch_count > 0:
        report.append("🚨 RECOMMENDED ACTIONS:")
        report.append("-" * 40)
        report.append(f"Found {mismatch_count} OneDrive(s) with NameId mismatches that need attention.")
        report.append("These users should be removed and re-added to fix PUID issues.")
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
        report.append("🚨 MISMATCHES (Require Action):")
        for result in mismatch_results:
            report.append(f"  • {result['onedrive_owner']}")
            report.append(f"    Current NameId: {result['current_nameid']}")
            report.append(f"    New NameId:     {result['new_nameid']}")
            report.append(f"    Site URL:       {result['site_url']}")
    
    if match_results:
        report.append("")
        report.append("✅ MATCHES (No Action Needed):")
        for result in match_results[:10]:  # Show first 10 matches
            report.append(f"  • {result['onedrive_owner']}")
        if len(match_results) > 10:
            report.append(f"  ... and {len(match_results) - 10} more")
    
    if incomplete_results:
        report.append("")
        report.append("⚠️  INCOMPLETE DATA (Manual Verification Needed):")
        for result in incomplete_results:
            report.append(f"  • {result['onedrive_owner']}")
            if not result['current_nameid']:
                report.append(f"    Missing Current NameId")
            if not result['new_nameid']:
                report.append(f"    Missing New NameId")
    
    if not_found_results:
        report.append("")
        report.append("❌ USER NOT FOUND:")
        for result in not_found_results[:10]:  # Show first 10 not found
            report.append(f"  • {result['onedrive_owner']}")
        if len(not_found_results) > 10:
            report.append(f"  ... and {len(not_found_results) - 10} more")
    
    if error_results:
        report.append("")
        report.append("🔴 ERRORS:")
        for result in error_results:
            report.append(f"  • {result['onedrive_owner']}: {result['error']}")
    
    report.append("")
    report.append("=" * 80)
    
    return "\n".join(report)

def main():
    """Main function to generate NameId mismatch report (NO REMOVAL)"""
    
    # Configuration
    TARGET_UPN = "user@geekbyteonline.onmicrosoft.com"  # Replace with target user UPN
    CSV_FILE_PATH = "onedrive_owners.csv"  # Path to CSV file with OneDrive owners
    MAX_WORKERS = 5  # Number of concurrent threads
    
    logger.info(f"🚀 Starting NameId mismatch report generation for user: {TARGET_UPN}")
    logger.info("📊 THIS IS A READ-ONLY REPORT - NO USERS WILL BE REMOVED")
    
    try:
        # Read OneDrive owners from CSV
        onedrive_owners = read_onedrive_owners_from_csv(CSV_FILE_PATH)
        
        if not onedrive_owners:
            logger.error("No OneDrive owners found in CSV file")
            return
        
        results = []
        
        # Process OneDrives in parallel for better performance
        logger.info(f"Processing {len(onedrive_owners)} OneDrive sites with {MAX_WORKERS} workers...")
        
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            # Submit all tasks
            future_to_owner = {
                executor.submit(process_onedrive_owner, TARGET_UPN, owner): owner 
                for owner in onedrive_owners
            }
            
            # Process completed tasks with progress tracking
            completed = 0
            total = len(onedrive_owners)
            
            for future in as_completed(future_to_owner):
                result = future.result()
                results.append(result)
                completed += 1
                logger.info(f"📈 Progress: {completed}/{total} ({completed/total*100:.1f}%) - {result['onedrive_owner']}: {result['nameid_status']}")
        
        # Generate and save reports
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # 1. Save detailed CSV report
        csv_output_file = f"nameid_report_{TARGET_UPN.split('@')[0]}_{timestamp}.csv"
        with open(csv_output_file, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = [
                'onedrive_owner', 'found', 'nameid_status', 'nameid_mismatch', 
                'current_nameid', 'new_nameid', 'user_email', 'user_login_name',
                'is_site_admin', 'site_url', 'user_id', 'error', 'processed_at'
            ]
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for result in results:
                writer.writerow(result)
        
        logger.info(f"💾 Detailed CSV report saved to: {csv_output_file}")
        
        # 2. Generate and save summary report
        summary_report = generate_detailed_report(results, TARGET_UPN)
        summary_output_file = f"nameid_summary_{TARGET_UPN.split('@')[0]}_{timestamp}.txt"
        
        with open(summary_output_file, 'w', encoding='utf-8') as f:
            f.write(summary_report)
        
        logger.info(f"💾 Summary report saved to: {summary_output_file}")
        
        # 3. Print summary to console
        print("\n" + "=" * 80)
        print("REPORT GENERATION COMPLETE")
        print("=" * 80)
        print(summary_report)
        
        # 4. Save only mismatch results for easy action planning
        mismatch_results = [r for r in results if r['nameid_mismatch']]
        if mismatch_results:
            mismatch_output_file = f"nameid_mismatches_only_{TARGET_UPN.split('@')[0]}_{timestamp}.csv"
            with open(mismatch_output_file, 'w', newline='', encoding='utf-8') as csvfile:
                fieldnames = ['onedrive_owner', 'current_nameid', 'new_nameid', 'site_url', 'user_id']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                
                writer.writeheader()
                for result in mismatch_results:
                    writer.writerow({
                        'onedrive_owner': result['onedrive_owner'],
                        'current_nameid': result['current_nameid'],
                        'new_nameid': result['new_nameid'],
                        'site_url': result['site_url'],
                        'user_id': result['user_id']
                    })
            
            logger.info(f"🎯 Mismatch-only CSV saved to: {mismatch_output_file}")
            logger.info(f"🔧 {len(mismatch_results)} OneDrive(s) require user removal and re-addition")
        
    except Exception as e:
        logger.exception("Fatal error during report generation")

if __name__ == "__main__":
    main()
