from flask import Flask, render_template_string, request, jsonify
import requests
import json
import uuid
import base64
import time
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend
import logging
import os
import xml.etree.ElementTree as ET
import csv
from datetime import datetime

app = Flask(__name__)
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Configuration
CONFIG = {
    "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
    "tenant_name": "geekbyteonline.onmicrosoft.com",
    "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
    "client_secret": "CyG8Q~FYHuCMSyVmt4sNxt5IejrMc2c24Ziz4a.t",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    "scopes": {
        "graph": "https://graph.microsoft.com/.default",
        "sharepoint": "https://geekbyteonline.sharepoint.com/.default"
    }
}
CONFIG.update({
    "log_file": "admin_changes.csv",
    "log_fields": ["timestamp", "action", "source_upn", "destination_upn", "site_url", "status"]
})
def log_admin_change(action, source_upn, destination_upn, site_url, status="success"):
    """Log admin changes to CSV file"""
    try:
        # Create file with headers if it doesn't exist
        file_exists = os.path.exists(CONFIG['log_file'])
        
        with open(CONFIG['log_file'], 'a', newline='') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=CONFIG['log_fields'])
            
            if not file_exists:
                writer.writeheader()
                
            writer.writerow({
                "timestamp": datetime.now().isoformat(),
                "action": action,
                "source_upn": source_upn,
                "destination_upn": destination_upn,
                "site_url": site_url,
                "status": status
            })
            
    except Exception as e:
        logger.exception("Failed to log admin change to CSV")

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Site Admin Manager</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        :root {
            --primary-color: #0078d4;
            --danger-color: #d83b01;
            --success-color: #107c10;
            --warning-color: #ffaa44;
            --light-gray: #f3f2f1;
            --medium-gray: #e1dfdd;
            --dark-gray: #323130;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 20px;
            background-color: #f9f9f9;
            color: var(--dark-gray);
        }
        
        .container {
            max-width: 1000px;
            margin: 0 auto;
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        
        h1 {
            color: var(--primary-color);
            margin-top: 0;
            padding-bottom: 10px;
            border-bottom: 1px solid var(--medium-gray);
        }
        
        .input-group {
            margin-bottom: 20px;
        }
        
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: 600;
        }
        
        input[type="text"], input[type="url"], input[type="email"] {
            width: 100%;
            padding: 10px;
            border: 1px solid var(--medium-gray);
            border-radius: 4px;
            font-size: 16px;
            box-sizing: border-box;
        }
        
        .btn {
            padding: 10px 20px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
            font-weight: 600;
            transition: background-color 0.2s;
        }
        
        .btn-primary {
            background-color: var(--primary-color);
            color: white;
        }
        
        .btn-primary:hover {
            background-color: #106ebe;
        }
        
        .btn-danger {
            background-color: var(--danger-color);
            color: white;
        }
        
        .btn-danger:hover {
            background-color: #b32d00;
        }
        
        .status-message {
            padding: 15px;
            border-radius: 4px;
            margin: 20px 0;
            display: none;
        }
        
        .status-success {
            background-color: #dff6dd;
            color: var(--success-color);
            border-left: 4px solid var(--success-color);
            display: block;
        }
        
        .status-error {
            background-color: #fde7e9;
            color: var(--danger-color);
            border-left: 4px solid var(--danger-color);
            display: block;
        }
        
        .admin-list {
            margin-top: 30px;
            border: 1px solid var(--medium-gray);
            border-radius: 4px;
            overflow: hidden;
        }
        
        .admin-list-header {
            background-color: var(--light-gray);
            padding: 15px;
            font-weight: 600;
            border-bottom: 1px solid var(--medium-gray);
        }
        
        .admin-item {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 15px;
            border-bottom: 1px solid var(--medium-gray);
        }
        
        .admin-item:last-child {
            border-bottom: none;
        }
        
        .admin-info {
            flex-grow: 1;
        }
        
        .admin-name {
            font-weight: 600;
            margin-bottom: 5px;
        }
        
        .admin-email {
            color: #666;
            font-size: 14px;
        }
        
        .add-admin-form {
            margin-top: 30px;
            padding: 20px;
            background-color: var(--light-gray);
            border-radius: 4px;
        }
        
        .form-row {
            display: flex;
            gap: 10px;
        }
        
        .form-row input {
            flex-grow: 1;
        }
        
        .loading {
            display: inline-block;
            width: 20px;
            height: 20px;
            border: 3px solid rgba(255,255,255,0.3);
            border-radius: 50%;
            border-top-color: white;
            animation: spin 1s ease-in-out infinite;
        }
        
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Site Admin Manager</h1>
        
        <div class="input-group">
            <label for="site-input">Enter Site URL, OneDrive URL, or User UPN:</label>
            <input type="text" id="site-input" placeholder="https://geekbyteonline.sharepoint.com/sites/sitename OR user@domain.com">
            <button id="load-btn" class="btn btn-primary" onclick="loadAdmins()">Load Admins</button>
        </div>
        
        <div id="status" class="status-message"></div>
        
        <div id="site-info" style="display: none;">
            <h2 id="site-title"></h2>
            <p id="site-url"></p>
            
            <div class="admin-list">
                <div class="admin-list-header">Current Site Administrators</div>
                <div id="admins-container"></div>
            </div>
            
            <div class="add-admin-form">
                <h3>Add New Admin</h3>
                <div class="form-row">
                    <input type="email" id="new-admin-upn" placeholder="user@geekbyteonline.onmicrosoft.com">
                    <button id="add-admin-btn" class="btn btn-primary" onclick="addAdmin()">Add Admin</button>
                </div>
            </div>
        </div>
    </div>

    <script>
        let currentSiteUrl = null;
        let currentSiteType = null;
        
        function showStatus(message, isSuccess) {
            const statusDiv = document.getElementById('status');
            statusDiv.textContent = message;
            statusDiv.className = isSuccess ? 'status-message status-success' : 'status-message status-error';
        }
        
        function clearStatus() {
            document.getElementById('status').className = 'status-message';
        }
        
        function setLoading(button, isLoading) {
            if (isLoading) {
                button.innerHTML = '<span class="loading"></span> Processing...';
                button.disabled = true;
            } else {
                button.textContent = button.getAttribute('data-original-text');
                button.disabled = false;
            }
        }
        
        function loadAdmins() {
            const input = document.getElementById('site-input').value.trim();
            const loadBtn = document.getElementById('load-btn');
            
            if (!input) {
                showStatus('Please enter a site URL, OneDrive URL, or user UPN', false);
                return;
            }
            
            loadBtn.setAttribute('data-original-text', loadBtn.textContent);
            setLoading(loadBtn, true);
            clearStatus();
            
            fetch('/get_admins', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ input: input })
            })
            .then(res => res.json())
            .then(data => {
                setLoading(loadBtn, false);
                
                if (data.error) {
                    showStatus('Error: ' + data.error, false);
                    return;
                }
                
                currentSiteUrl = data.site_url;
                currentSiteType = data.site_type;
                
                document.getElementById('site-title').textContent = 
                    `${data.site_type} Administrators`;
                document.getElementById('site-url').textContent = data.site_url;
                
                const adminsContainer = document.getElementById('admins-container');
                adminsContainer.innerHTML = '';
                
                if (data.admins.length === 0) {
                    adminsContainer.innerHTML = '<div class="admin-item">No administrators found</div>';
                } else {
                    data.admins.forEach(admin => {
                        const adminDiv = document.createElement('div');
                        adminDiv.className = 'admin-item';
                        adminDiv.innerHTML = `
                            <div class="admin-info">
                                <div class="admin-name">${admin.title || admin.email || admin.login_name}</div>
                                <div class="admin-email">${admin.email || admin.login_name}</div>
                            </div>
                            <button class="btn btn-danger" onclick="removeAdmin('${admin.user_id}', '${admin.email || admin.login_name}')">Remove</button>
                        `;
                        adminsContainer.appendChild(adminDiv);
                    });
                }
                
                document.getElementById('site-info').style.display = 'block';
                showStatus(`Successfully loaded ${data.admins.length} admins`, true);
            })
            .catch(err => {
                setLoading(loadBtn, false);
                showStatus('Request failed: ' + err.message, false);
            });
        }
        
        function addAdmin() {
            const upn = document.getElementById('new-admin-upn').value.trim();
            const addBtn = document.getElementById('add-admin-btn');
            
            if (!upn || !currentSiteUrl) {
                showStatus('Please enter a valid UPN and load site admins first', false);
                return;
            }
            
            addBtn.setAttribute('data-original-text', addBtn.textContent);
            setLoading(addBtn, true);
            clearStatus();
            
            fetch('/manage_admin', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    site_url: currentSiteUrl,
                    upn: upn,
                    action: 'add'
                })
            })
            .then(res => res.json())
            .then(data => {
                setLoading(addBtn, false);
                
                if (data.error) {
                    showStatus('Error: ' + data.error, false);
                } else {
                    document.getElementById('new-admin-upn').value = '';
                    showStatus(data.message, true);
                    loadAdmins(); // Always refresh the admin list
                }
            })
            .catch(err => {
                setLoading(addBtn, false);
                showStatus('Request failed: ' + err.message, false);
            });
        }

        function removeAdmin(userId, userName) {        
            if (!confirm(`Are you sure you want to remove ${userName} from this site?`)) {
                return;
            }
            
            fetch('/manage_admin', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    site_url: currentSiteUrl,
                    user_id: userId,
                    upn: userName,
                    action: 'remove'
                })
            })
            .then(res => res.json())
            .then(data => {
                if (data.error) {
                    showStatus('Error: ' + data.error, false);
                } else {
                    showStatus(`Successfully removed ${userName}`, true);
                    loadAdmins(); // Always refresh the admin list
                }
            })
            .catch(err => {
                showStatus('Request failed: ' + err.message, false);
            });
        }
    </script>
</body>
</html>
"""

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

def parse_site_users_xml(xml_content):
    """Parse SharePoint site users XML and return all admins"""
    try:
        logger.debug("Parsing XML for site admins")
        
        # Parse XML
        root = ET.fromstring(xml_content)
        
        # Define namespaces
        namespaces = {
            'd': 'http://schemas.microsoft.com/ado/2007/08/dataservices',
            'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata',
            'atom': 'http://www.w3.org/2005/Atom'
        }
        
        # Find all entries
        entries = root.findall('.//atom:entry', namespaces)
        logger.debug(f"Found {len(entries)} entries in XML")
        
        admins = []
        
        for entry in entries:
            content = entry.find('atom:content', namespaces)
            if content is not None:
                properties = content.find('m:properties', namespaces)
                if properties is not None:
                    # Extract user details
                    user_id_elem = properties.find('d:Id', namespaces)
                    title_elem = properties.find('d:Title', namespaces)  
                    email_elem = properties.find('d:Email', namespaces)
                    login_name_elem = properties.find('d:LoginName', namespaces)
                    is_site_admin_elem = properties.find('d:IsSiteAdmin', namespaces)
                    
                    # Get values safely
                    user_id = user_id_elem.text if user_id_elem is not None else None
                    title = title_elem.text if title_elem is not None else None
                    email = email_elem.text if email_elem is not None else None
                    login_name = login_name_elem.text if login_name_elem is not None else None
                    is_site_admin = is_site_admin_elem.text == 'true' if is_site_admin_elem is not None else False
                    
                    if is_site_admin:
                        admins.append({
                            'user_id': user_id,
                            'title': title,
                            'email': email,
                            'login_name': login_name,
                            'is_site_admin': is_site_admin
                        })
        
        logger.info(f"Found {len(admins)} site administrators")
        return admins
        
    except Exception as e:
        logger.exception("Failed to parse site users XML")
        raise Exception(f"Failed to parse XML response: {str(e)}")
        
        
def ensure_user(site_url, token, request_digest, user_upn):
    """Ensure user exists on site and return user ID"""
    try:
        ensure_url = f"{site_url}_api/web/ensureuser('{user_upn}')"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": request_digest
        }
        
        response = requests.post(ensure_url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            return data['d']['Id']
        else:
            logger.error(f"Failed to ensure user {user_upn}: {response.text}")
            return None
    except Exception as e:
        logger.exception(f"Error ensuring user {user_upn}")
        return None
        
        
def get_site_admins(site_url):
    """Get all site administrators for a SharePoint site/OneDrive"""
    try:
        # Ensure site URL ends properly
        if not site_url.endswith('/'):
            site_url += '/'
        
        # Construct SharePoint API URL
        site_users_url = f"{site_url}_api/web/siteusers"
        logger.info(f"Fetching site admins from URL: {site_users_url}")
        
        # Get SharePoint token
        sharepoint_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sharepoint_token:
            sharepoint_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Call SharePoint API to get site users
        sharepoint_headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/xml"
        }
        
        logger.info(f"Making request to: {site_users_url}")
        site_users_response = requests.get(site_users_url, headers=sharepoint_headers)
        
        if site_users_response.status_code != 200:
            logger.error(f"SharePoint API failed: {site_users_response.status_code} - {site_users_response.text}")
            raise Exception(f"Failed to get site users from {site_users_url}: {site_users_response.text}")
        
        # Parse XML and find all admins
        admins = parse_site_users_xml(site_users_response.text)
        logger.info(f"Found admins: {[a['email'] or a['login_name'] for a in admins]}")
        return admins
        
    except Exception as e:
        logger.exception(f"Failed to get site admins from {site_url}")
        raise

def resolve_site_url(input_str):
    """Resolve the actual site URL from various input types"""
    try:
        # If it's already a URL
        if input_str.startswith('http'):
            if "my.sharepoint.com" in input_str.lower():
                # OneDrive URL
                if input_str.endswith('/Documents'):
                    return input_str[:-10]
                return input_str
            else:
                # SharePoint site URL
                return input_str
        
        # If it's a UPN (contains @)
        elif '@' in input_str:
            return get_onedrive_url(input_str)
        
        # Otherwise assume it's a site name
        else:
            if not input_str.startswith('sites/'):
                input_str = f"sites/{input_str}"
            return f"https://{CONFIG['tenant_name'].split('.')[0]}.sharepoint.com/{input_str}"
            
    except Exception as e:
        logger.exception("Failed to resolve site URL")
        raise

def get_request_digest(site_url, token):
    """Get request digest for SharePoint API calls"""
    try:
        if not site_url.endswith('/'):
            site_url += '/'
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=verbose"
        }
        
        response = requests.post(
            f"{site_url}_api/contextinfo",
            headers=headers
        )
        
        if response.status_code == 200:
            return response.json()['d']['GetContextWebInformation']['FormDigestValue']
        else:
            raise Exception(f"Failed to get request digest: {response.text}")
    except Exception as e:
        logger.exception("Failed to get request digest")
        raise

def add_user_as_admin(site_url, upn):
    """Add a user as admin to a SharePoint site/OneDrive"""
    try:
        if not site_url.endswith('/'):
            site_url += '/'
        
        # Get SharePoint token
        sharepoint_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sharepoint_token:
            sharepoint_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Get request digest
        request_digest = get_request_digest(site_url, sharepoint_token)
        headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": request_digest
        }

        # Step 1: Ensure user exists on site
        ensure_url = f"{site_url}_api/web/ensureuser"
        ensure_data = {
            'logonName': upn
        }
        
        logger.info(f"Ensuring user exists: {upn}")
        ensure_response = requests.post(
            ensure_url,
            headers=headers,
            json=ensure_data
        )
        
        if ensure_response.status_code != 200:
            raise Exception(f"Failed to ensure user {upn}: {ensure_response.text}")
        
        user_info = ensure_response.json()['d']
        user_id = user_info['Id']
        
        # Step 2: Set user as admin
        update_url = f"{site_url}_api/web/siteusers/getbyid({user_id})"
        update_headers = headers.copy()
        update_headers.update({
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
        })
        
        update_data = {
            '__metadata': {'type': 'SP.User'},
            'IsSiteAdmin': True
        }
        
        logger.info(f"Setting admin privileges for user ID: {user_id}")
        update_response = requests.post(
            update_url,
            headers=update_headers,
            json=update_data
        )
        
        if update_response.status_code not in [200, 204]:
            raise Exception(f"Failed to set admin privileges: {update_response.text}")
        
        # Step 3: Verify admin status
        verify_url = f"{site_url}_api/web/siteusers/getbyid({user_id})"
        verify_response = requests.get(verify_url, headers=headers)
        
        if verify_response.status_code != 200:
            logger.warning(f"Could not verify admin status for {upn}, but operation completed")
            return True
            
        user_data = verify_response.json()['d']
        if not user_data.get('IsSiteAdmin', False):
            logger.error(f"Admin status verification failed for {upn}")
            raise Exception(f"User {upn} was added but not set as admin")
        
        logger.info(f"Successfully added {upn} as admin to {site_url}")
        return True
        
    except Exception as e:
        logger.exception(f"Failed to add {upn} as admin to {site_url}")
        raise
def remove_admin_privileges(site_url, user_id, user_upn=None):
    """Remove a user from a SharePoint site/OneDrive"""
    try:
        if not site_url.endswith('/'):
            site_url += '/'
        
        # Get SharePoint token
        sharepoint_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sharepoint_token:
            sharepoint_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
        if not sharepoint_token:
            raise Exception("Failed to obtain SharePoint access token")
        
        # Get request digest
        request_digest = get_request_digest(site_url, sharepoint_token)
        headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": request_digest
        }
        
        # Step 1: First remove admin privileges (set IsSiteAdmin=False)
        update_url = f"{site_url}_api/web/siteusers/getbyid({user_id})"
        update_headers = headers.copy()
        update_headers.update({
            "IF-MATCH": "*",
            "X-HTTP-Method": "MERGE"
        })
        
        update_data = {
            '__metadata': {'type': 'SP.User'},
            'IsSiteAdmin': False
        }
        
        logger.info(f"Removing admin privileges for user ID: {user_id}")
        update_response = requests.post(
            update_url,
            headers=update_headers,
            json=update_data
        )
        
        if update_response.status_code not in [200, 204]:
            raise Exception(f"Failed to remove admin privileges: {update_response.text}")
        
        # Step 2: Remove user from site
        remove_url = f"{site_url}_api/web/siteusers/removebyid({user_id})"
        logger.info(f"Removing user ID {user_id} from site")
        remove_response = requests.post(
            remove_url,
            headers=headers
        )
        
        if remove_response.status_code not in [200, 204]:
            error_data = remove_response.json()
            error_msg = error_data.get('error', {}).get('message', {}).get('value', 'Unknown error')
            
            # If we can't remove the user, at least we've removed their admin privileges
            if "cannot delete the owners" in error_msg:
                logger.warning(f"Couldn't remove user {user_id} but admin privileges were revoked")
                return True
            raise Exception(f"Failed to remove user: {error_msg}")
        
        logger.info(f"Successfully removed user {user_id} from {site_url}")
        return True
        
    except Exception as e:
        logger.exception(f"Failed to remove user {user_id} from {site_url}")
        raise
        
@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/get_admins', methods=['POST'])
def get_admins():
    try:
        data = request.json
        input_str = data.get('input', '').strip()
        
        logger.info(f"Getting admins for input: {input_str}")
        
        if not input_str:
            return jsonify({"error": "Input is required"}), 400
        
        # Resolve the site URL from various input types
        site_url = resolve_site_url(input_str)
        logger.info(f"Resolved site URL: {site_url}")
        
        # Determine site type
        if "my.sharepoint.com" in site_url.lower():
            site_type = "OneDrive"
        else:
            site_type = "SharePoint Site"
        
        # Get site admins
        admins = get_site_admins(site_url)
        
        logger.info(f"Found {len(admins)} admins for site {site_url}")
        
        return jsonify({
            "success": True,
            "site_url": site_url,
            "site_type": site_type,
            "admins": admins
        })
        
    except Exception as e:
        logger.exception("Error occurred during admin listing")
        return jsonify({"error": str(e)}), 500
@app.route('/manage_admin', methods=['POST'])
def manage_admin():
    try:
        data = request.json
        site_url = data.get('site_url', '').strip()
        action = data.get('action', '').strip()
        upn = data.get('upn', '').strip()
        user_id = data.get('user_id', '').strip()
        
        # Get the current user (source UPN) from the request
        # You might need to adjust this based on your authentication
        source_upn = "system"  # Default if not available
        if request.headers.get('X-User-Email'):
            source_upn = request.headers['X-User-Email']
        
        if not site_url or action not in ['add', 'remove']:
            log_admin_change(action, source_upn, upn, site_url, "invalid parameters")
            return jsonify({"error": "Invalid parameters provided"}), 400
        
        if action == 'add' and not upn:
            log_admin_change(action, source_upn, upn, site_url, "missing UPN")
            return jsonify({"error": "User UPN is required for add operations"}), 400
        if action == 'remove' and not user_id:
            log_admin_change(action, source_upn, upn, site_url, "missing user ID")
            return jsonify({"error": "User ID is required for remove operations"}), 400
        
        # Perform the action
        if action == 'add':
            try:
                add_user_as_admin(site_url, upn)
                log_admin_change(action, source_upn, upn, site_url, "success")
                return jsonify({
                    "success": True,
                    "message": f"Successfully added {upn} as admin",
                    "user_id": "",
                    "site_url": site_url,
                    "action": action
                })
            except Exception as e:
                log_admin_change(action, source_upn, upn, site_url, f"error: {str(e)}")
                raise
        else:
            try:
                remove_admin_privileges(site_url, user_id, upn)
                log_admin_change(action, source_upn, upn, site_url, "success")
                return jsonify({
                    "success": True,
                    "message": f"Successfully removed user",
                    "user_id": user_id,
                    "site_url": site_url,
                    "action": action
                })
            except Exception as e:
                log_admin_change(action, source_upn, upn, site_url, f"error: {str(e)}")
                raise
        
    except Exception as e:
        logger.exception("Error occurred during admin management")
        return jsonify({"error": str(e)}), 500
        
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=6005, debug=True)