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
                    console.error('Detailed error:', data.error);
                } else {
                    document.getElementById('new-admin-upn').value = '';
                    showStatus(`Successfully added ${upn} as admin`, true);
                    loadAdmins(); // Refresh the admin list
                }
            })
            .catch(err => {
                setLoading(addBtn, false);
                showStatus('Request failed: ' + err.message, false);
                console.error('Full error:', err);
            });
        }
        
        function removeAdmin(userId, userName) {
            if (!confirm(`Are you sure you want to remove admin privileges from ${userName}?`)) {
                return;
            }
            
            if (!userId || !currentSiteUrl) {
                showStatus('Invalid parameters for removal', false);
                return;
            }
            
            fetch('/manage_admin', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    site_url: currentSiteUrl,
                    user_id: userId,
                    action: 'remove'
                })
            })
            .then(res => res.json())
            .then(data => {
                if (data.error) {
                    showStatus('Error: ' + data.error, false);
                } else {
                    showStatus(`Successfully removed admin privileges from ${userName}`, true);
                    loadAdmins(); // Refresh the admin list
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
                    login_name_elem = properties.fi
