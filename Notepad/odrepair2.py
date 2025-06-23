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
    <title>SharePoint Site User Remover</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1000px;
            margin: 0 auto;
            background: rgba(255, 255, 255, 0.95);
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            backdrop-filter: blur(10px);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5rem;
            margin-bottom: 10px;
            font-weight: 300;
        }
        
        .header p {
            opacity: 0.9;
            font-size: 1.1rem;
        }
        
        .warning-banner {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            color: #856404;
            padding: 15px;
            margin: 20px 30px;
            border-radius: 8px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .main-content {
            padding: 30px;
        }
        
        .search-section {
            background: white;
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 25px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        }
        
        .form-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 20px;
            margin-bottom: 20px;
        }
        
        .form-group {
            margin-bottom: 20px;
        }
        
        .form-group label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #495057;
            font-size: 1.1rem;
        }
        
        .form-control {
            width: 100%;
            padding: 15px;
            border: 2px solid #e9ecef;
            border-radius: 8px;
            font-size: 16px;
            transition: all 0.3s ease;
        }
        
        .form-control:focus {
            outline: none;
            border-color: #dc3545;
            box-shadow: 0 0 0 3px rgba(220, 53, 69, 0.1);
        }
        
        .btn {
            padding: 15px 30px;
            border: none;
            border-radius: 8px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
            display: inline-flex;
            align-items: center;
            gap: 10px;
            text-decoration: none;
        }
        
        .btn-danger {
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
            color: white;
        }
        
        .btn-danger:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 16px rgba(220, 53, 69, 0.3);
        }
        
        .btn-danger:disabled {
            opacity: 0.6;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        
        .btn-secondary {
            background: linear-gradient(135deg, #6c757d 0%, #5a6268 100%);
            color: white;
        }
        
        .btn-secondary:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 16px rgba(108, 117, 125, 0.3);
        }
        
        .btn-block {
            width: 100%;
            justify-content: center;
        }
        
        .btn-row {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
        }
        
        .status-section {
            display: flex;
            align-items: center;
            gap: 10px;
            margin: 20px 0;
            padding: 15px;
            border-radius: 8px;
            font-weight: 500;
        }
        
        .status-ready {
            background: #d1ecf1;
            color: #0c5460;
        }
        
        .status-loading {
            background: #fff3cd;
            color: #856404;
        }
        
        .status-success {
            background: #d4edda;
            color: #155724;
        }
        
        .status-error {
            background: #f8d7da;
            color: #721c24;
        }
        
        .results-section {
            background: white;
            border-radius: 12px;
            padding: 25px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
            margin-bottom: 25px;
        }
        
        .user-info {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
        }
        
        .user-info h3 {
            color: #dc3545;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 15px;
        }
        
        .info-item {
            display: flex;
            flex-direction: column;
            gap: 5px;
        }
        
        .info-label {
            font-weight: 600;
            color: #495057;
            font-size: 0.9rem;
        }
        
        .info-value {
            font-size: 0.95rem;
            color: #6c757d;
            word-break: break-all;
        }
        
        .removal-confirmation {
            background: #fff5f5;
            border: 2px solid #fed7d7;
            border-radius: 8px;
            padding: 20px;
            margin-top: 20px;
        }
        
        .removal-confirmation h4 {
            color: #dc3545;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        @media (max-width: 768px) {
            .container {
                margin: 10px;
                border-radius: 15px;
            }
            
            .main-content {
                padding: 20px;
            }
            
            .header h1 {
                font-size: 2rem;
            }
            
            .form-row {
                grid-template-columns: 1fr;
            }
            
            .btn-row {
                grid-template-columns: 1fr;
            }
            
            .info-grid {
                grid-template-columns: 1fr;
            }
        }
        
        .spinner {
            display: inline-block;
            width: 20px;
            height: 20px;
            border: 3px solid rgba(255,255,255,.3);
            border-radius: 50%;
            border-top-color: #fff;
            animation: spin 1s ease-in-out infinite;
        }
        
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
        
        .help-text {
            font-size: 0.9rem;
            color: #6c757d;
            margin-top: 5px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i class="fas fa-user-minus"></i> SharePoint Site User Remover</h1>
            <p>Remove users from SharePoint sites</p>
        </div>
        
        <div class="warning-banner">
            <i class="fas fa-exclamation-triangle"></i>
            <strong>Warning:</strong> This action will permanently remove the user from the specified site. Use with caution.
        </div>
        
        <div class="main-content">
            <div class="search-section">
                <h3 style="margin-bottom: 20px; color: #dc3545;">
                    <i class="fas fa-user-times"></i> User Removal
                </h3>
                
                <div class="form-group">
                    <label for="removal-type">Removal Type:</label>
                    <select id="removal-type" class="form-control" onchange="toggleRemovalType()">
                        <option value="site">Remove from SharePoint Site</option>
                        <option value="onedrive">Remove from User's OneDrive</option>
                    </select>
                    <div class="help-text">Choose whether to remove from a site or from someone's OneDrive</div>
                </div>
                
                <div class="form-row">
                    <div class="form-group">
                        <label for="user-upn">User to Remove (UPN):</label>
                        <input type="email" id="user-upn" class="form-control" 
                               placeholder="user@geekbyteonline.onmicrosoft.com">
                        <div class="help-text">Enter the email/UPN of the user to remove</div>
                    </div>
                    
                    <div class="form-group" id="site-url-group">
                        <label for="site-url">SharePoint Site URL:</label>
                        <input type="url" id="site-url" class="form-control" 
                               placeholder="https://geekbyteonline.sharepoint.com/sites/sitename">
                        <div class="help-text">Enter the full URL of the SharePoint site</div>
                    </div>
                    
                    <div class="form-group" id="onedrive-owner-group" style="display: none;">
                        <label for="onedrive-owner-upn">OneDrive Owner (UPN):</label>
                        <input type="email" id="onedrive-owner-upn" class="form-control" 
                               placeholder="owner@geekbyteonline.onmicrosoft.com">
                        <div class="help-text">Enter the email/UPN of the OneDrive owner</div>
                    </div>
                </div>
                
                <div class="btn-row">
                    <button id="find-user-btn" class="btn btn-secondary" onclick="findUser()">
                        <i class="fas fa-search"></i> Find User
                    </button>
                    <button id="remove-user-btn" class="btn btn-danger" onclick="removeUser()" disabled>
                        <i class="fas fa-user-minus"></i> Remove User
                    </button>
                </div>
            </div>
            
            <div id="status" class="status-section status-ready">
                <i class="fas fa-info-circle"></i> Enter user UPN and site URL to start
            </div>
            
            <div id="results-section" class="results-section" style="display: none;">
                <div id="user-info" class="user-info"></div>
                <div id="confirmation-section" class="removal-confirmation" style="display: none;"></div>
            </div>
        </div>
    </div>

    <script>
        let currentUserId = null;
        let currentSiteUrl = null;
        let currentRemovalType = 'site';
        
        function toggleRemovalType() {
            const removalType = document.getElementById('removal-type').value;
            const siteUrlGroup = document.getElementById('site-url-group');
            const onedriveOwnerGroup = document.getElementById('onedrive-owner-group');
            
            currentRemovalType = removalType;
            
            if (removalType === 'site') {
                siteUrlGroup.style.display = 'block';
                onedriveOwnerGroup.style.display = 'none';
            } else {
                siteUrlGroup.style.display = 'none';
                onedriveOwnerGroup.style.display = 'block';
            }
            
            // Reset form state
            currentUserId = null;
            currentSiteUrl = null;
            document.getElementById('remove-user-btn').disabled = true;
            document.getElementById('results-section').style.display = 'none';
            setStatus('ready', 'Enter user details to start');
        }
        
        function setStatus(type, message) {
            const statusDiv = document.getElementById('status');
            statusDiv.className = `status-section status-${type}`;
            
            let icon = 'fas fa-info-circle';
            if (type === 'loading') icon = 'fas fa-spinner fa-spin';
            else if (type === 'success') icon = 'fas fa-check-circle';
            else if (type === 'error') icon = 'fas fa-exclamation-circle';
            
            statusDiv.innerHTML = `<i class="${icon}"></i> ${message}`;
        }
        
        function findUser() {
            const upn = document.getElementById('user-upn').value.trim();
            const removalType = document.getElementById('removal-type').value;
            const findBtn = document.getElementById('find-user-btn');
            const removeBtn = document.getElementById('remove-user-btn');
            const resultsSection = document.getElementById('results-section');
            
            let siteUrl = '';
            let onedriveOwnerUpn = '';
            
            if (removalType === 'site') {
                siteUrl = document.getElementById('site-url').value.trim();
                if (!upn || !siteUrl) {
                    setStatus('error', 'Please enter both user UPN and site URL');
                    return;
                }
            } else {
                onedriveOwnerUpn = document.getElementById('onedrive-owner-upn').value.trim();
                if (!upn || !onedriveOwnerUpn) {
                    setStatus('error', 'Please enter both user UPN and OneDrive owner UPN');
                    return;
                }
            }
            
            // Disable buttons and show loading
            findBtn.disabled = true;
            removeBtn.disabled = true;
            findBtn.innerHTML = '<div class="spinner"></div> Finding User...';
            setStatus('loading', removalType === 'site' ? 'Searching for user on the specified site...' : 'Getting OneDrive URL and searching for user...');
            resultsSection.style.display = 'none';
            currentUserId = null;
            currentSiteUrl = null;
            
            const requestData = {
                upn: upn,
                removal_type: removalType
            };
            
            if (removalType === 'site') {
                requestData.site_url = siteUrl;
                currentSiteUrl = siteUrl;
            } else {
                requestData.onedrive_owner_upn = onedriveOwnerUpn;
            }
            
            fetch('/find_user', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(requestData)
            })
            .then(res => res.json())
            .then(data => {
                findBtn.disabled = false;
                findBtn.innerHTML = '<i class="fas fa-search"></i> Find User';
                
                if (data.error) {
                    setStatus('error', `Error: ${data.error}`);
                    resultsSection.style.display = 'none';
                    removeBtn.disabled = true;
                } else {
                    setStatus('success', removalType === 'site' ? 'User found on site' : 'User found on OneDrive');
                    displayUserInfo(data);
                    resultsSection.style.display = 'block';
                    removeBtn.disabled = false;
                    currentUserId = data.user_id;
                    currentSiteUrl = data.site_url;
                }
            })
            .catch(err => {
                findBtn.disabled = false;
                findBtn.innerHTML = '<i class="fas fa-search"></i> Find User';
                setStatus('error', 'Request failed: ' + err.message);
                resultsSection.style.display = 'none';
                removeBtn.disabled = true;
            });
        }
        
        function removeUser() {
            if (!currentUserId || !currentSiteUrl) {
                setStatus('error', 'Please find a user first');
                return;
            }
            
            if (!confirm('Are you sure you want to remove this user from the site? This action cannot be undone.')) {
                return;
            }
            
            const removeBtn = document.getElementById('remove-user-btn');
            const findBtn = document.getElementById('find-user-btn');
            
            // Disable buttons and show loading
            removeBtn.disabled = true;
            findBtn.disabled = true;
            removeBtn.innerHTML = '<div class="spinner"></div> Removing User...';
            setStatus('loading', 'Removing user from site...');
            
            fetch('/remove_user', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    user_id: currentUserId, 
                    site_url: currentSiteUrl 
                })
            })
            .then(res => res.json())
            .then(data => {
                removeBtn.disabled = false;
                findBtn.disabled = false;
                removeBtn.innerHTML = '<i class="fas fa-user-minus"></i> Remove User';
                
                if (data.error) {
                    setStatus('error', `Error: ${data.error}`);
                } else {
                    setStatus('success', 'User successfully removed from site');
                    displayRemovalConfirmation(data);
                    // Reset form
                    currentUserId = null;
                    removeBtn.disabled = true;
                }
            })
            .catch(err => {
                removeBtn.disabled = false;
                findBtn.disabled = false;
                removeBtn.innerHTML = '<i class="fas fa-user-minus"></i> Remove User';
                setStatus('error', 'Request failed: ' + err.message);
            });
        }
        
        function displayUserInfo(data) {
            const userInfoDiv = document.getElementById('user-info');
            
            userInfoDiv.innerHTML = `
                <h3><i class="fas fa-user"></i> User Found</h3>
                <div class="info-grid">
                    <div class="info-item">
                        <div class="info-label">Display Name</div>
                        <div class="info-value">${data.title || 'N/A'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">Email</div>
                        <div class="info-value">${data.email || 'N/A'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">User ID</div>
                        <div class="info-value">${data.user_id || 'N/A'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">Login Name</div>
                        <div class="info-value">${data.login_name || 'N/A'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">Site/OneDrive Type</div>
                        <div class="info-value">${data.site_type || 'SharePoint Site'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">Site URL</div>
                        <div class="info-value">${data.site_url || 'N/A'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">OneDrive Owner</div>
                        <div class="info-value">${data.onedrive_owner || 'N/A'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">Is Site Admin</div>
                        <div class="info-value">${data.is_site_admin ? 'Yes' : 'No'}</div>
                    </div>
                </div>
            `;
        }
        
        function displayRemovalConfirmation(data) {
            const confirmationDiv = document.getElementById('confirmation-section');
            
            confirmationDiv.innerHTML = `
                <h4><i class="fas fa-check-circle"></i> User Removed Successfully</h4>
                <p><strong>User ID ${data.user_id}</strong> has been successfully removed from the site.</p>
                <p><strong>Removal Time:</strong> ${new Date().toLocaleString()}</p>
            `;
            
            confirmationDiv.style.display = 'block';
        }
        
        // Allow Enter key to submit
        document.getElementById('user-upn').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                findUser();
            }
        });
        
        document.getElementById('site-url').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                findUser();
            }
        });
        
        document.getElementById('onedrive-owner-upn').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                findUser();
            }
        });
    </script>
</body>
</html>
"""

def get_token_with_certificate(scope):
    """Get access token using certificate-based authentication"""
    try:
        if not os.path.exists(CONFIG['certificate_path']) or not os.path.exists(CONFIG['private_key_path']):
            raise Exception("Certificate or private key file not found")
            
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
            return token_response.json()["access_token"]
        else:
            logger.error(f"Token request failed: {token_response.text}")
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
            return token_response.json()["access_token"]
        else:
            logger.error(f"Token request failed: {token_response.text}")
            return None
            
    except Exception as e:
        logger.exception("Client secret authentication failed")
        return None

def parse_site_users_xml(xml_content, target_upn):
    """Parse SharePoint site users XML and find specific user"""
    try:
        logger.debug(f"Parsing XML for user: {target_upn}")
        
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
                    
                    logger.debug(f"Found user - ID: {user_id}, Email: {email}, Login: {login_name}")
                    
                    # Check if this is the target user (with proper null checks)
                    email_match = email and email.lower() == target_upn.lower()
                    login_match = login_name and target_upn.lower() in login_name.lower()
                    
                    if email_match or login_match:
                        user_info = {
                            'user_id': user_id,
                            'title': title,
                            'email': email,
                            'login_name': login_name,
                            'is_site_admin': is_site_admin
                        }
                        logger.info(f"Found target user: {user_info}")
                        return user_info
        
        logger.warning(f"User {target_upn} not found in site users")
        return None
        
    except Exception as e:
        logger.exception("Failed to parse site users XML")
        raise Exception(f"Failed to parse XML response: {str(e)}")


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/find_user', methods=['POST'])
def find_user():
    try:
        data = request.json
        upn = data.get('upn', '').strip()
        removal_type = data.get('removal_type', 'site')
        
        if not upn:
            return jsonify({"error": "User UPN is required"}), 400
        
        site_url = ''
        site_type = 'SharePoint Site'
        onedrive_owner = 'N/A'
        
        if removal_type == 'site':
            site_url = data.get('site_url', '').strip()
            if not site_url:
                return jsonify({"error": "Site URL is required"}), 400
        else:  # onedrive
            onedrive_owner_upn = data.get('onedrive_owner_upn', '').strip()
            if not onedrive_owner_upn:
                return jsonify({"error": "OneDrive owner UPN is required"}), 400
            
            # Get Graph API token to fetch OneDrive URL
            graph_token = get_token_with_certificate(CONFIG['scopes']['graph'])
            if not graph_token:
                graph_token = get_token_with_secret(CONFIG['scopes']['graph'])
            
            if not graph_token:
                return jsonify({"error": "Failed to obtain Graph API access token"}), 500
            
            # Get OneDrive URL for the owner
            graph_headers = {
                "Authorization": f"Bearer {graph_token}",
                "Content-Type": "application/json"
            }
            
            onedrive_response = requests.get(
                f"https://graph.microsoft.com/v1.0/users/{onedrive_owner_upn}/drive?$select=webUrl",
                headers=graph_headers
            )
            
            if onedrive_response.status_code != 200:
                return jsonify({"error": f"OneDrive not found for owner: {onedrive_response.text}"}), 404
            
            onedrive_info = onedrive_response.json()
            onedrive_url = onedrive_info.get('webUrl', '')
            
            if not onedrive_url:
                return jsonify({"error": "OneDrive URL not found for owner"}), 404
            
            # Remove /Documents from the end if present
            if onedrive_url.endswith('/Documents'):
                site_url = onedrive_url[:-10]  # Remove /Documents
            else:
                site_url = onedrive_url
            
            site_type = 'OneDrive'
            onedrive_owner = onedrive_owner_upn
        
        # Ensure site URL ends properly
        if not site_url.endswith('/'):
            site_url += '/'
        
        # Construct SharePoint API URL
        site_users_url = f"{site_url}_api/web/siteusers"
        
        # Get SharePoint token
        sharepoint_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sharepoint_token:
            sharepoint_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
        if not sharepoint_token:
            return jsonify({"error": "Failed to obtain SharePoint access token"}), 500
        
        # Call SharePoint API to get site users
        sharepoint_headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/xml"
        }
        
        site_users_response = requests.get(site_users_url, headers=sharepoint_headers)
        
        if site_users_response.status_code != 200:
            return jsonify({"error": f"Failed to get site users: {site_users_response.text}"}), 500
        
        # Parse XML and find the specific user
        user_info = parse_site_users_xml(site_users_response.text, upn)
        
        if not user_info:
            return jsonify({"error": f"User not found on the specified {site_type.lower()}"}), 404
        
        # Add additional info to response
        user_info['site_url'] = site_url
        user_info['site_type'] = site_type
        user_info['onedrive_owner'] = onedrive_owner
        
        return jsonify(user_info)
        
    except Exception as e:
        logger.exception("Error occurred during user search")
        return jsonify({"error": str(e)}), 500

@app.route('/remove_user', methods=['POST'])
def remove_user():
    try:
        data = request.json
        user_id = data.get('user_id', '').strip()
        site_url = data.get('site_url', '').strip()
        
        if not user_id or not site_url:
            return jsonify({"error": "Both user ID and site URL are required"}), 400
        
        # Ensure site URL ends properly
        if not site_url.endswith('/'):
            site_url += '/'
        
        # Construct SharePoint API URL for user removal
        remove_user_url = f"{site_url}_api/web/siteusers/removebyid('{user_id}')"
        
        # Get SharePoint token
        sharepoint_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sharepoint_token:
            sharepoint_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
        if not sharepoint_token:
            return jsonify({"error": "Failed to obtain SharePoint access token"}), 500
        
        # Call SharePoint API to remove user
        sharepoint_headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": "",  # This might be needed for some environments
            "X-HTTP-Method": "POST"
        }
        
        remove_response = requests.post(remove_user_url, headers=sharepoint_headers)
        
        if remove_response.status_code not in [200, 204]:
            return jsonify({"error": f"Failed to remove user: {remove_response.text}"}), 500
        
        return jsonify({
            "success": True,
            "message": "User successfully removed from site",
            "user_id": user_id,
            "site_url": site_url
        })
        
    except Exception as e:
        logger.exception("Error occurred during user removal")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5002, debug=True)