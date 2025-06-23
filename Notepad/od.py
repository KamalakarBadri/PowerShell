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
    "repair_account": "edit@geekbyte.online",
    "scopes": {
        "graph": "https://graph.microsoft.com/.default",
        "sharepoint": "https://geekbyteonline.sharepoint.com/.default"
    }
}

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>OneDrive Repair Tool</title>
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
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
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
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
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
        
        .info-banner {
            background: #d1ecf1;
            border: 1px solid #bee5eb;
            color: #0c5460;
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
        
        .repair-section {
            background: white;
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 25px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
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
            border-color: #28a745;
            box-shadow: 0 0 0 3px rgba(40, 167, 69, 0.1);
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
        
        .btn-success {
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
            color: white;
        }
        
        .btn-success:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 16px rgba(40, 167, 69, 0.3);
        }
        
        .btn-success:disabled {
            opacity: 0.6;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        
        .btn-info {
            background: linear-gradient(135deg, #17a2b8 0%, #138496 100%);
            color: white;
        }
        
        .btn-info:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 16px rgba(23, 162, 184, 0.3);
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
        
        .progress-container {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 20px;
            margin: 20px 0;
        }
        
        .progress-step {
            display: flex;
            align-items: center;
            padding: 10px 0;
            border-bottom: 1px solid #e9ecef;
        }
        
        .progress-step:last-child {
            border-bottom: none;
        }
        
        .step-icon {
            width: 30px;
            height: 30px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            margin-right: 15px;
            font-size: 12px;
        }
        
        .step-pending {
            background: #e9ecef;
            color: #6c757d;
        }
        
        .step-running {
            background: #fff3cd;
            color: #856404;
        }
        
        .step-success {
            background: #d4edda;
            color: #155724;
        }
        
        .step-error {
            background: #f8d7da;
            color: #721c24;
        }
        
        .step-text {
            flex: 1;
            font-weight: 500;
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
            
            .btn-row {
                grid-template-columns: 1fr;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i class="fas fa-tools"></i> OneDrive Repair Tool</h1>
            <p>Fix OneDrive access issues due to user PUID mismatches</p>
        </div>
        
        <div class="info-banner">
            <i class="fas fa-info-circle"></i>
            <strong>Info:</strong> This tool repairs OneDrive access by resetting user permissions and fixing PUID mismatches.
        </div>
        
        <div class="main-content">
            <div class="repair-section">
                <h3 style="margin-bottom: 20px; color: #28a745;">
                    <i class="fas fa-user-cog"></i> OneDrive Repair
                </h3>
                
                <div class="form-group">
                    <label for="user-upn">User UPN to Repair:</label>
                    <input type="email" id="user-upn" class="form-control" 
                           placeholder="user@geekbyteonline.onmicrosoft.com">
                    <div class="help-text">Enter the email/UPN of the user whose OneDrive needs repair</div>
                </div>
                
                <div class="btn-row">
                    <button id="get-onedrive-btn" class="btn btn-info" onclick="getOneDriveInfo()">
                        <i class="fas fa-search"></i> Get OneDrive Info
                    </button>
                    <button id="repair-onedrive-btn" class="btn btn-success" onclick="repairOneDrive()" disabled>
                        <i class="fas fa-tools"></i> Repair OneDrive
                    </button>
                </div>
            </div>
            
            <div id="status" class="status-section status-ready">
                <i class="fas fa-info-circle"></i> Enter user UPN to start OneDrive repair
            </div>
            
            <div id="results-section" class="results-section" style="display: none;">
                <div id="onedrive-info"></div>
                <div id="progress-container" class="progress-container" style="display: none;">
                    <h4 style="margin-bottom: 15px; color: #28a745;">
                        <i class="fas fa-cogs"></i> Repair Progress
                    </h4>
                    <div id="progress-steps"></div>
                </div>
            </div>
        </div>
    </div>

    <script>
        let currentOneDriveUrl = null;
        let currentUserUpn = null;
        
        function setStatus(type, message) {
            const statusDiv = document.getElementById('status');
            statusDiv.className = `status-section status-${type}`;
            
            let icon = 'fas fa-info-circle';
            if (type === 'loading') icon = 'fas fa-spinner fa-spin';
            else if (type === 'success') icon = 'fas fa-check-circle';
            else if (type === 'error') icon = 'fas fa-exclamation-circle';
            
            statusDiv.innerHTML = `<i class="${icon}"></i> ${message}`;
        }
        
        function getOneDriveInfo() {
            const upn = document.getElementById('user-upn').value.trim();
            const getBtn = document.getElementById('get-onedrive-btn');
            const repairBtn = document.getElementById('repair-onedrive-btn');
            const resultsSection = document.getElementById('results-section');
            
            if (!upn) {
                setStatus('error', 'Please enter a user UPN');
                return;
            }
            
            // Disable buttons and show loading
            getBtn.disabled = true;
            repairBtn.disabled = true;
            getBtn.innerHTML = '<div class="spinner"></div> Getting OneDrive Info...';
            setStatus('loading', 'Getting OneDrive information...');
            resultsSection.style.display = 'none';
            currentOneDriveUrl = null;
            currentUserUpn = upn;
            
            fetch('/get_onedrive_info', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ upn: upn })
            })
            .then(res => res.json())
            .then(data => {
                getBtn.disabled = false;
                getBtn.innerHTML = '<i class="fas fa-search"></i> Get OneDrive Info';
                
                if (data.error) {
                    setStatus('error', `Error: ${data.error}`);
                    resultsSection.style.display = 'none';
                    repairBtn.disabled = true;
                } else {
                    setStatus('success', 'OneDrive information retrieved successfully');
                    displayOneDriveInfo(data);
                    resultsSection.style.display = 'block';
                    repairBtn.disabled = false;
                    currentOneDriveUrl = data.onedrive_url;
                }
            })
            .catch(err => {
                getBtn.disabled = false;
                getBtn.innerHTML = '<i class="fas fa-search"></i> Get OneDrive Info';
                setStatus('error', 'Request failed: ' + err.message);
                resultsSection.style.display = 'none';
                repairBtn.disabled = true;
            });
        }
        
        function repairOneDrive() {
            if (!currentOneDriveUrl || !currentUserUpn) {
                setStatus('error', 'Please get OneDrive info first');
                return;
            }
            
            if (!confirm('Are you sure you want to repair this OneDrive? This will reset user permissions and may take several minutes.')) {
                return;
            }
            
            const repairBtn = document.getElementById('repair-onedrive-btn');
            const getBtn = document.getElementById('get-onedrive-btn');
            const progressContainer = document.getElementById('progress-container');
            
            // Disable buttons and show loading
            repairBtn.disabled = true;
            getBtn.disabled = true;
            repairBtn.innerHTML = '<div class="spinner"></div> Repairing OneDrive...';
            setStatus('loading', 'Starting OneDrive repair process...');
            progressContainer.style.display = 'block';
            
            // Initialize progress steps
            initializeProgressSteps();
            
            fetch('/repair_onedrive', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    upn: currentUserUpn,
                    onedrive_url: currentOneDriveUrl 
                })
            })
            .then(res => res.json())
            .then(data => {
                repairBtn.disabled = false;
                getBtn.disabled = false;
                repairBtn.innerHTML = '<i class="fas fa-tools"></i> Repair OneDrive';
                
                if (data.error) {
                    setStatus('error', `Error: ${data.error}`);
                    updateProgressStep(data.failed_step || 0, 'error', data.error);
                } else {
                    setStatus('success', 'OneDrive repair completed successfully');
                    updateAllStepsSuccess();
                }
            })
            .catch(err => {
                repairBtn.disabled = false;
                getBtn.disabled = false;
                repairBtn.innerHTML = '<i class="fas fa-tools"></i> Repair OneDrive';
                setStatus('error', 'Request failed: ' + err.message);
                updateProgressStep(0, 'error', err.message);
            });
        }
        
        function displayOneDriveInfo(data) {
            const oneDriveInfoDiv = document.getElementById('onedrive-info');
            
            oneDriveInfoDiv.innerHTML = `
                <h3 style="color: #28a745; margin-bottom: 15px;">
                    <i class="fas fa-cloud"></i> OneDrive Information
                </h3>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 15px;">
                    <div style="display: flex; flex-direction: column; gap: 5px;">
                        <div style="font-weight: 600; color: #495057; font-size: 0.9rem;">User UPN</div>
                        <div style="font-size: 0.95rem; color: #6c757d; word-break: break-all;">${data.upn || 'N/A'}</div>
                    </div>
                    <div style="display: flex; flex-direction: column; gap: 5px;">
                        <div style="font-weight: 600; color: #495057; font-size: 0.9rem;">OneDrive URL</div>
                        <div style="font-size: 0.95rem; color: #6c757d; word-break: break-all;">${data.onedrive_url || 'N/A'}</div>
                    </div>
                    <div style="display: flex; flex-direction: column; gap: 5px;">
                        <div style="font-weight: 600; color: #495057; font-size: 0.9rem;">Site Title</div>
                        <div style="font-size: 0.95rem; color: #6c757d;">${data.site_title || 'N/A'}</div>
                    </div>
                    <div style="display: flex; flex-direction: column; gap: 5px;">
                        <div style="font-weight: 600; color: #495057; font-size: 0.9rem;">Owner</div>
                        <div style="font-size: 0.95rem; color: #6c757d;">${data.owner || 'N/A'}</div>
                    </div>
                </div>
            `;
        }
        
        function initializeProgressSteps() {
            const steps = [
                "Get OneDrive URL",
                "Add repair account as user", 
                "Make repair account site admin",
                "Remove original user admin rights",
                "Remove original user from site",
                "Re-add original user to site",
                "Make original user site admin",
                "Remove repair account admin rights",
                "Remove repair account from site"
            ];
            
            const progressSteps = document.getElementById('progress-steps');
            progressSteps.innerHTML = '';
            
            steps.forEach((step, index) => {
                const stepDiv = document.createElement('div');
                stepDiv.className = 'progress-step';
                stepDiv.id = `step-${index}`;
                stepDiv.innerHTML = `
                    <div class="step-icon step-pending">
                        <i class="fas fa-clock"></i>
                    </div>
                    <div class="step-text">${step}</div>
                `;
                progressSteps.appendChild(stepDiv);
            });
        }
        
        function updateProgressStep(stepIndex, status, message = '') {
            const stepDiv = document.getElementById(`step-${stepIndex}`);
            if (!stepDiv) return;
            
            const iconDiv = stepDiv.querySelector('.step-icon');
            const textDiv = stepDiv.querySelector('.step-text');
            
            iconDiv.className = `step-icon step-${status}`;
            
            let icon = 'fas fa-clock';
            if (status === 'running') icon = 'fas fa-spinner fa-spin';
            else if (status === 'success') icon = 'fas fa-check';
            else if (status === 'error') icon = 'fas fa-times';
            
            iconDiv.innerHTML = `<i class="${icon}"></i>`;
            
            if (message && status === 'error') {
                textDiv.innerHTML += `<br><small style="color: #721c24;">${message}</small>`;
            }
        }
        
        function updateAllStepsSuccess() {
            for (let i = 0; i < 9; i++) {
                updateProgressStep(i, 'success');
            }
        }
        
        // Allow Enter key to submit
        document.getElementById('user-upn').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                getOneDriveInfo();
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

def set_site_admin(site_url, token, request_digest, user_id, is_admin=True):
    """Set or remove site admin permissions for user"""
    try:
        admin_url = f"{site_url}_api/web/getuserbyid('{user_id}')"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": request_digest,
            "X-HTTP-Method": "MERGE"
        }
        
        body = {
            "__metadata": {"type": "SP.User"},
            "IsSiteAdmin": is_admin
        }
        
        response = requests.post(admin_url, headers=headers, json=body)
        if response.status_code in [200, 204]:
            logger.info(f"Successfully set site admin to {is_admin} for user ID {user_id}")
            return True
        else:
            logger.error(f"Failed to set site admin for user ID {user_id}: {response.text}")
            return False
    except Exception as e:
        logger.exception(f"Error setting site admin for user ID {user_id}")
        return False

def remove_user_by_id(site_url, token, request_digest, user_id):
    """Remove user from site by ID"""
    try:
        remove_url = f"{site_url}_api/web/siteusers/removebyid('{user_id}')"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": request_digest
        }
        
        response = requests.post(remove_url, headers=headers)
        if response.status_code in [200, 204]:
            logger.info(f"Successfully removed user ID {user_id} from site")
            return True
        else:
            logger.error(f"Failed to remove user ID {user_id}: {response.text}")
            return False
    except Exception as e:
        logger.exception(f"Error removing user ID {user_id}")
        return False

def get_user_id_from_site(site_url, token, user_upn):
    """Get user ID from site users"""
    try:
        users_url = f"{site_url}_api/web/siteusers"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=verbose"
        }
        
        response = requests.get(users_url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            users = data.get('d', {}).get('results', [])
            
            for user in users:
                if (user.get('Email', '').lower() == user_upn.lower() or 
                    user_upn.lower() in user.get('LoginName', '').lower()):
                    return user.get('Id')
        
        return None
    except Exception as e:
        logger.exception("Error getting user ID from site")
        return None

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/get_onedrive_info', methods=['POST'])
def get_onedrive_info():
    try:
        data = request.json
        upn = data.get('upn', '').strip()
        
        if not upn:
            return jsonify({"error": "User UPN is required"}), 400
        
        # Get Graph API token
        graph_token = get_token_with_certificate(CONFIG['scopes']['graph'])
        if not graph_token:
            graph_token = get_token_with_secret(CONFIG['scopes']['graph'])
        
        if not graph_token:
            return jsonify({"error": "Failed to get Graph API token"}), 500
        
        # Get user OneDrive URL
        user_url = f"https://graph.microsoft.com/v1.0/users/{upn}/drive"
        headers = {
            "Authorization": f"Bearer {graph_token}",
            "Accept": "application/json"
        }
        
        response = requests.get(user_url, headers=headers)
        if response.status_code != 200:
            return jsonify({"error": f"Failed to get user OneDrive: {response.text}"}), 400
        
        drive_data = response.json()
        onedrive_url = drive_data.get('webUrl', '')
        
        if not onedrive_url:
            return jsonify({"error": "OneDrive URL not found"}), 400
        
        # Extract site URL from OneDrive URL
        if '/personal/' in onedrive_url:
            site_url = onedrive_url.split('/Documents')[0] + '/'
        else:
            return jsonify({"error": "Invalid OneDrive URL format"}), 400
        
        # Get SharePoint token
        sp_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sp_token:
            sp_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
        if not sp_token:
            return jsonify({"error": "Failed to get SharePoint token"}), 500
        
        # Get site information
        site_info_url = f"{site_url}_api/web"
        headers = {
            "Authorization": f"Bearer {sp_token}",
            "Accept": "application/json;odata=verbose"
        }
        
        response = requests.get(site_info_url, headers=headers)
        if response.status_code == 200:
            site_data = response.json()
            site_title = site_data.get('d', {}).get('Title', 'N/A')
            
            # Get site owner information
            owner_url = f"{site_url}_api/web/associatedownergroup"
            response = requests.get(owner_url, headers=headers)
            owner = 'N/A'
            if response.status_code == 200:
                owner_data = response.json()
                owner = owner_data.get('d', {}).get('Title', 'N/A')
        else:
            site_title = 'N/A'
            owner = 'N/A'
        
        return jsonify({
            "upn": upn,
            "onedrive_url": onedrive_url,
            "site_title": site_title,
            "owner": owner
        })
        
    except Exception as e:
        logger.exception("Error getting OneDrive info")
        return jsonify({"error": str(e)}), 500

@app.route('/repair_onedrive', methods=['POST'])
def repair_onedrive():
    try:
        data = request.json
        upn = data.get('upn', '').strip()
        onedrive_url = data.get('onedrive_url', '').strip()
        
        if not upn or not onedrive_url:
            return jsonify({"error": "User UPN and OneDrive URL are required"}), 400
        
        # Extract site URL from OneDrive URL
        if '/personal/' in onedrive_url:
            site_url = onedrive_url.split('/Documents')[0] + '/'
        else:
            return jsonify({"error": "Invalid OneDrive URL format", "failed_step": 0}), 400
        
        # Get SharePoint token
        sp_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sp_token:
            sp_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
        if not sp_token:
            return jsonify({"error": "Failed to get SharePoint token", "failed_step": 0}), 500
        
        # Get request digest
        request_digest = get_request_digest(site_url, sp_token)
        if not request_digest:
            return jsonify({"error": "Failed to get request digest", "failed_step": 0}), 500
        
        logger.info(f"Starting OneDrive repair for {upn}")
        
        # Step 1: Add repair account as user
        logger.info("Step 1: Adding repair account as user")
        repair_user_id = ensure_user(site_url, sp_token, request_digest, CONFIG['repair_account'])
        if not repair_user_id:
            return jsonify({"error": "Failed to add repair account as user", "failed_step": 1}), 500
        
        # Step 2: Make repair account site admin
        logger.info("Step 2: Making repair account site admin")
        if not set_site_admin(site_url, sp_token, request_digest, repair_user_id, True):
            return jsonify({"error": "Failed to make repair account site admin", "failed_step": 2}), 500
        
        # Step 3: Get original user ID and remove admin rights
        logger.info("Step 3: Getting original user and removing admin rights")
        original_user_id = get_user_id_from_site(site_url, sp_token, upn)
        if original_user_id:
            if not set_site_admin(site_url, sp_token, request_digest, original_user_id, False):
                logger.warning("Failed to remove original user admin rights, continuing...")
        
        # Step 4: Remove original user from site
        logger.info("Step 4: Removing original user from site")
        if original_user_id:
            if not remove_user_by_id(site_url, sp_token, request_digest, original_user_id):
                logger.warning("Failed to remove original user, continuing...")
        
        # Wait a moment for changes to propagate
        time.sleep(2)
        
        # Step 5: Re-add original user to site
        logger.info("Step 5: Re-adding original user to site")
        new_user_id = ensure_user(site_url, sp_token, request_digest, upn)
        if not new_user_id:
            return jsonify({"error": "Failed to re-add original user to site", "failed_step": 5}), 500
        
        # Step 6: Make original user site admin
        logger.info("Step 6: Making original user site admin")
        if not set_site_admin(site_url, sp_token, request_digest, new_user_id, True):
            return jsonify({"error": "Failed to make original user site admin", "failed_step": 6}), 500
        
        # Step 7: Remove repair account admin rights
        logger.info("Step 7: Removing repair account admin rights")
        if not set_site_admin(site_url, sp_token, request_digest, repair_user_id, False):
            logger.warning("Failed to remove repair account admin rights, continuing...")
        
        # Step 8: Remove repair account from site
        logger.info("Step 8: Removing repair account from site")
        if not remove_user_by_id(site_url, sp_token, request_digest, repair_user_id):
            logger.warning("Failed to remove repair account from site")
        
        logger.info(f"OneDrive repair completed successfully for {upn}")
        return jsonify({"success": True, "message": "OneDrive repair completed successfully"})
        
    except Exception as e:
        logger.exception("Error during OneDrive repair")
        return jsonify({"error": str(e), "failed_step": 0}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5050)
		
		