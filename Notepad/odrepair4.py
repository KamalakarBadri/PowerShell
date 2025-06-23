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
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed

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
    <title>SharePoint User Remover</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .tab-container { margin-bottom: 20px; }
        .tab-button { padding: 10px 20px; margin-right: 5px; border: 1px solid #ccc; background: #f9f9f9; cursor: pointer; }
        .tab-button.active { background: #007cba; color: white; }
        .tab-content { display: none; border: 1px solid #ccc; padding: 20px; }
        .tab-content.active { display: block; }
        .form-group { margin-bottom: 15px; }
        .form-group label { display: block; margin-bottom: 5px; font-weight: bold; }
        .form-group input, .form-group select, .form-group textarea { width: 100%; padding: 8px; border: 1px solid #ccc; border-radius: 4px; }
        .form-group textarea { height: 100px; resize: vertical; }
        .button { padding: 10px 20px; margin: 5px; border: none; border-radius: 4px; cursor: pointer; }
        .button-primary { background: #007cba; color: white; }
        .button-secondary { background: #6c757d; color: white; }
        .button:disabled { background: #ccc; cursor: not-allowed; }
        .results-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        .results-table th, .results-table td { padding: 10px; border: 1px solid #ccc; text-align: left; }
        .results-table th { background: #f9f9f9; }
        .status-success { color: green; font-weight: bold; }
        .status-error { color: red; font-weight: bold; }
        .status-pending { color: orange; font-weight: bold; }
        .progress-bar { width: 100%; height: 20px; background: #f0f0f0; border-radius: 10px; margin: 10px 0; }
        .progress-fill { height: 100%; background: #007cba; border-radius: 10px; transition: width 0.3s ease; }
        .warning { background: #fff3cd; border: 1px solid #ffeaa7; padding: 15px; border-radius: 4px; margin-bottom: 20px; }
    </style>
</head>
<body>
    <h1>SharePoint User Remover</h1>
    
    <div class="tab-container">
        <button class="tab-button active" onclick="showTab('single')">Single Removal</button>
        <button class="tab-button" onclick="showTab('bulk')">Bulk OneDrive Removal</button>
    </div>
    
    <!-- Single Removal Tab -->
    <div id="single-tab" class="tab-content active">
        <div class="warning">
            <strong>Warning:</strong> This will permanently remove the user from the specified site.
        </div>
        
        <form id="singleUserForm">
            <div class="form-group">
                <label for="single-removal-type">Removal Type:</label>
                <select id="single-removal-type" onchange="toggleSingleRemovalType()">
                    <option value="site">Remove from SharePoint Site</option>
                    <option value="onedrive">Remove from User's OneDrive</option>
                </select>
            </div>
            
            <div class="form-group">
                <label for="single-user-upn">User to Remove (Email/UPN):</label>
                <input type="email" id="single-user-upn" placeholder="user@geekbyteonline.onmicrosoft.com" required>
            </div>
            
            <div id="single-site-url-group" class="form-group">
                <label for="single-site-url">SharePoint Site URL:</label>
                <input type="url" id="single-site-url" placeholder="https://geekbyteonline.sharepoint.com/sites/sitename">
            </div>
            
            <div id="single-onedrive-owner-group" class="form-group" style="display: none;">
                <label for="single-onedrive-owner-upn">OneDrive Owner (Email/UPN):</label>
                <input type="email" id="single-onedrive-owner-upn" placeholder="owner@geekbyteonline.onmicrosoft.com">
            </div>
            
            <button type="button" id="single-find-user-btn" class="button button-primary" onclick="findSingleUser()">Find User</button>
            <button type="button" id="single-remove-user-btn" class="button button-secondary" onclick="removeSingleUser()" disabled>Remove User</button>
        </form>
        
        <div id="single-status">
            <p>Enter user details to start</p>
        </div>
        
        <div id="single-results-section" style="display: none;">
            <h3>User Information</h3>
            <div id="single-user-info"></div>
            <div id="single-confirmation-section" style="display: none;"></div>
        </div>
    </div>
    
    <!-- Bulk Removal Tab -->
    <div id="bulk-tab" class="tab-content">
        <div class="warning">
            <strong>Warning:</strong> This will remove the specified user from ALL listed OneDrive sites. This action cannot be undone!
        </div>
        
        <form id="bulkUserForm">
            <div class="form-group">
                <label for="bulk-user-upn">User to Remove from Multiple OneDrives (Email/UPN):</label>
                <input type="email" id="bulk-user-upn" placeholder="user@geekbyteonline.onmicrosoft.com" required>
            </div>
            
            <div class="form-group">
                <label for="onedrive-owners">OneDrive Owners (one email per line):</label>
                <textarea id="onedrive-owners" placeholder="owner1@geekbyteonline.onmicrosoft.com&#10;owner2@geekbyteonline.onmicrosoft.com&#10;owner3@geekbyteonline.onmicrosoft.com" required></textarea>
                <small>Enter one email address per line for each OneDrive owner</small>
            </div>
            
            <button type="button" id="bulk-find-btn" class="button button-primary" onclick="findBulkUsers()">Find User on All OneDrives</button>
            <button type="button" id="bulk-remove-btn" class="button button-secondary" onclick="removeBulkUsers()" disabled>Remove from All Found OneDrives</button>
        </form>
        
        <div id="bulk-status">
            <p>Enter user details to start bulk removal</p>
        </div>
        
        <div id="bulk-progress" style="display: none;">
            <h3>Progress</h3>
            <div class="progress-bar">
                <div id="progress-fill" class="progress-fill" style="width: 0%;"></div>
            </div>
            <p id="progress-text">0 / 0 completed</p>
        </div>
        
        <div id="bulk-results-section" style="display: none;">
            <h3>Results</h3>
            <table id="bulk-results-table" class="results-table">
                <thead>
                    <tr>
                        <th>OneDrive Owner</th>
                        <th>User Found</th>
                        <th>Removal Status</th>
                        <th>Details</th>
                    </tr>
                </thead>
                <tbody id="bulk-results-body">
                </tbody>
            </table>
        </div>
    </div>

    <script>
        let singleCurrentUserId = null;
        let singleCurrentSiteUrl = null;
        let bulkResults = [];
        
        function showTab(tabName) {
            // Hide all tabs
            document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
            document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
            
            // Show selected tab
            document.getElementById(tabName + '-tab').classList.add('active');
            event.target.classList.add('active');
            
            // Reset forms when switching tabs
            if (tabName === 'single') {
                resetSingleForm();
            } else {
                resetBulkForm();
            }
        }
        
        function toggleSingleRemovalType() {
            const removalType = document.getElementById('single-removal-type').value;
            const siteUrlGroup = document.getElementById('single-site-url-group');
            const onedriveOwnerGroup = document.getElementById('single-onedrive-owner-group');
            
            if (removalType === 'site') {
                siteUrlGroup.style.display = 'block';
                onedriveOwnerGroup.style.display = 'none';
            } else {
                siteUrlGroup.style.display = 'none';
                onedriveOwnerGroup.style.display = 'block';
            }
            
            resetSingleForm();
        }
        
        function resetSingleForm() {
            singleCurrentUserId = null;
            singleCurrentSiteUrl = null;
            document.getElementById('single-remove-user-btn').disabled = true;
            document.getElementById('single-results-section').style.display = 'none';
            setSingleStatus('Enter user details to start');
        }
        
        function resetBulkForm() {
            bulkResults = [];
            document.getElementById('bulk-remove-btn').disabled = true;
            document.getElementById('bulk-results-section').style.display = 'none';
            document.getElementById('bulk-progress').style.display = 'none';
            setBulkStatus('Enter user details to start bulk removal');
        }
        
        function setSingleStatus(message) {
            document.getElementById('single-status').innerHTML = '<p>' + message + '</p>';
        }
        
        function setBulkStatus(message) {
            document.getElementById('bulk-status').innerHTML = '<p>' + message + '</p>';
        }
        
        function updateProgress(completed, total) {
            const progressFill = document.getElementById('progress-fill');
            const progressText = document.getElementById('progress-text');
            const percentage = total > 0 ? (completed / total) * 100 : 0;
            
            progressFill.style.width = percentage + '%';
            progressText.textContent = `${completed} / ${total} completed`;
        }
        
        // Single user functions (existing functionality)
        function findSingleUser() {
            const upn = document.getElementById('single-user-upn').value.trim();
            const removalType = document.getElementById('single-removal-type').value;
            const findBtn = document.getElementById('single-find-user-btn');
            const removeBtn = document.getElementById('single-remove-user-btn');
            const resultsSection = document.getElementById('single-results-section');
            
            let siteUrl = '';
            let onedriveOwnerUpn = '';
            
            if (removalType === 'site') {
                siteUrl = document.getElementById('single-site-url').value.trim();
                if (!upn || !siteUrl) {
                    setSingleStatus('Error: Please enter both user UPN and site URL');
                    return;
                }
            } else {
                onedriveOwnerUpn = document.getElementById('single-onedrive-owner-upn').value.trim();
                if (!upn || !onedriveOwnerUpn) {
                    setSingleStatus('Error: Please enter both user UPN and OneDrive owner UPN');
                    return;
                }
            }
            
            findBtn.disabled = true;
            removeBtn.disabled = true;
            findBtn.textContent = 'Finding User...';
            setSingleStatus(removalType === 'site' ? 'Searching for user on the specified site...' : 'Getting OneDrive URL and searching for user...');
            resultsSection.style.display = 'none';
            singleCurrentUserId = null;
            singleCurrentSiteUrl = null;
            
            const requestData = {
                upn: upn,
                removal_type: removalType
            };
            
            if (removalType === 'site') {
                requestData.site_url = siteUrl;
                singleCurrentSiteUrl = siteUrl;
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
                findBtn.textContent = 'Find User';
                
                if (data.error) {
                    setSingleStatus('Error: ' + data.error);
                    resultsSection.style.display = 'none';
                    removeBtn.disabled = true;
                } else {
                    setSingleStatus(removalType === 'site' ? 'User found on site' : 'User found on OneDrive');
                    displaySingleUserInfo(data);
                    resultsSection.style.display = 'block';
                    removeBtn.disabled = false;
                    singleCurrentUserId = data.user_id;
                    singleCurrentSiteUrl = data.site_url;
                }
            })
            .catch(err => {
                findBtn.disabled = false;
                findBtn.textContent = 'Find User';
                setSingleStatus('Request failed: ' + err.message);
                resultsSection.style.display = 'none';
                removeBtn.disabled = true;
            });
        }
        
        function removeSingleUser() {
            if (!singleCurrentUserId || !singleCurrentSiteUrl) {
                setSingleStatus('Error: Please find a user first');
                return;
            }
            
            if (!confirm('Are you sure you want to remove this user? This action cannot be undone.')) {
                return;
            }
            
            const removeBtn = document.getElementById('single-remove-user-btn');
            const findBtn = document.getElementById('single-find-user-btn');
            
            removeBtn.disabled = true;
            findBtn.disabled = true;
            removeBtn.textContent = 'Removing User...';
            setSingleStatus('Removing user from site...');
            
            fetch('/remove_user', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    user_id: singleCurrentUserId, 
                    site_url: singleCurrentSiteUrl 
                })
            })
            .then(res => res.json())
            .then(data => {
                removeBtn.disabled = false;
                findBtn.disabled = false;
                removeBtn.textContent = 'Remove User';
                
                if (data.error) {
                    setSingleStatus('Error: ' + data.error);
                } else {
                    setSingleStatus('User successfully removed from site');
                    displaySingleRemovalConfirmation(data);
                    singleCurrentUserId = null;
                    removeBtn.disabled = true;
                }
            })
            .catch(err => {
                removeBtn.disabled = false;
                findBtn.disabled = false;
                removeBtn.textContent = 'Remove User';
                setSingleStatus('Request failed: ' + err.message);
            });
        }
        
        function displaySingleUserInfo(data) {
            const userInfoDiv = document.getElementById('single-user-info');
            
            userInfoDiv.innerHTML = `
                <h4>User Found</h4>
                <p><strong>Display Name:</strong> ${data.title || 'N/A'}</p>
                <p><strong>Email:</strong> ${data.email || 'N/A'}</p>
                <p><strong>User ID:</strong> ${data.user_id || 'N/A'}</p>
                <p><strong>Login Name:</strong> ${data.login_name || 'N/A'}</p>
                <p><strong>Site Type:</strong> ${data.site_type || 'SharePoint Site'}</p>
                <p><strong>Site URL:</strong> ${data.site_url || 'N/A'}</p>
                <p><strong>OneDrive Owner:</strong> ${data.onedrive_owner || 'N/A'}</p>
                <p><strong>Is Site Admin:</strong> ${data.is_site_admin ? 'Yes' : 'No'}</p>
            `;
        }
        
        function displaySingleRemovalConfirmation(data) {
            const confirmationDiv = document.getElementById('single-confirmation-section');
            
            confirmationDiv.innerHTML = `
                <h4>User Removed Successfully</h4>
                <p>User ID <strong>${data.user_id}</strong> has been successfully removed from the site.</p>
                <p><strong>Removal Time:</strong> ${new Date().toLocaleString()}</p>
            `;
            
            confirmationDiv.style.display = 'block';
        }
        
        // Bulk removal functions
        function findBulkUsers() {
            const upn = document.getElementById('bulk-user-upn').value.trim();
            const ownersText = document.getElementById('onedrive-owners').value.trim();
            
            if (!upn || !ownersText) {
                setBulkStatus('Error: Please enter both user UPN and OneDrive owners');
                return;
            }
            
            const owners = ownersText.split('\\n').map(owner => owner.trim()).filter(owner => owner.length > 0);
            
            if (owners.length === 0) {
                setBulkStatus('Error: Please enter at least one OneDrive owner');
                return;
            }
            
            const findBtn = document.getElementById('bulk-find-btn');
            const removeBtn = document.getElementById('bulk-remove-btn');
            
            findBtn.disabled = true;
            removeBtn.disabled = true;
            findBtn.textContent = 'Finding Users...';
            setBulkStatus(`Searching for user on ${owners.length} OneDrive sites...`);
            
            document.getElementById('bulk-progress').style.display = 'block';
            updateProgress(0, owners.length);
            
            bulkResults = [];
            
            fetch('/bulk_find_user', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    upn: upn,
                    onedrive_owners: owners
                })
            })
            .then(res => res.json())
            .then(data => {
                findBtn.disabled = false;
                findBtn.textContent = 'Find User on All OneDrives';
                
                if (data.error) {
                    setBulkStatus('Error: ' + data.error);
                    document.getElementById('bulk-progress').style.display = 'none';
                } else {
                    bulkResults = data.results;
                    const foundCount = bulkResults.filter(r => r.found).length;
                    setBulkStatus(`Search completed. User found on ${foundCount} out of ${owners.length} OneDrives.`);
                    displayBulkResults(bulkResults);
                    document.getElementById('bulk-results-section').style.display = 'block';
                    removeBtn.disabled = foundCount === 0;
                    updateProgress(owners.length, owners.length);
                }
            })
            .catch(err => {
                findBtn.disabled = false;
                findBtn.textContent = 'Find User on All OneDrives';
                setBulkStatus('Request failed: ' + err.message);
                document.getElementById('bulk-progress').style.display = 'none';
            });
        }
        
        function removeBulkUsers() {
            const foundUsers = bulkResults.filter(r => r.found);
            
            if (foundUsers.length === 0) {
                setBulkStatus('Error: No users found to remove');
                return;
            }
            
            if (!confirm(`Are you sure you want to remove the user from ${foundUsers.length} OneDrive sites? This action cannot be undone.`)) {
                return;
            }
            
            const findBtn = document.getElementById('bulk-find-btn');
            const removeBtn = document.getElementById('bulk-remove-btn');
            
            findBtn.disabled = true;
            removeBtn.disabled = true;
            removeBtn.textContent = 'Removing Users...';
            setBulkStatus(`Removing user from ${foundUsers.length} OneDrive sites...`);
            
            updateProgress(0, foundUsers.length);
            
            fetch('/bulk_remove_user', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    removal_requests: foundUsers.map(user => ({
                        user_id: user.user_id,
                        site_url: user.site_url,
                        onedrive_owner: user.onedrive_owner
                    }))
                })
            })
            .then(res => res.json())
            .then(data => {
                findBtn.disabled = false;
                removeBtn.disabled = false;
                removeBtn.textContent = 'Remove from All Found OneDrives';
                
                if (data.error) {
                    setBulkStatus('Error: ' + data.error);
                } else {
                    bulkResults = data.results;
                    const successCount = bulkResults.filter(r => r.removal_success).length;
                    setBulkStatus(`Bulk removal completed. Successfully removed from ${successCount} out of ${foundUsers.length} OneDrives.`);
                    displayBulkResults(bulkResults);
                    updateProgress(foundUsers.length, foundUsers.length);
                }
            })
            .catch(err => {
                findBtn.disabled = false;
                removeBtn.disabled = false;
                removeBtn.textContent = 'Remove from All Found OneDrives';
                setBulkStatus('Request failed: ' + err.message);
            });
        }
        
        function displayBulkResults(results) {
            const tbody = document.getElementById('bulk-results-body');
            tbody.innerHTML = '';
            
            results.forEach(result => {
                const row = tbody.insertRow();
                
                // OneDrive Owner
                row.insertCell(0).textContent = result.onedrive_owner;
                
                // User Found
                const foundCell = row.insertCell(1);
                if (result.found) {
                    foundCell.innerHTML = '<span class="status-success">✓ Found</span>';
                } else {
                    foundCell.innerHTML = '<span class="status-error">✗ Not Found</span>';
                }
                
                // Removal Status
                const statusCell = row.insertCell(2);
                if (!result.found) {
                    statusCell.innerHTML = '<span class="status-pending">N/A</span>';
                } else if (result.removal_success === undefined) {
                    statusCell.innerHTML = '<span class="status-pending">Pending</span>';
                } else if (result.removal_success) {
                    statusCell.innerHTML = '<span class="status-success">✓ Removed</span>';
                } else {
                    statusCell.innerHTML = '<span class="status-error">✗ Failed</span>';
                }
                
                // Details
                const detailsCell = row.insertCell(3);
                if (result.error) {
                    detailsCell.textContent = result.error;
                } else if (result.found) {
                    detailsCell.textContent = `User: ${result.title || 'N/A'} (${result.email || 'N/A'})`;
                } else {
                    detailsCell.textContent = 'User not found on this OneDrive';
                }
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
        
        # Call SharePoint API to get site users
        sharepoint_headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/xml"
        }
        
        logger.info("Calling SharePoint API to get site users")
        site_users_response = requests.get(site_users_url, headers=sharepoint_headers)
        
        if site_users_response.status_code != 200:
            logger.error(f"SharePoint API failed: {site_users_response.status_code} - {site_users_response.text}")
            raise Exception(f"Failed to get site users: {site_users_response.text}")
        
        # Parse XML and find the specific user
        user_info = parse_site_users_xml(site_users_response.text, target_upn)
        return user_info
        
    except Exception as e:
        logger.exception(f"Failed to find user {target_upn} on site {site_url}")
        raise

def remove_user_from_site(user_id, site_url):
    """Remove a user from a SharePoint site/OneDrive"""
    try:
        # Ensure site URL ends properly
        if not site_url.endswith('/'):
            site_url += '/'
        
        # Construct SharePoint API URL for user removal
        remove_user_url = f"{site_url}_api/web/siteusers/removebyid({user_id})"
        logger.info(f"Remove user URL: {remove_user_url}")
        
        # Get SharePoint token
        sharepoint_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sharepoint_token:
            sharepoint_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
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

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/find_user', methods=['POST'])
def find_user():
    try:
        data = request.json
        upn = data.get('upn', '').strip()
        removal_type = data.get('removal_type', 'site')
        
        logger.info(f"Finding user: {upn}, type: {removal_type}")
        
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
            
            site_url = get_onedrive_url(onedrive_owner_upn)
            site_type = 'OneDrive'
            onedrive_owner = onedrive_owner_upn
        
        # Find user on site
        user_info = find_user_on_site(upn, site_url)
        
        if not user_info:
            return jsonify({"error": f"User '{upn}' not found on the specified {site_type.lower()}"}), 404
        
        # Add additional info to response
        user_info['site_url'] = site_url
        user_info['site_type'] = site_type
        user_info['onedrive_owner'] = onedrive_owner
        
        logger.info(f"Successfully found user: {user_info}")
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
        
        logger.info(f"Removing user ID: {user_id} from site: {site_url}")
        
        if not user_id or not site_url:
            return jsonify({"error": "Both user ID and site URL are required"}), 400
        
        # Remove user from site
        remove_user_from_site(user_id, site_url)
        
        logger.info(f"Successfully removed user {user_id} from {site_url}")
        return jsonify({
            "success": True,
            "message": "User successfully removed from site",
            "user_id": user_id,
            "site_url": site_url
        })
        
    except Exception as e:
        logger.exception("Error occurred during user removal")
        return jsonify({"error": str(e)}), 500

@app.route('/bulk_find_user', methods=['POST'])
def bulk_find_user():
    try:
        data = request.json
        target_upn = data.get('upn', '').strip()
        onedrive_owners = data.get('onedrive_owners', [])
        
        logger.info(f"Bulk finding user: {target_upn} on {len(onedrive_owners)} OneDrives")
        
        if not target_upn or not onedrive_owners:
            return jsonify({"error": "Both user UPN and OneDrive owners are required"}), 400
        
        results = []
        
        def process_onedrive_owner(owner):
            result = {
                'onedrive_owner': owner,
                'found': False,
                'error': None,
                'site_url': None,
                'user_id': None,
                'title': None,
                'email': None,
                'login_name': None,
                'is_site_admin': False
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
                    result['title'] = user_info['title']
                    result['email'] = user_info['email']
                    result['login_name'] = user_info['login_name']
                    result['is_site_admin'] = user_info['is_site_admin']
                    logger.info(f"Found user on {owner}'s OneDrive")
                else:
                    logger.info(f"User not found on {owner}'s OneDrive")
                
            except Exception as e:
                logger.exception(f"Error processing OneDrive for {owner}")
                result['error'] = str(e)
            
            return result
        
        # Process OneDrives in parallel for better performance
        with ThreadPoolExecutor(max_workers=5) as executor:
            future_to_owner = {executor.submit(process_onedrive_owner, owner): owner for owner in onedrive_owners}
            
            for future in as_completed(future_to_owner):
                result = future.result()
                results.append(result)
        
        # Sort results by owner name for consistent display
        results.sort(key=lambda x: x['onedrive_owner'])
        
        found_count = len([r for r in results if r['found']])
        logger.info(f"Bulk search completed. Found user on {found_count} out of {len(onedrive_owners)} OneDrives")
        
        return jsonify({
            "success": True,
            "results": results,
            "summary": {
                "total_searched": len(onedrive_owners),
                "found_count": found_count,
                "not_found_count": len(onedrive_owners) - found_count
            }
        })
        
    except Exception as e:
        logger.exception("Error occurred during bulk user search")
        return jsonify({"error": str(e)}), 500

@app.route('/bulk_remove_user', methods=['POST'])
def bulk_remove_user():
    try:
        data = request.json
        removal_requests = data.get('removal_requests', [])
        
        logger.info(f"Bulk removing user from {len(removal_requests)} OneDrives")
        
        if not removal_requests:
            return jsonify({"error": "No removal requests provided"}), 400
        
        results = []
        
        def process_removal(request):
            result = {
                'onedrive_owner': request['onedrive_owner'],
                'site_url': request['site_url'],
                'user_id': request['user_id'],
                'found': True,  # These are pre-filtered to only found users
                'removal_success': False,
                'error': None
            }
            
            try:
                remove_user_from_site(request['user_id'], request['site_url'])
                result['removal_success'] = True
                logger.info(f"Successfully removed user from {request['onedrive_owner']}'s OneDrive")
                
            except Exception as e:
                logger.exception(f"Error removing user from {request['onedrive_owner']}'s OneDrive")
                result['error'] = str(e)
            
            return result
        
        # Process removals in parallel for better performance
        with ThreadPoolExecutor(max_workers=5) as executor:
            future_to_request = {executor.submit(process_removal, req): req for req in removal_requests}
            
            for future in as_completed(future_to_request):
                result = future.result()
                results.append(result)
        
        # Sort results by owner name for consistent display
        results.sort(key=lambda x: x['onedrive_owner'])
        
        success_count = len([r for r in results if r['removal_success']])
        logger.info(f"Bulk removal completed. Successfully removed from {success_count} out of {len(removal_requests)} OneDrives")
        
        return jsonify({
            "success": True,
            "results": results,
            "summary": {
                "total_attempts": len(removal_requests),
                "success_count": success_count,
                "failure_count": len(removal_requests) - success_count
            }
        })
        
    except Exception as e:
        logger.exception("Error occurred during bulk user removal")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5005, debug=True)