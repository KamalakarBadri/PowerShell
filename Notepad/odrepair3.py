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
    <title>SharePoint User Remover</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body>
    <h1>SharePoint Site User Remover</h1>
    <p><strong>Warning:</strong> This will permanently remove the user from the specified site.</p>
    
    <form id="userForm">
        <h3>User Removal Settings</h3>
        
        <label for="removal-type">Removal Type:</label>
        <select id="removal-type" onchange="toggleRemovalType()">
            <option value="site">Remove from SharePoint Site</option>
            <option value="onedrive">Remove from User's OneDrive</option>
        </select>
        <br><br>
        
        <label for="user-upn">User to Remove (Email/UPN):</label>
        <input type="email" id="user-upn" placeholder="user@geekbyteonline.onmicrosoft.com" required>
        <br><br>
        
        <div id="site-url-group">
            <label for="site-url">SharePoint Site URL:</label>
            <input type="url" id="site-url" placeholder="https://geekbyteonline.sharepoint.com/sites/sitename">
            <br><br>
        </div>
        
        <div id="onedrive-owner-group" style="display: none;">
            <label for="onedrive-owner-upn">OneDrive Owner (Email/UPN):</label>
            <input type="email" id="onedrive-owner-upn" placeholder="owner@geekbyteonline.onmicrosoft.com">
            <br><br>
        </div>
        
        <button type="button" id="find-user-btn" onclick="findUser()">Find User</button>
        <button type="button" id="remove-user-btn" onclick="removeUser()" disabled>Remove User</button>
    </form>
    
    <div id="status">
        <p>Enter user details to start</p>
    </div>
    
    <div id="results-section" style="display: none;">
        <h3>User Information</h3>
        <div id="user-info"></div>
        <div id="confirmation-section" style="display: none;"></div>
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
            resetForm();
        }
        
        function resetForm() {
            currentUserId = null;
            currentSiteUrl = null;
            document.getElementById('remove-user-btn').disabled = true;
            document.getElementById('results-section').style.display = 'none';
            setStatus('Enter user details to start');
        }
        
        function setStatus(message) {
            document.getElementById('status').innerHTML = '<p>' + message + '</p>';
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
                    setStatus('Error: Please enter both user UPN and site URL');
                    return;
                }
            } else {
                onedriveOwnerUpn = document.getElementById('onedrive-owner-upn').value.trim();
                if (!upn || !onedriveOwnerUpn) {
                    setStatus('Error: Please enter both user UPN and OneDrive owner UPN');
                    return;
                }
            }
            
            // Disable buttons and show loading
            findBtn.disabled = true;
            removeBtn.disabled = true;
            findBtn.textContent = 'Finding User...';
            setStatus(removalType === 'site' ? 'Searching for user on the specified site...' : 'Getting OneDrive URL and searching for user...');
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
                findBtn.textContent = 'Find User';
                
                if (data.error) {
                    setStatus('Error: ' + data.error);
                    resultsSection.style.display = 'none';
                    removeBtn.disabled = true;
                } else {
                    setStatus(removalType === 'site' ? 'User found on site' : 'User found on OneDrive');
                    displayUserInfo(data);
                    resultsSection.style.display = 'block';
                    removeBtn.disabled = false;
                    currentUserId = data.user_id;
                    currentSiteUrl = data.site_url;
                }
            })
            .catch(err => {
                findBtn.disabled = false;
                findBtn.textContent = 'Find User';
                setStatus('Request failed: ' + err.message);
                resultsSection.style.display = 'none';
                removeBtn.disabled = true;
            });
        }
        
        function removeUser() {
            if (!currentUserId || !currentSiteUrl) {
                setStatus('Error: Please find a user first');
                return;
            }
            
            if (!confirm('Are you sure you want to remove this user? This action cannot be undone.')) {
                return;
            }
            
            const removeBtn = document.getElementById('remove-user-btn');
            const findBtn = document.getElementById('find-user-btn');
            
            // Disable buttons and show loading
            removeBtn.disabled = true;
            findBtn.disabled = true;
            removeBtn.textContent = 'Removing User...';
            setStatus('Removing user from site...');
            
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
                removeBtn.textContent = 'Remove User';
                
                if (data.error) {
                    setStatus('Error: ' + data.error);
                } else {
                    setStatus('User successfully removed from site');
                    displayRemovalConfirmation(data);
                    // Reset form
                    currentUserId = null;
                    removeBtn.disabled = true;
                }
            })
            .catch(err => {
                removeBtn.disabled = false;
                findBtn.disabled = false;
                removeBtn.textContent = 'Remove User';
                setStatus('Request failed: ' + err.message);
            });
        }
        
        function displayUserInfo(data) {
            const userInfoDiv = document.getElementById('user-info');
            
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
        
        function displayRemovalConfirmation(data) {
            const confirmationDiv = document.getElementById('confirmation-section');
            
            confirmationDiv.innerHTML = `
                <h4>User Removed Successfully</h4>
                <p>User ID <strong>${data.user_id}</strong> has been successfully removed from the site.</p>
                <p><strong>Removal Time:</strong> ${new Date().toLocaleString()}</p>
            `;
            
            confirmationDiv.style.display = 'block';
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
            
            logger.info(f"Getting OneDrive for owner: {onedrive_owner_upn}")
            onedrive_response = requests.get(
                f"https://graph.microsoft.com/v1.0/users/{onedrive_owner_upn}/drive?$select=webUrl",
                headers=graph_headers
            )
            
            if onedrive_response.status_code != 200:
                logger.error(f"OneDrive lookup failed: {onedrive_response.text}")
                return jsonify({"error": f"OneDrive not found for owner: {onedrive_response.text}"}), 404
            
            onedrive_info = onedrive_response.json()
            onedrive_url = onedrive_info.get('webUrl', '')
            
            if not onedrive_url:
                return jsonify({"error": "OneDrive URL not found for owner"}), 404
            
            # Remove /Documents from the end if present
            if onedrive_url.endswith('/Documents'):
                site_url = onedrive_url[:-10]
            else:
                site_url = onedrive_url
            
            site_type = 'OneDrive'
            onedrive_owner = onedrive_owner_upn
            logger.info(f"OneDrive URL: {site_url}")
        
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
            return jsonify({"error": "Failed to obtain SharePoint access token"}), 500
        
        # Call SharePoint API to get site users
        sharepoint_headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/xml"
        }
        
        logger.info("Calling SharePoint API to get site users")
        site_users_response = requests.get(site_users_url, headers=sharepoint_headers)
        
        if site_users_response.status_code != 200:
            logger.error(f"SharePoint API failed: {site_users_response.status_code} - {site_users_response.text}")
            return jsonify({"error": f"Failed to get site users: {site_users_response.text}"}), 500
        
        # Parse XML and find the specific user
        user_info = parse_site_users_xml(site_users_response.text, upn)
        
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
            return jsonify({"error": "Failed to obtain SharePoint access token"}), 500
        
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
            return jsonify({"error": f"Failed to remove user: {remove_response.text}"}), 500
        
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

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5002, debug=True)