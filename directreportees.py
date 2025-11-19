from flask import Flask, render_template_string, request, jsonify, Response
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
import csv
import io

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
        "graph": "https://graph.microsoft.com/.default"
    }
}

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Nested Direct Reports Export</title>
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
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
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
        
        .checkbox-group {
            display: flex;
            align-items: center;
            gap: 10px;
            margin-bottom: 15px;
        }
        
        .checkbox-group input[type="checkbox"] {
            width: 18px;
            height: 18px;
        }
        
        .checkbox-group label {
            margin-bottom: 0;
            font-weight: 500;
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
        
        .btn-primary {
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
            color: white;
        }
        
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 16px rgba(40, 167, 69, 0.3);
        }
        
        .btn-primary:disabled {
            opacity: 0.6;
            cursor: not-allowed;
            transform: none;
            box-shadow: none;
        }
        
        .btn-block {
            width: 100%;
            justify-content: center;
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
            color: #28a745;
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
        
        .stats-section {
            background: #e3f2fd;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            text-align: center;
        }
        
        .stat-item {
            padding: 15px;
        }
        
        .stat-number {
            font-size: 2rem;
            font-weight: 700;
            color: #007bff;
            margin-bottom: 5px;
        }
        
        .stat-label {
            font-size: 0.9rem;
            color: #6c757d;
            font-weight: 500;
        }
        
        .reports-section {
            margin-top: 20px;
        }
        
        .reports-section h3 {
            color: #28a745;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
            gap: 10px;
        }
        
        .hierarchy-tree {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 20px;
            margin-bottom: 20px;
            max-height: 400px;
            overflow-y: auto;
        }
        
        .tree-item {
            padding: 8px 0;
            border-left: 2px solid #dee2e6;
            margin-left: 10px;
            padding-left: 15px;
        }
        
        .tree-item:last-child {
            border-left: 2px solid transparent;
        }
        
        .tree-level-0 { margin-left: 0; border-left: none; padding-left: 0; }
        .tree-level-1 { margin-left: 20px; }
        .tree-level-2 { margin-left: 40px; }
        .tree-level-3 { margin-left: 60px; }
        .tree-level-4 { margin-left: 80px; }
        .tree-level-5 { margin-left: 100px; }
        
        .tree-user {
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 8px 12px;
            background: white;
            border-radius: 6px;
            border: 1px solid #e9ecef;
            margin-bottom: 5px;
        }
        
        .tree-user .name {
            font-weight: 600;
            color: #495057;
        }
        
        .tree-user .upn {
            color: #6c757d;
            font-size: 0.9rem;
        }
        
        .tree-user .count {
            background: #007bff;
            color: white;
            padding: 2px 8px;
            border-radius: 12px;
            font-size: 0.8rem;
            margin-left: auto;
        }
        
        .no-reports {
            text-align: center;
            color: #6c757d;
            padding: 40px;
            font-size: 1.1rem;
        }
        
        .export-section {
            text-align: center;
            margin-top: 20px;
        }
        
        .btn-export {
            background: linear-gradient(135deg, #007bff 0%, #0056b3 100%);
            color: white;
        }
        
        .btn-export:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 16px rgba(0, 123, 255, 0.3);
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
            
            .stats-grid {
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
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i class="fas fa-sitemap"></i> Nested Direct Reports Export</h1>
            <p>Get complete organizational hierarchy with nested direct reports</p>
        </div>
        
        <div class="main-content">
            <div class="search-section">
                <h3 style="margin-bottom: 20px; color: #28a745;">
                    <i class="fas fa-search"></i> User Search
                </h3>
                
                <div class="form-group">
                    <label for="user-upn">User Principal Name (UPN):</label>
                    <input type="email" id="user-upn" class="form-control" 
                           placeholder="user@geekbyteonline.onmicrosoft.com" 
                           value="">
                </div>
                
                <div class="checkbox-group">
                    <input type="checkbox" id="include-nested" checked>
                    <label for="include-nested">Include nested direct reports (full hierarchy)</label>
                </div>
                
                <button id="search-btn" class="btn btn-primary btn-block" onclick="searchUser()">
                    <i class="fas fa-search"></i> Get Organizational Hierarchy
                </button>
            </div>
            
            <div id="status" class="status-section status-ready">
                <i class="fas fa-info-circle"></i> Enter a user UPN to get organizational hierarchy
            </div>
            
            <div id="results-section" class="results-section" style="display: none;">
                <div id="user-info" class="user-info"></div>
                <div id="stats-section" class="stats-section"></div>
                <div id="reports-section" class="reports-section"></div>
                <div id="export-section" class="export-section"></div>
            </div>
        </div>
    </div>

    <script>
        function setStatus(type, message) {
            const statusDiv = document.getElementById('status');
            statusDiv.className = `status-section status-${type}`;
            
            let icon = 'fas fa-info-circle';
            if (type === 'loading') icon = 'fas fa-spinner fa-spin';
            else if (type === 'success') icon = 'fas fa-check-circle';
            else if (type === 'error') icon = 'fas fa-exclamation-circle';
            
            statusDiv.innerHTML = `<i class="${icon}"></i> ${message}`;
        }
        
        function searchUser() {
            const upn = document.getElementById('user-upn').value.trim();
            const includeNested = document.getElementById('include-nested').checked;
            const searchBtn = document.getElementById('search-btn');
            const resultsSection = document.getElementById('results-section');
            
            if (!upn) {
                setStatus('error', 'Please enter a valid UPN');
                return;
            }
            
            // Disable button and show loading
            searchBtn.disabled = true;
            searchBtn.innerHTML = '<div class="spinner"></div> Building hierarchy...';
            setStatus('loading', 'Building organizational hierarchy... This may take a while for large organizations.');
            resultsSection.style.display = 'none';
            
            fetch('/get_nested_direct_reports', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    upn: upn,
                    include_nested: includeNested 
                })
            })
            .then(res => res.json())
            .then(data => {
                searchBtn.disabled = false;
                searchBtn.innerHTML = '<i class="fas fa-search"></i> Get Organizational Hierarchy';
                
                if (data.error) {
                    setStatus('error', `Error: ${data.error}`);
                    resultsSection.style.display = 'none';
                } else {
                    setStatus('success', `Found ${data.total_users} users across ${data.max_depth} levels of hierarchy`);
                    displayResults(data);
                    resultsSection.style.display = 'block';
                }
            })
            .catch(err => {
                searchBtn.disabled = false;
                searchBtn.innerHTML = '<i class="fas fa-search"></i> Get Organizational Hierarchy';
                setStatus('error', 'Request failed: ' + err.message);
                resultsSection.style.display = 'none';
            });
        }
        
        function displayResults(data) {
            const userInfoDiv = document.getElementById('user-info');
            const statsSectionDiv = document.getElementById('stats-section');
            const reportsSectionDiv = document.getElementById('reports-section');
            const exportSectionDiv = document.getElementById('export-section');
            
            // Display user information
            userInfoDiv.innerHTML = `
                <h3><i class="fas fa-user"></i> Root Manager</h3>
                <div class="info-grid">
                    <div class="info-item">
                        <div class="info-label">Display Name</div>
                        <div class="info-value">${data.root_user.displayName || 'N/A'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">User Principal Name</div>
                        <div class="info-value">${data.root_user.userPrincipalName || 'N/A'}</div>
                    </div>
                    <div class="info-item">
                        <div class="info-label">User ID</div>
                        <div class="info-value">${data.root_user.id || 'N/A'}</div>
                    </div>
                </div>
            `;
            
            // Display statistics
            statsSectionDiv.innerHTML = `
                <h3><i class="fas fa-chart-bar"></i> Hierarchy Statistics</h3>
                <div class="stats-grid">
                    <div class="stat-item">
                        <div class="stat-number">${data.total_users}</div>
                        <div class="stat-label">Total Users</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-number">${data.max_depth}</div>
                        <div class="stat-label">Max Depth</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-number">${data.total_direct_reports}</div>
                        <div class="stat-label">Direct Reports</div>
                    </div>
                    <div class="stat-item">
                        <div class="stat-number">${data.total_nested_reports}</div>
                        <div class="stat-label">Nested Reports</div>
                    </div>
                </div>
            `;
            
            // Display hierarchy tree
            let reportsHtml = `
                <h3><i class="fas fa-sitemap"></i> Organizational Hierarchy</h3>
                <div class="hierarchy-tree">
                    ${buildTreeHTML(data.hierarchy, 0)}
                </div>
            `;
            
            reportsSectionDiv.innerHTML = reportsHtml;
            
            // Display export button
            if (data.total_users > 1) {
                exportSectionDiv.innerHTML = `
                    <a href="/export_nested_csv?upn=${encodeURIComponent(data.root_user.userPrincipalName)}&include_nested=true" class="btn btn-export">
                        <i class="fas fa-download"></i> Export Full Hierarchy to CSV
                    </a>
                `;
            } else {
                exportSectionDiv.innerHTML = '';
            }
        }
        
        function buildTreeHTML(hierarchy, level) {
            if (!hierarchy || hierarchy.length === 0) return '';
            
            let html = '';
            hierarchy.forEach(item => {
                const levelClass = `tree-level-${Math.min(level, 5)}`;
                const reportCount = item.direct_reports ? item.direct_reports.length : 0;
                
                html += `
                    <div class="tree-item ${levelClass}">
                        <div class="tree-user">
                            <i class="fas fa-user${level === 0 ? '-tie' : ''}"></i>
                            <div>
                                <div class="name">${item.user.displayName || 'Unknown'}</div>
                                <div class="upn">${item.user.userPrincipalName || 'No UPN'}</div>
                            </div>
                            ${reportCount > 0 ? `<span class="count">${reportCount}</span>` : ''}
                        </div>
                        ${item.direct_reports && item.direct_reports.length > 0 ? buildTreeHTML(item.direct_reports, level + 1) : ''}
                    </div>
                `;
            });
            return html;
        }
        
        // Allow Enter key to submit
        document.getElementById('user-upn').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                searchUser();
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

def get_direct_reports(graph_token, user_id):
    """Get direct reports for a user using Microsoft Graph API"""
    try:
        graph_headers = {
            "Authorization": f"Bearer {graph_token}",
            "Content-Type": "application/json"
        }
        
        direct_reports_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/directReports"
        
        response = requests.get(direct_reports_url, headers=graph_headers)
        
        if response.status_code == 200:
            direct_reports_data = response.json()
            # Extract user information from direct reports
            direct_reports = []
            for report in direct_reports_data.get('value', []):
                if report.get('@odata.type') == '#microsoft.graph.user':
                    direct_reports.append({
                        'id': report.get('id'),
                        'displayName': report.get('displayName'),
                        'userPrincipalName': report.get('userPrincipalName'),
                        'mail': report.get('mail'),
                        'jobTitle': report.get('jobTitle'),
                        'businessPhones': report.get('businessPhones', []),
                        'officeLocation': report.get('officeLocation'),
                        'department': report.get('department')
                    })
            return direct_reports
        else:
            logger.error(f"Failed to get direct reports for user {user_id}: {response.text}")
            return []
            
    except Exception as e:
        logger.exception(f"Failed to get direct reports for user {user_id}")
        return []

def get_nested_direct_reports_recursive(graph_token, user_info, current_depth=0, max_depth=10):
    """Recursively get nested direct reports with depth tracking"""
    if current_depth >= max_depth:
        return {
            'user': user_info,
            'direct_reports': [],
            'depth': current_depth
        }
    
    direct_reports = get_direct_reports(graph_token, user_info['id'])
    
    result = {
        'user': user_info,
        'direct_reports': [],
        'depth': current_depth
    }
    
    # Recursively get reports for each direct report
    for report in direct_reports:
        nested_result = get_nested_direct_reports_recursive(
            graph_token, 
            report, 
            current_depth + 1, 
            max_depth
        )
        result['direct_reports'].append(nested_result)
    
    return result

def flatten_hierarchy(hierarchy, flat_list=None, level=0):
    """Flatten the nested hierarchy into a flat list with level information"""
    if flat_list is None:
        flat_list = []
    
    # Add current user to flat list
    flat_list.append({
        **hierarchy['user'],
        'hierarchy_level': level,
        'manager_id': None,  # This will be filled in later for nested reports
        'manager_name': None,
        'manager_upn': None
    })
    
    # Process direct reports
    for report in hierarchy['direct_reports']:
        # Set manager information for the report
        report['user']['manager_id'] = hierarchy['user']['id']
        report['user']['manager_name'] = hierarchy['user']['displayName']
        report['user']['manager_upn'] = hierarchy['user']['userPrincipalName']
        
        # Recursively flatten
        flatten_hierarchy(report, flat_list, level + 1)
    
    return flat_list

def calculate_hierarchy_stats(hierarchy):
    """Calculate statistics about the hierarchy"""
    def count_nodes(node):
        total = 1  # Count current node
        direct_count = len(node['direct_reports'])
        nested_count = 0
        
        for child in node['direct_reports']:
            child_total, child_direct, child_nested = count_nodes(child)
            total += child_total
            nested_count += child_total  # All children count as nested for this node
        
        return total, direct_count, nested_count
    
    total_users, direct_reports, nested_reports = count_nodes(hierarchy)
    
    # Calculate max depth
    def max_depth(node):
        if not node['direct_reports']:
            return node['depth']
        return max(max_depth(child) for child in node['direct_reports'])
    
    max_depth_value = max_depth(hierarchy)
    
    return {
        'total_users': total_users,
        'total_direct_reports': direct_reports,
        'total_nested_reports': nested_reports,
        'max_depth': max_depth_value
    }

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/get_nested_direct_reports', methods=['POST'])
def get_nested_direct_reports_route():
    try:
        data = request.json
        upn = data.get('upn', '').strip()
        include_nested = data.get('include_nested', True)
        
        if not upn:
            return jsonify({"error": "UPN is required"}), 400
        
        # Get Graph API token
        graph_token = get_token_with_certificate(CONFIG['scopes']['graph'])
        if not graph_token:
            graph_token = get_token_with_secret(CONFIG['scopes']['graph'])
        
        if not graph_token:
            return jsonify({"error": "Failed to obtain Graph API access token"}), 500
        
        # Get root user info
        graph_headers = {
            "Authorization": f"Bearer {graph_token}",
            "Content-Type": "application/json"
        }
        
        user_response = requests.get(
            f"https://graph.microsoft.com/v1.0/users/{upn}?$select=id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation,businessPhones",
            headers=graph_headers
        )
        
        if user_response.status_code != 200:
            return jsonify({"error": f"User not found or access denied: {user_response.text}"}), 404
        
        root_user = user_response.json()
        
        if include_nested:
            # Get nested hierarchy
            hierarchy = get_nested_direct_reports_recursive(graph_token, root_user)
            stats = calculate_hierarchy_stats(hierarchy)
            
            return jsonify({
                "root_user": root_user,
                "hierarchy": [hierarchy],  # Wrap in list for consistency
                "total_users": stats['total_users'],
                "total_direct_reports": stats['total_direct_reports'],
                "total_nested_reports": stats['total_nested_reports'],
                "max_depth": stats['max_depth']
            })
        else:
            # Get only direct reports (non-recursive)
            direct_reports = get_direct_reports(graph_token, root_user['id'])
            
            return jsonify({
                "root_user": root_user,
                "hierarchy": [{
                    'user': root_user,
                    'direct_reports': [{'user': report, 'direct_reports': []} for report in direct_reports],
                    'depth': 0
                }],
                "total_users": len(direct_reports) + 1,
                "total_direct_reports": len(direct_reports),
                "total_nested_reports": 0,
                "max_depth": 1
            })
        
    except Exception as e:
        logger.exception("Error occurred while getting nested direct reports")
        return jsonify({"error": str(e)}), 500

@app.route('/export_nested_csv')
def export_nested_csv():
    try:
        upn = request.args.get('upn', '').strip()
        include_nested = request.args.get('include_nested', 'true').lower() == 'true'
        
        if not upn:
            return "UPN parameter is required", 400
        
        # Get Graph API token
        graph_token = get_token_with_certificate(CONFIG['scopes']['graph'])
        if not graph_token:
            graph_token = get_token_with_secret(CONFIG['scopes']['graph'])
        
        if not graph_token:
            return "Failed to obtain access token", 500
        
        # Get root user info
        graph_headers = {
            "Authorization": f"Bearer {graph_token}",
            "Content-Type": "application/json"
        }
        
        user_response = requests.get(
            f"https://graph.microsoft.com/v1.0/users/{upn}?$select=id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation,businessPhones",
            headers=graph_headers
        )
        
        if user_response.status_code != 200:
            return f"User not found: {user_response.text}", 404
        
        root_user = user_response.json()
        
        # Get nested hierarchy
        hierarchy = get_nested_direct_reports_recursive(graph_token, root_user)
        flat_list = flatten_hierarchy(hierarchy)
        
        # Create CSV
        output = io.StringIO()
        writer = csv.writer(output)
        
        # Write header
        writer.writerow([
            'Hierarchy Level',
            'Display Name', 
            'User Principal Name', 
            'Email', 
            'Job Title', 
            'Department',
            'Office Location',
            'Business Phones',
            'User ID',
            'Manager Name',
            'Manager UPN',
            'Manager ID'
        ])
        
        # Write data
        for user in flat_list:
            writer.writerow([
                user.get('hierarchy_level', 0),
                user.get('displayName', ''),
                user.get('userPrincipalName', ''),
                user.get('mail', ''),
                user.get('jobTitle', ''),
                user.get('department', ''),
                user.get('officeLocation', ''),
                '; '.join(user.get('businessPhones', [])),
                user.get('id', ''),
                user.get('manager_name', ''),
                user.get('manager_upn', ''),
                user.get('manager_id', '')
            ])
        
        # Prepare response
        output.seek(0)
        filename = f"organizational_hierarchy_{upn.replace('@', '_').replace('.', '_')}.csv"
        
        return Response(
            output.getvalue(),
            mimetype="text/csv",
            headers={"Content-Disposition": f"attachment;filename={filename}"}
        )
        
    except Exception as e:
        logger.exception("Error occurred during CSV export")
        return f"Error: {str(e)}", 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5002, debug=True)
