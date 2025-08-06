from flask import Flask, render_template, request, jsonify
import requests
from datetime import datetime, timedelta
import os

app = Flask(__name__)

# Configuration - Replace these with your actual credentials
AUTH_CONFIG = {
    'client_id': '73efa35d-6188-42d4-b258-838a977eb149',
    'client_secret': 'CyG8Q~FYHuCMSyVmt4sNxt5IejrMc2c24Ziz4a.t',
    'tenant_id': '0e439a1f-a497-462b-9e6b-4e582e203607',
    'access_token': '',
    'token_expires': None
}

class SharePointManager:
    def __init__(self):
        self.base_url = "https://graph.microsoft.com/v1.0"
    
    def _get_token(self):
        """Get or refresh access token"""
        if AUTH_CONFIG['access_token'] and datetime.now() < AUTH_CONFIG['token_expires']:
            return AUTH_CONFIG['access_token']
        
        url = f"https://login.microsoftonline.com/{AUTH_CONFIG['tenant_id']}/oauth2/v2.0/token"
        data = {
            'client_id': AUTH_CONFIG['client_id'],
            'client_secret': AUTH_CONFIG['client_secret'],
            'scope': 'https://graph.microsoft.com/.default',
            'grant_type': 'client_credentials'
        }
        
        response = requests.post(url, data=data)
        response.raise_for_status()
        token_data = response.json()
        
        AUTH_CONFIG['access_token'] = token_data['access_token']
        AUTH_CONFIG['token_expires'] = datetime.now() + timedelta(seconds=token_data['expires_in'])
        return AUTH_CONFIG['access_token']
    
    def _make_request(self, method, endpoint, json_data=None):
        """Make authenticated API request"""
        token = self._get_token()
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        url = f"{self.base_url}{endpoint}"
        response = requests.request(method, url, headers=headers, json=json_data)
        response.raise_for_status()
        return response.json() if response.content else None
    
    def search_sites(self, query):
        """Search SharePoint sites"""
        endpoint = f"/sites?search={query}" if query else "/sites?$top=10"
        result = self._make_request('GET', endpoint)
        return result.get('value', [])
    
    def get_site_details(self, site_id):
        """Get site details"""
        return self._make_request('GET', f"/sites/{site_id}")
    
    def get_site_permissions(self, site_id):
        """Get all permissions for a site"""
        return self._make_request('GET', f"/sites/{site_id}/permissions")
    
    def add_permission(self, site_id, app_id, roles):
        """Add new permission"""
        data = {
            "roles": roles,
            "grantedToIdentities": [{
                "application": {
                    "id": app_id,
                    "displayName": "Added via API"
                }
            }]
        }
        return self._make_request('POST', f"/sites/{site_id}/permissions", data)
    
    def delete_permission(self, site_id, permission_id):
        """Delete a permission"""
        self._make_request('DELETE', f"/sites/{site_id}/permissions/{permission_id}")
        return True

# Initialize SharePoint manager
sp_manager = SharePointManager()

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/search', methods=['POST'])
def search():
    query = request.form.get('query', '')
    sites = sp_manager.search_sites(query)
    return jsonify({'sites': sites})

@app.route('/site/<site_id>')
def site_details(site_id):
    site = sp_manager.get_site_details(site_id)
    return jsonify(site)

@app.route('/site/<site_id>/permissions')
def site_permissions(site_id):
    permissions = sp_manager.get_site_permissions(site_id)
    return jsonify({'permissions': permissions.get('value', [])})

@app.route('/site/<site_id>/permissions/add', methods=['POST'])
def add_site_permission(site_id):
    data = request.get_json()
    app_id = data.get('app_id')
    roles = data.get('roles', ['read'])
    result = sp_manager.add_permission(site_id, app_id, roles)
    return jsonify(result)

@app.route('/site/<site_id>/permissions/delete/<permission_id>', methods=['DELETE'])
def delete_site_permission(site_id, permission_id):
    success = sp_manager.delete_permission(site_id, permission_id)
    return jsonify({'success': success})

# HTML Template
html_template = '''
<!DOCTYPE html>
<html>
<head>
    <title>SharePoint Permissions Manager</title>
    <style>
        body { font-family: Arial, sans-serif; max-width: 1200px; margin: 0 auto; padding: 20px; }
        .container { display: flex; flex-direction: column; gap: 20px; }
        .search-section, .site-section, .permissions-section { padding: 20px; border: 1px solid #ddd; border-radius: 5px; }
        button { padding: 8px 16px; margin: 5px; cursor: pointer; }
        table { width: 100%; border-collapse: collapse; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        .hidden { display: none; }
        .permission-form { margin-top: 20px; }
        .permission-form input, .permission-form select { padding: 8px; margin: 5px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>SharePoint Permissions Manager</h1>
        
        <div class="search-section">
            <h2>Search Sites</h2>
            <input type="text" id="searchQuery" placeholder="Enter site name...">
            <button onclick="searchSites()">Search</button>
            <div id="searchResults"></div>
        </div>
        
        <div id="siteDetails" class="site-section hidden">
            <h2>Site Details</h2>
            <div id="siteInfo"></div>
            <button onclick="showAddPermission()">Add Permission</button>
            <button onclick="showPermissions()">Display Permissions</button>
            
            <div id="addPermissionForm" class="permission-form hidden">
                <h3>Add New Permission</h3>
                <input type="text" id="appId" placeholder="Application ID">
                <input type="text" id="appName" placeholder="Application Name">
                <select id="permissionRoles" multiple>
                    <option value="read">Read</option>
                    <option value="write">Write</option>
                    <option value="manage">Manage</option>
                </select>
                <button onclick="addPermission()">Add Permission</button>
            </div>
        </div>
        
        <div id="permissionsList" class="permissions-section hidden">
            <h2>Site Permissions</h2>
            <table id="permissionsTable">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Application</th>
                        <th>Roles</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody></tbody>
            </table>
        </div>
    </div>

    <script>
        let currentSiteId = '';
        
        function searchSites() {
            const query = document.getElementById('searchQuery').value;
            fetch('/search', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: `query=${encodeURIComponent(query)}`
            })
            .then(response => response.json())
            .then(data => {
                let html = '<h3>Search Results</h3><ul>';
                data.sites.forEach(site => {
                    html += `<li>
                        ${site.displayName || site.name} 
                        <button onclick="selectSite('${site.id}')">Select</button>
                    </li>`;
                });
                html += '</ul>';
                document.getElementById('searchResults').innerHTML = html;
            });
        }
        
        function selectSite(siteId) {
            currentSiteId = siteId;
            fetch(`/site/${siteId}`)
                .then(response => response.json())
                .then(site => {
                    document.getElementById('siteInfo').innerHTML = `
                        <p><strong>Name:</strong> ${site.displayName}</p>
                        <p><strong>URL:</strong> ${site.webUrl}</p>
                        <p><strong>ID:</strong> ${site.id}</p>
                    `;
                    document.getElementById('siteDetails').classList.remove('hidden');
                });
        }
        
        function showAddPermission() {
            document.getElementById('addPermissionForm').classList.remove('hidden');
            document.getElementById('permissionsList').classList.add('hidden');
        }
        
        function showPermissions() {
            document.getElementById('addPermissionForm').classList.add('hidden');
            fetch(`/site/${currentSiteId}/permissions`)
                .then(response => response.json())
                .then(data => {
                    const tbody = document.querySelector('#permissionsTable tbody');
                    tbody.innerHTML = '';
                    data.permissions.forEach(perm => {
                        const app = perm.grantedToIdentities?.[0]?.application || {};
                        const row = document.createElement('tr');
                        row.innerHTML = `
                            <td>${perm.id}</td>
                            <td>${app.displayName || app.id || 'N/A'}</td>
                            <td>${perm.roles?.join(', ') || 'N/A'}</td>
                            <td><button onclick="deletePermission('${perm.id}')">Delete</button></td>
                        `;
                        tbody.appendChild(row);
                    });
                    document.getElementById('permissionsList').classList.remove('hidden');
                });
        }
        
        function addPermission() {
            const appId = document.getElementById('appId').value;
            const roles = Array.from(document.getElementById('permissionRoles').selectedOptions)
                .map(opt => opt.value);
            
            fetch(`/site/${currentSiteId}/permissions/add`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ app_id: appId, roles })
            })
            .then(response => response.json())
            .then(data => {
                alert('Permission added successfully!');
                showPermissions();
            })
            .catch(error => alert('Error adding permission: ' + error));
        }
        
        function deletePermission(permissionId) {
            if (confirm('Are you sure you want to delete this permission?')) {
                fetch(`/site/${currentSiteId}/permissions/delete/${permissionId}`, {
                    method: 'DELETE'
                })
                .then(response => response.json())
                .then(data => {
                    if (data.success) {
                        showPermissions();
                    }
                });
            }
        }
    </script>
</body>
</html>
'''

# Create templates directory and save the HTML template
if not os.path.exists('templates'):
    os.makedirs('templates')

with open('templates/index.html', 'w', encoding='utf-8') as f:
    f.write(html_template)

if __name__ == '__main__':
    print("Starting SharePoint Permissions Manager...")
    print("Access the application at: http://localhost:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)
