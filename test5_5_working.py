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
from collections import deque
import logging
import os
import xml.dom.minidom

# History file path
HISTORY_FILE = 'api_history.json'

def save_history(url):
    """Save URL to history file"""
    try:
        # Load existing history
        history = []
        if os.path.exists(HISTORY_FILE):
            with open(HISTORY_FILE, 'r') as f:
                history = json.load(f)
        
        # Remove duplicates and add new URL at the beginning
        if url in history:
            history.remove(url)
        history.insert(0, url)
        
        # Keep only last 20 items
        history = history[:20]
        
        # Save back to file
        with open(HISTORY_FILE, 'w') as f:
            json.dump(history, f)
            
    except Exception as e:
        logger.error(f"Error saving history: {str(e)}")

def load_history():
    """Load history from file"""
    try:
        if os.path.exists(HISTORY_FILE):
            with open(HISTORY_FILE, 'r') as f:
                return json.load(f)
    except Exception as e:
        logger.error(f"Error loading history: {str(e)}")
    return []

app = Flask(__name__)
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Configuration for multiple tenants
TENANTS = {
    "tenant1": {
        "name": "GeekByte Online",
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
    },
    "tenant2": {
        "name": "Contoso",
        "tenant_id": "YOUR_SECOND_TENANT_ID",
        "tenant_name": "contoso.onmicrosoft.com",
        "app_id": "YOUR_SECOND_APP_ID",
        "client_secret": "YOUR_SECOND_CLIENT_SECRET",
        "certificate_path": "certificate2.pem",
        "private_key_path": "private_key2.pem",
        "scopes": {
            "graph": "https://graph.microsoft.com/.default",
            "sharepoint": "https://contoso.sharepoint.com/.default"
        }
    }
}

RECENT_URLS = deque(maxlen=20)
TOKEN_CACHE = {}  # Store tokens with expiration

PREDEFINED_URLS = {
    "Graph": [
        "https://graph.microsoft.com/v1.0/users",
        "https://graph.microsoft.com/v1.0/sites",
        "https://graph.microsoft.com/v1.0/me",
        "https://graph.microsoft.com/v1.0/groups"
    ],
    "SharePoint": [
        "https://graph.microsoft.com/v1.0/sites/geekbyteonline.sharepoint.com",
        "https://graph.microsoft.com/v1.0/sites/geekbyteonline.sharepoint.com/drives",
        "https://graph.microsoft.com/v1.0/sites/geekbyteonline.sharepoint.com/lists"
    ]
}

HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>Graph API Explorer</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Arial, sans-serif; font-size: 12px; background: #f5f5f5; }
        .container { max-width: 1200px; margin: 5px auto; background: white; padding: 8px; border-radius: 3px; }
        .header { background: #0078d4; color: white; padding: 6px; margin: -8px -8px 8px -8px; text-align: center; }
        .header h1 { font-size: 16px; margin: 0; }
        
        .config-row { display: flex; gap: 8px; margin-bottom: 8px; flex-wrap: wrap; }
        .config-group { background: #f8f9fa; border: 1px solid #ddd; padding: 6px; border-radius: 2px; }
        .config-group-title { font-weight: bold; font-size: 10px; color: #666; margin-bottom: 4px; }
        .config-options { display: flex; gap: 8px; flex-wrap: wrap; }
        .config-option { display: flex; align-items: center; }
        .config-option label { margin-left: 3px; font-size: 11px; }
        
        .url-row { display: flex; gap: 8px; margin-bottom: 8px; }
        .url-input { flex: 1; padding: 4px; border: 1px solid #ccc; border-radius: 2px; }
        .btn { padding: 4px 8px; background: #0078d4; color: white; border: none; border-radius: 2px; cursor: pointer; font-size: 11px; }
        .btn:hover { background: #106ebe; }
        
        .token-box { background: #f8f9fa; border: 1px solid #ddd; padding: 4px; margin: 4px 0; border-radius: 2px; }
        .token-display { font-family: monospace; font-size: 10px; color: #666; word-break: break-all; max-height: 40px; overflow-y: auto; }
        
        .shortcuts { background: #f8f9fa; border: 1px solid #ddd; padding: 4px; margin: 4px 0; border-radius: 2px; }
        .shortcuts-header { font-weight: bold; font-size: 10px; color: #0078d4; cursor: pointer; }
        .shortcuts-content { display: none; margin-top: 4px; }
        .url-group { margin-bottom: 4px; }
        .url-group-title { font-weight: bold; font-size: 10px; color: #666; }
        .url-item { background: #fff; border: 1px solid #ddd; padding: 2px 4px; margin: 1px 0; font-size: 9px; cursor: pointer; border-radius: 1px; }
        .url-item:hover { background: #e3f2fd; }
        
        .history { background: #f8f9fa; border: 1px solid #ddd; padding: 4px; margin: 4px 0; border-radius: 2px; }
        .history-header { font-weight: bold; font-size: 10px; color: #0078d4; cursor: pointer; }
        .history-content { display: none; margin-top: 4px; max-height: 80px; overflow-y: auto; }
        .history-item { background: #fff; border: 1px solid #ddd; padding: 2px 4px; margin: 1px 0; font-size: 9px; cursor: pointer; border-radius: 1px; }
        .history-item:hover { background: #e3f2fd; }
        
        .status { padding: 4px 6px; margin: 4px 0; border-radius: 2px; font-size: 11px; }
        .status-ready { background: #d1ecf1; color: #0c5460; }
        .status-loading { background: #fff3cd; color: #856404; }
        .status-success { background: #d4edda; color: #155724; }
        .status-error { background: #f8d7da; color: #721c24; }
        
        .response-area { border: 1px solid #ddd; border-radius: 2px; margin-top: 8px; }
        .response-header { background: #f8f9fa; padding: 4px 6px; border-bottom: 1px solid #ddd; display: flex; justify-content: space-between; align-items: center; }
        .tab-btn { padding: 2px 6px; background: #e9ecef; border: 1px solid #ccc; cursor: pointer; font-size: 10px; margin-right: 2px; }
        .tab-btn.active { background: #0078d4; color: white; }
        .response-content { padding: 6px; font-family: monospace; font-size: 10px; white-space: pre-wrap; word-break: break-word; min-height: 300px; background: #f8f9fa; }
        
        /* Table view styles */
        .table-view { 
            display: none; 
            width: 100%; 
            overflow: auto;
            max-height: 400px;
        }
        .data-table {
            width: 100%;
            border-collapse: collapse;
            font-family: Arial, sans-serif;
            font-size: 10px;
        }
        .data-table th {
            background-color: #0078d4;
            color: white;
            padding: 4px 8px;
            text-align: left;
            position: sticky;
            top: 0;
            white-space: nowrap;
        }
        .data-table td {
            padding: 4px 8px;
            border: 1px solid #ddd;
            vertical-align: top;
        }
        .data-table tr:nth-child(even) {
            background-color: #f2f2f2;
        }
        .data-table tr:hover {
            background-color: #e3f2fd;
        }
        .json-value {
            font-family: monospace;
        }
        .json-string {
            color: #d14;
        }
        .json-number {
            color: #099;
        }
        .json-boolean {
            color: #00c;
        }
        .json-null {
            color: #999;
        }
        .json-object {
            color: #708;
        }
        .json-array {
            color: #708;
        }
        .nested-table {
            width: 100%;
            margin: 0;
            padding: 0;
            border: none;
        }
        .nested-table td {
            border: none;
            padding: 2px 4px;
        }
        .expand-btn {
            cursor: pointer;
            color: #0078d4;
            font-weight: bold;
            margin-right: 5px;
        }
        
        .grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 4px; }
        
        @media (max-width: 768px) {
            .config-row, .url-row { flex-direction: column; }
            .config-options { flex-wrap: wrap; }
            .grid-2 { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Graph API Explorer</h1>
        </div>
        
        <div class="config-row">
            <div class="config-group">
                <div class="config-group-title">Tenant:</div>
                <div class="config-options">
                    <div class="config-option">
                        <input type="radio" id="tenant1" name="tenant" value="tenant1" checked>
                        <label for="tenant1">GeekByte</label>
                    </div>
                    <div class="config-option">
                        <input type="radio" id="tenant2" name="tenant" value="tenant2">
                        <label for="tenant2">Contoso</label>
                    </div>
                </div>
            </div>
            
            <div class="config-group">
                <div class="config-group-title">Method:</div>
                <div class="config-options">
                    <select id="http-method" style="padding: 3px; font-size: 11px;">
                        <option value="GET">GET</option>
                        <option value="POST">POST</option>
                        <option value="PUT">PUT</option>
                        <option value="DELETE">DELETE</option>
                    </select>
                </div>
            </div>
            
            <div class="config-group">
                <div class="config-group-title">Auth:</div>
                <div class="config-options">
                    <div class="config-option">
                        <input type="radio" id="auth-cert" name="auth-method" value="certificate" checked>
                        <label for="auth-cert">Certificate</label>
                    </div>
                    <div class="config-option">
                        <input type="radio" id="auth-secret" name="auth-method" value="secret">
                        <label for="auth-secret">Secret</label>
                    </div>
                </div>
            </div>
            
            <div class="config-group">
                <div class="config-group-title">Scope:</div>
                <div class="config-options">
                    <div class="config-option">
                        <input type="radio" id="scope-graph" name="scope" value="graph" checked>
                        <label for="scope-graph">Graph</label>
                    </div>
                    <div class="config-option">
                        <input type="radio" id="scope-sp" name="scope" value="sharepoint">
                        <label for="scope-sp">SharePoint</label>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="url-row">
            <input type="text" id="api-url" class="url-input" value="https://graph.microsoft.com/v1.0/users" placeholder="Enter API URL">
            <button class="btn" onclick="callApi()">Send</button>
        </div>
        
        <div id="post-options" style="display: none; margin-bottom: 8px;">
            <div style="display: flex; gap: 8px;">
                <div style="flex: 1;">
                    <div style="font-size: 10px; font-weight: bold; margin-bottom: 2px;">Headers (JSON):</div>
                    <textarea id="headers-input" style="width: 100%; height: 60px; font-family: monospace; font-size: 10px; padding: 4px; border: 1px solid #ccc; border-radius: 2px;" placeholder='{"Content-Type": "application/json"}'></textarea>
                </div>
                <div style="flex: 1;">
                    <div style="font-size: 10px; font-weight: bold; margin-bottom: 2px;">Body (JSON):</div>
                    <textarea id="body-input" style="width: 100%; height: 60px; font-family: monospace; font-size: 10px; padding: 4px; border: 1px solid #ccc; border-radius: 2px;" placeholder='{"key": "value"}'></textarea>
                </div>
            </div>
        </div>
        
        <div class="token-box">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 2px;">
                <div style="font-size: 10px; font-weight: bold;">Token:</div>
                <div>
                    <button class="btn" onclick="copyToken()" style="font-size: 9px; padding: 2px 4px;">Copy</button>
                    <button class="btn" onclick="refreshToken()" style="font-size: 9px; padding: 2px 4px;">Refresh</button>
                </div>
            </div>
            <div class="token-display" id="token-display">No token obtained yet</div>
        </div>
        
        <div class="grid-2">
            <div class="shortcuts">
                <div class="shortcuts-header" onclick="toggleSection('shortcuts')">🔗 Quick URLs</div>
                <div class="shortcuts-content" id="shortcuts-content">
                    {% for category, urls in predefined_urls.items() %}
                    <div class="url-group">
                        <div class="url-group-title">{{ category }}</div>
                        {% for url in urls %}
                        <div class="url-item" onclick="setUrl('{{ url }}')">{{ url }}</div>
                        {% endfor %}
                    </div>
                    {% endfor %}
                </div>
            </div>
            
            <div class="history">
                <div class="history-header" onclick="toggleSection('history')">📋 History <span id="history-count">(0)</span></div>
                <div class="history-content" id="history-content">
                    <div style="color: #666; text-align: center; padding: 10px; font-size: 9px;">No history</div>
                </div>
            </div>
        </div>
        
        <div id="status" class="status status-ready">Ready</div>
        
        <div class="response-area">
            <div class="response-header">
                <div>
                    <button class="tab-btn active" onclick="switchTab('formatted')">Formatted</button>
                    <button class="tab-btn" onclick="switchTab('raw')">Raw</button>
                    <button class="tab-btn" onclick="switchTab('table')">Table</button>
                </div>
                <button class="btn" onclick="copyResponse()">Copy</button>
            </div>
            <div class="response-content" id="response-formatted">Response will appear here...</div>
            <div class="response-content" id="response-raw" style="display: none;">Raw response will appear here...</div>
            <div class="response-content table-view" id="response-table">
                <table class="data-table" id="data-table">
                    <!-- Table content will be generated here -->
                </table>
            </div>
        </div>
    </div>

    <script>
        let currentResponse = '';
        let currentRawResponse = '';
        let currentJsonData = null;
        let activeTab = 'formatted';
        let currentToken = '';
        let tokenCache = {};
        
        // Initialize history from server-side data
        window.apiHistory = [];
        
        // Load initial history
        fetch('/load_history')
            .then(response => response.json())
            .then(history => {
                window.apiHistory = history;
                renderHistory();
            })
            .catch(error => console.error('Error loading history:', error));
        
        // Show/hide POST options based on method selection
        document.getElementById('http-method').addEventListener('change', function() {
            const postOptions = document.getElementById('post-options');
            const method = this.value;
            postOptions.style.display = (method === 'POST' || method === 'PUT') ? 'block' : 'none';
        });
        
        function setUrl(url) {
            document.getElementById('api-url').value = url;
        }
        
        function toggleSection(sectionName) {
            const content = document.getElementById(sectionName + '-content');
            const isVisible = content.style.display === 'block';
            
            // Close all sections
            document.querySelectorAll('.shortcuts-content, .history-content').forEach(el => {
                el.style.display = 'none';
            });
            
            // Open clicked section if it wasn't visible
            if (!isVisible) {
                content.style.display = 'block';
            }
        }
        
        function switchTab(tabName) {
            const tabs = document.querySelectorAll('.tab-btn');
            const formattedDiv = document.getElementById('response-formatted');
            const rawDiv = document.getElementById('response-raw');
            const tableView = document.getElementById('response-table');
            
            tabs.forEach(tab => tab.classList.remove('active'));
            
            if (tabName === 'formatted') {
                tabs[0].classList.add('active');
                formattedDiv.style.display = 'block';
                rawDiv.style.display = 'none';
                tableView.style.display = 'none';
                activeTab = 'formatted';
            } else if (tabName === 'raw') {
                tabs[1].classList.add('active');
                formattedDiv.style.display = 'none';
                rawDiv.style.display = 'block';
                tableView.style.display = 'none';
                activeTab = 'raw';
            } else {
                tabs[2].classList.add('active');
                formattedDiv.style.display = 'none';
                rawDiv.style.display = 'none';
                tableView.style.display = 'block';
                activeTab = 'table';
                renderTable();
            }
        }
        
        function updateTokenDisplay(token) {
            const tokenDisplay = document.getElementById('token-display');
            if (token) {
                currentToken = token;
                // Show first 20 and last 20 characters
                const truncated = token.length > 40 ? 
                    token.substring(0, 20) + '...' + token.substring(token.length - 20) : token;
                tokenDisplay.textContent = truncated;
                tokenDisplay.title = token; // Full token in tooltip
            } else {
                tokenDisplay.textContent = 'Failed to obtain token';
                tokenDisplay.title = '';
            }
        }
        
        function copyToken() {
            if (currentToken) {
                navigator.clipboard.writeText(currentToken).then(() => {
                    setStatus('success', 'Token copied!');
                    setTimeout(() => setStatus('ready', 'Ready'), 2000);
                });
            } else {
                setStatus('error', 'No token to copy');
            }
        }
        
        function refreshToken() {
            // Clear token cache and get new token
            tokenCache = {};
            setStatus('loading', 'Refreshing token...');
            
            const tenant = document.querySelector('input[name="tenant"]:checked').value;
            const authMethod = document.querySelector('input[name="auth-method"]:checked').value;
            const scope = document.querySelector('input[name="scope"]:checked').value;
            
            fetch('/get_token', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    tenant: tenant,
                    auth_method: authMethod,
                    scope: scope,
                    force_refresh: true
                })
            })
            .then(res => res.json())
            .then(data => {
                if (data.token) {
                    updateTokenDisplay(data.token);
                    setStatus('success', 'Token refreshed');
                } else {
                    setStatus('error', 'Failed to refresh token');
                }
            })
            .catch(err => {
                setStatus('error', 'Token refresh failed');
            });
        }
        
        function callApi() {
            const url = document.getElementById('api-url').value.trim();
            const method = document.getElementById('http-method').value;
            const tenant = document.querySelector('input[name="tenant"]:checked').value;
            const authMethod = document.querySelector('input[name="auth-method"]:checked').value;
            const scope = document.querySelector('input[name="scope"]:checked').value;
            const statusDiv = document.getElementById('status');
            const formattedResponseDiv = document.getElementById('response-formatted');
            const rawResponseDiv = document.getElementById('response-raw');
            const tableView = document.getElementById('response-table');
            
            if (!url) {
                setStatus('error', 'Enter URL');
                return;
            }

            // Get headers and body for POST/PUT requests
            let headers = {};
            let body = null;
            
            if (method === 'POST' || method === 'PUT') {
                const headersInput = document.getElementById('headers-input').value.trim();
                const bodyInput = document.getElementById('body-input').value.trim();
                
                if (headersInput) {
                    try {
                        headers = JSON.parse(headersInput);
                    } catch (e) {
                        setStatus('error', 'Invalid JSON in headers');
                        return;
                    }
                }
                
                if (bodyInput) {
                    try {
                        body = JSON.parse(bodyInput);
                    } catch (e) {
                        setStatus('error', 'Invalid JSON in body');
                        return;
                    }
                }
            }

            setStatus('loading', 'Sending...');
            formattedResponseDiv.textContent = 'Loading...';
            rawResponseDiv.textContent = 'Loading...';
            tableView.innerHTML = '<table class="data-table" id="data-table"></table>';

            fetch('/call_graph', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    url: url,
                    method: method,
                    tenant: tenant,
                    auth_method: authMethod,
                    scope: scope,
                    headers: headers,
                    body: body
                })
            })
            .then(res => res.json())
            .then(data => {
                // Update token display
                if (data.token) {
                    updateTokenDisplay(data.token);
                }
                
                if (data.error) {
                    setStatus('error', 'Error: ' + data.error);
                    const errorResponse = JSON.stringify(data, null, 2);
                    formattedResponseDiv.textContent = errorResponse;
                    rawResponseDiv.textContent = errorResponse;
                    currentResponse = errorResponse;
                    currentRawResponse = errorResponse;
                    currentJsonData = null;
                } else {
                    setStatus('success', 'Success (' + data.status_code + ')');
                    
                    currentRawResponse = data.raw_response || JSON.stringify(data.response, null, 2);
                    currentResponse = data.formatted_response || JSON.stringify(data.response, null, 2);
                    currentJsonData = data.response;
                    
                    formattedResponseDiv.textContent = currentResponse;
                    rawResponseDiv.textContent = currentRawResponse;
                    
                    // Add to history if request was successful
                    if (data.status_code < 400) {
                        addToHistory(url);
                    }
                    
                    // If table view is active, render it
                    if (activeTab === 'table') {
                        renderTable();
                    }
                }
            })
            .catch(err => {
                setStatus('error', 'Request failed');
                const errorMsg = 'Error: ' + err.message;
                formattedResponseDiv.textContent = errorMsg;
                rawResponseDiv.textContent = errorMsg;
                currentResponse = errorMsg;
                currentRawResponse = errorMsg;
                currentJsonData = null;
            });
        }
        
        function setStatus(type, message) {
            const statusDiv = document.getElementById('status');
            statusDiv.className = 'status status-' + type;
            statusDiv.textContent = message;
        }
        
        function addToHistory(url) {
            if (!window.apiHistory) {
                window.apiHistory = [];
            }
            
            // Remove if exists
            window.apiHistory = window.apiHistory.filter(item => item !== url);
            
            // Add to beginning
            window.apiHistory.unshift(url);
            
            // Keep only last 20
            window.apiHistory = window.apiHistory.slice(0, 20);
            
            // Save to server
            fetch('/save_history', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ url: url })
            })
            .catch(error => {
                console.error('Error saving history:', error);
                setStatus('error', 'Failed to save to history');
            });
            
            renderHistory();
        }
        
        function renderHistory() {
            const container = document.getElementById('history-content');
            const countSpan = document.getElementById('history-count');
            const history = window.apiHistory || [];
            
            countSpan.textContent = `(${history.length})`;
            
            if (history.length === 0) {
                container.innerHTML = '<div style="color: #666; text-align: center; padding: 10px; font-size: 9px;">No history</div>';
                return;
            }
            
            container.innerHTML = '';
            history.forEach(url => {
                const div = document.createElement('div');
                div.className = 'history-item';
                div.textContent = url;
                div.onclick = () => setUrl(url);
                container.appendChild(div);
            });
        }
        
        function copyResponse() {
            let responseText;
            if (activeTab === 'formatted') {
                responseText = currentResponse;
            } else if (activeTab === 'raw') {
                responseText = currentRawResponse;
            } else {
                responseText = document.getElementById('data-table').outerHTML;
            }
            
            navigator.clipboard.writeText(responseText).then(() => {
                setStatus('success', 'Copied!');
                setTimeout(() => setStatus('ready', 'Ready'), 2000);
            });
        }
        
        function renderTable() {
            if (!currentJsonData) {
                document.getElementById('data-table').innerHTML = '<tr><td>No data available</td></tr>';
                return;
            }
            
            const table = document.getElementById('data-table');
            table.innerHTML = '';
            
            try {
                // Check if the response is an array of items (like users, groups, etc.)
                if (Array.isArray(currentJsonData)) {
                    renderArrayAsTable(table, currentJsonData);
                } 
                // Check if the response is an object with a 'value' property that's an array
                else if (currentJsonData.value && Array.isArray(currentJsonData.value)) {
                    renderArrayAsTable(table, currentJsonData.value);
                } 
                // Check if it's a single object
                else if (typeof currentJsonData === 'object') {
                    renderObjectAsTable(table, currentJsonData);
                }
                // Handle other cases
                else {
                    table.innerHTML = '<tr><td>Data format not supported for table view</td></tr>';
                }
            } catch (e) {
                console.error('Error rendering table:', e);
                table.innerHTML = '<tr><td>Error rendering table view</td></tr>';
            }
        }
        
        function renderArrayAsTable(table, items) {
            if (!items || items.length === 0) {
                table.innerHTML = '<tr><td>No data available</td></tr>';
                return;
            }
            
            // Collect all possible property names from all items
            const allProperties = new Set();
            items.forEach(item => {
                if (typeof item === 'object') {
                    Object.keys(item).forEach(prop => allProperties.add(prop));
                }
            });
            
            if (allProperties.size === 0) {
                table.innerHTML = '<tr><td>No properties found in data</td></tr>';
                return;
            }
            
            // Create header row
            const headerRow = document.createElement('tr');
            Array.from(allProperties).forEach(prop => {
                const th = document.createElement('th');
                th.textContent = prop;
                headerRow.appendChild(th);
            });
            table.appendChild(headerRow);
            
            // Create data rows
            items.forEach(item => {
                const row = document.createElement('tr');
                
                Array.from(allProperties).forEach(prop => {
                    const td = document.createElement('td');
                    const value = item[prop];
                    td.innerHTML = formatValueForTable(value, prop);
                    row.appendChild(td);
                });
                
                table.appendChild(row);
            });
        }
        
        function renderObjectAsTable(table, obj, depth = 0) {
            if (!obj || Object.keys(obj).length === 0) {
                table.innerHTML = '<tr><td>No data available</td></tr>';
                return;
            }
            
            // Create header row if top level
            if (depth === 0) {
                const headerRow = document.createElement('tr');
                const th1 = document.createElement('th');
                th1.textContent = 'Property';
                const th2 = document.createElement('th');
                th2.textContent = 'Value';
                headerRow.appendChild(th1);
                headerRow.appendChild(th2);
                table.appendChild(headerRow);
            }
            
            // Create data rows
            Object.keys(obj).forEach(key => {
                const row = document.createElement('tr');
                
                // Property name cell
                const tdKey = document.createElement('td');
                tdKey.textContent = key;
                if (depth > 0) {
                    tdKey.style.paddingLeft = (depth * 10) + 'px';
                }
                row.appendChild(tdKey);
                
                // Value cell
                const tdValue = document.createElement('td');
                const value = obj[key];
                
                if (typeof value === 'object' && value !== null && !Array.isArray(value)) {
                    // For nested objects, create a nested table
                    const expandBtn = document.createElement('span');
                    expandBtn.className = 'expand-btn';
                    expandBtn.textContent = '▶';
                    expandBtn.onclick = function(e) {
                        e.stopPropagation();
                        const nestedTable = this.nextSibling;
                        if (nestedTable.style.display === 'none') {
                            nestedTable.style.display = 'table';
                            this.textContent = '▼';
                        } else {
                            nestedTable.style.display = 'none';
                            this.textContent = '▶';
                        }
                    };
                    tdValue.appendChild(expandBtn);
                    
                    const nestedTable = document.createElement('table');
                    nestedTable.className = 'nested-table';
                    nestedTable.style.display = 'none';
                    renderObjectAsTable(nestedTable, value, depth + 1);
                    tdValue.appendChild(nestedTable);
                    
                    // Add a preview of the first few properties
                    const preview = document.createElement('span');
                    preview.className = 'json-object';
                    const keys = Object.keys(value);
                    preview.textContent = `{${keys.slice(0, 3).join(', ')}${keys.length > 3 ? '...' : ''}}`;
                    tdValue.appendChild(preview);
                } else {
                    tdValue.innerHTML = formatValueForTable(value, key);
                }
                
                row.appendChild(tdValue);
                table.appendChild(row);
            });
        }
        
        function formatValueForTable(value, key, depth = 0) {
            if (value === undefined || value === null) {
                return '<span class="json-value json-null">null</span>';
            } else if (typeof value === 'string') {
                // Format dates
                if (key.toLowerCase().includes('date') || key.toLowerCase().includes('time')) {
                    try {
                        const date = new Date(value);
                        return `<span class="json-value json-string">${date.toLocaleString()}</span>`;
                    } catch (e) {
                        return `<span class="json-value json-string">"${escapeHtml(value)}"</span>`;
                    }
                }
                // Format URLs
                if (key.toLowerCase().includes('url') && value.startsWith('http')) {
                    return `<a href="${escapeHtml(value)}" target="_blank" class="json-value json-string">"${escapeHtml(value)}"</a>`;
                }
                // Format emails
                if (key.toLowerCase().includes('email') && value.includes('@')) {
                    return `<a href="mailto:${escapeHtml(value)}" class="json-value json-string">"${escapeHtml(value)}"</a>`;
                }
                return `<span class="json-value json-string">"${escapeHtml(value)}"</span>`;
            } else if (typeof value === 'number') {
                // Format sizes
                if (key.toLowerCase().includes('size') || key.toLowerCase().includes('bytes')) {
                    return `<span class="json-value json-number">${formatBytes(value)}</span>`;
                }
                // Format other numbers
                return `<span class="json-value json-number">${value.toLocaleString()}</span>`;
            } else if (typeof value === 'boolean') {
                return `<span class="json-value json-boolean">${value}</span>`;
            } else if (Array.isArray(value)) {
                if (value.length === 0) {
                    return '<span class="json-value json-array">[ ]</span>';
                }
                // Create expandable array
                const id = `array-${Math.random().toString(36).substr(2, 9)}`;
                let preview = '';
                if (value.every(v => typeof v !== 'object')) {
                    preview = value.slice(0, 3).map(v => {
                        if (typeof v === 'string') return `"${v}"`;
                        return v;
                    }).join(', ');
                }
                return `
                    <div class="expandable-container">
                        <span class="expand-btn" onclick="toggleExpand('${id}')">▶</span>
                        <span class="json-value json-array">[${preview}${value.length > 3 ? ', ...' : ''}] (${value.length})</span>
                        <div id="${id}" class="expanded-content" style="display:none">
                            ${value.map((item, i) => `
                                <div class="array-item">
                                    <span class="array-index">${i}:</span>
                                    ${formatValueForTable(item, key, depth + 1)}
                                </div>
                            `).join('')}
                        </div>
                    </div>
                `;
            } else if (typeof value === 'object') {
                // Create expandable object
                const id = `obj-${Math.random().toString(36).substr(2, 9)}`;
                const keys = Object.keys(value);
                if (keys.length === 0) {
                    return '<span class="json-value json-object">{ }</span>';
                }
                return `
                    <div class="expandable-container" style="margin-left: ${depth * 10}px">
                        <span class="expand-btn" onclick="toggleExpand('${id}')">▶</span>
                        <span class="json-value json-object">{${keys.slice(0, 3).join(', ')}${keys.length > 3 ? ', ...' : ''}}</span>
                        <div id="${id}" class="expanded-content" style="display:none">
                            <table class="nested-table">
                                ${keys.map(k => `
                                    <tr>
                                        <td class="nested-key">${k}</td>
                                        <td class="nested-value">${formatValueForTable(value[k], k, depth + 1)}</td>
                                    </tr>
                                `).join('')}
                            </table>
                        </div>
                    </div>
                `;
            }
            return escapeHtml(String(value));
        }

        // Add these helper functions to your script
        function toggleExpand(id) {
            const element = document.getElementById(id);
            const btn = element.previousElementSibling.previousElementSibling;
            if (element.style.display === 'none') {
                element.style.display = 'block';
                btn.textContent = '▼';
            } else {
                element.style.display = 'none';
                btn.textContent = '▶';
            }
        }

        function formatBytes(bytes, decimals = 2) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const dm = decimals < 0 ? 0 : decimals;
            const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB', 'PB', 'EB', 'ZB', 'YB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
        }

        function escapeHtml(unsafe) {
            return unsafe
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#039;");
        }


                
        // Enter key to submit
        document.getElementById('api-url').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                callApi();
            }
        });
        
        // Initialize the table view
        document.getElementById('data-table').innerHTML = '<tr><td>Make a request to see data in table view</td></tr>';
    </script>
</body>
</html>
"""

def format_xml_response(xml_string):
    """Format XML response with proper indentation and line breaks"""
    try:
        dom = xml.dom.minidom.parseString(xml_string)
        return dom.toprettyxml(indent="  ")
    except Exception:
        return xml_string

def is_xml_content(content_type, response_text):
    """Check if the response is XML content"""
    if content_type and ('xml' in content_type.lower() or 'application/xml' in content_type.lower()):
        return True
    
    # Check if the response text looks like XML
    if response_text.strip().startswith('<?xml') or response_text.strip().startswith('<'):
        return True
    
    return False

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE, predefined_urls=PREDEFINED_URLS)

@app.route('/load_history')
def get_history():
    """Get saved API history"""
    try:
        history = load_history()
        return jsonify(history)
    except Exception as e:
        logger.exception("Error loading history")
        return jsonify([])

@app.route('/save_history', methods=['POST'])
def save_history_endpoint():
    """Save a single URL to history"""
    try:
        if not request.is_json:
            return jsonify({"error": "Content-Type must be application/json"}), 400
            
        data = request.get_json()
        if not data or 'url' not in data:
            return jsonify({"error": "URL not provided"}), 400
            
        url = data['url']
        save_history(url)
        return jsonify({"message": "URL saved successfully"})
    except Exception as e:
        logger.exception("Error saving history")
        return jsonify({"error": str(e)}), 500

@app.route('/get_token', methods=['POST'])
def get_token():
    """Get or refresh access token"""
    try:
        if not request.is_json:
            return jsonify({"error": "Content-Type must be application/json"}), 400
            
        data = request.get_json()
        if not data:
            return jsonify({"error": "No data provided"}), 400
            
        tenant_id = data.get('tenant', 'tenant1')
        auth_method = data.get('auth_method', 'certificate')
        scope_type = data.get('scope', 'graph')
        force_refresh = data.get('force_refresh', False)
        
        # Create cache key
        cache_key = f"{tenant_id}_{auth_method}_{scope_type}"
        
        # Check if we have a valid cached token
        if not force_refresh and cache_key in TOKEN_CACHE:
            token_data = TOKEN_CACHE[cache_key]
            if time.time() < token_data['expires_at']:
                return jsonify({"token": token_data['token']})
        
        # Get new token
        tenant_config = TENANTS.get(tenant_id, TENANTS['tenant1'])
        scope = tenant_config['scopes'][scope_type]
        
        if auth_method == 'certificate':
            access_token = get_token_with_certificate(tenant_config, scope)
        else:
            access_token = get_token_with_secret(tenant_config, scope)
        
        if access_token:
            # Cache token (expires in 55 minutes to be safe)
            TOKEN_CACHE[cache_key] = {
                'token': access_token,
                'expires_at': time.time() + 3300  # 55 minutes
            }
            return jsonify({"token": access_token})
        else:
            return jsonify({"error": "Failed to obtain token"}), 500
            
    except Exception as e:
        logger.exception("Error getting token")
        return jsonify({"error": str(e)}), 500

@app.route('/call_graph', methods=['POST'])
def handle_call_graph():
    """Handle Graph API calls"""
    try:
        if not request.is_json:
            return jsonify({"error": "Content-Type must be application/json"}), 400
            
        data = request.get_json()
        if not data:
            return jsonify({"error": "No data provided"}), 400
            
        api_url = data.get('url')
        if not api_url:
            return jsonify({"error": "URL not provided"}), 400
            
        method = data.get('method', 'GET')
        tenant_id = data.get('tenant', 'tenant1')
        auth_method = data.get('auth_method', 'certificate')
        scope_type = data.get('scope', 'graph')
        custom_headers = data.get('headers', {})
        body = data.get('body')
        
        # Get tenant config
        tenant_config = TENANTS.get(tenant_id, TENANTS['tenant1'])
        
        # Create cache key
        cache_key = f"{tenant_id}_{auth_method}_{scope_type}"
        
        # Get token (either from cache or new)
        access_token = None
        if cache_key in TOKEN_CACHE:
            token_data = TOKEN_CACHE[cache_key]
            if time.time() < token_data['expires_at']:
                access_token = token_data['token']
        
        if not access_token:
            scope = tenant_config['scopes'][scope_type]
            if auth_method == 'certificate':
                access_token = get_token_with_certificate(tenant_config, scope)
            else:
                access_token = get_token_with_secret(tenant_config, scope)
            
            if access_token:
                TOKEN_CACHE[cache_key] = {
                    'token': access_token,
                    'expires_at': time.time() + 3300
                }
        
        if not access_token:
            return jsonify({"error": "Failed to obtain access token"}), 500
        
        # Prepare headers
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        headers.update(custom_headers)
        
        # Make API request
        try:
            if method == 'GET':
                response = requests.get(api_url, headers=headers)
            elif method == 'POST':
                response = requests.post(api_url, headers=headers, json=body)
            elif method == 'PUT':
                response = requests.put(api_url, headers=headers, json=body)
            elif method == 'DELETE':
                response = requests.delete(api_url, headers=headers)
            else:
                return jsonify({"error": f"Unsupported method: {method}"}), 400
        except requests.exceptions.RequestException as e:
            return jsonify({"error": f"Request failed: {str(e)}"}), 500
        
        # Handle token expiration
        if response.status_code == 401:
            scope = tenant_config['scopes'][scope_type]
            new_token = get_token_with_certificate(tenant_config, scope) if auth_method == 'certificate' else get_token_with_secret(tenant_config, scope)
            
            if new_token:
                TOKEN_CACHE[cache_key] = {
                    'token': new_token,
                    'expires_at': time.time() + 3300
                }
                headers["Authorization"] = f"Bearer {new_token}"
                
                try:
                    if method == 'GET':
                        response = requests.get(api_url, headers=headers)
                    elif method == 'POST':
                        response = requests.post(api_url, headers=headers, json=body)
                    elif method == 'PUT':
                        response = requests.put(api_url, headers=headers, json=body)
                    elif method == 'DELETE':
                        response = requests.delete(api_url, headers=headers)
                except requests.exceptions.RequestException as e:
                    return jsonify({"error": f"Request failed after token refresh: {str(e)}"}), 500
                
                access_token = new_token
        
        # Process response
        content_type = response.headers.get('content-type', '')
        raw_response = response.text
        
        # Try to parse response
        try:
            response_data = response.json()
            formatted_response = json.dumps(response_data, indent=2, ensure_ascii=False)
        except ValueError:
            if is_xml_content(content_type, raw_response):
                response_data = {"content_type": "XML", "raw": raw_response}
                formatted_response = format_xml_response(raw_response)
            else:
                response_data = {"content_type": content_type or "text/plain", "raw": raw_response}
                formatted_response = raw_response

        # Save successful requests to history
        if response.status_code < 400:
            try:
                save_history(api_url)
            except Exception as e:
                logger.error(f"Failed to save history: {str(e)}")

        # Return response
        result = {
            "status_code": response.status_code,
            "response": response_data,
            "formatted_response": formatted_response,
            "raw_response": raw_response,
            "content_type": content_type,
            "auth_method": auth_method,
            "scope": scope_type,
            "tenant": tenant_id,
            "token": access_token,
            "method": method
        }
        
        return jsonify(result)

    except Exception as e:
        logger.exception("Error in call_graph")
        return jsonify({"error": str(e), "token": None}), 500

def get_token_with_certificate(tenant_config, scope):
    """Get access token using certificate-based authentication"""
    try:
        if not os.path.exists(tenant_config['certificate_path']) or not os.path.exists(tenant_config['private_key_path']):
            raise Exception("Certificate or private key file not found")
            
        with open(tenant_config['certificate_path'], "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
        with open(tenant_config['private_key_path'], "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())

        now = int(time.time())
        jwt_header = {
            "alg": "RS256",
            "typ": "JWT",
            "x5t": base64.urlsafe_b64encode(certificate.fingerprint(hashes.SHA1())).decode().rstrip('=')
        }
        jwt_payload = {
            "aud": f"https://login.microsoftonline.com/{tenant_config['tenant_id']}/oauth2/v2.0/token",
            "exp": now + 300,
            "iss": tenant_config['app_id'],
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": tenant_config['app_id']
        }

        encoded_header = base64.urlsafe_b64encode(json.dumps(jwt_header).encode()).decode().rstrip('=')
        encoded_payload = base64.urlsafe_b64encode(json.dumps(jwt_payload).encode()).decode().rstrip('=')
        jwt_unsigned = f"{encoded_header}.{encoded_payload}"
        signature = private_key.sign(jwt_unsigned.encode(), padding.PKCS1v15(), hashes.SHA256())
        encoded_signature = base64.urlsafe_b64encode(signature).decode().rstrip('=')
        jwt = f"{jwt_unsigned}.{encoded_signature}"

        token_response = requests.post(
            f"https://login.microsoftonline.com/{tenant_config['tenant_id']}/oauth2/v2.0/token",
            data={
                "client_id": tenant_config['app_id'],
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

def get_token_with_secret(tenant_config, scope):
    """Get access token using client secret authentication"""
    try:
        token_url = f"https://login.microsoftonline.com/{tenant_config['tenant_id']}/oauth2/v2.0/token"
        
        token_data = {
            "client_id": tenant_config['app_id'],
            "client_secret": tenant_config['client_secret'],
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

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)