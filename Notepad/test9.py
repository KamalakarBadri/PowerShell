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

app = Flask(__name__)
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Configuration for both authentication methods
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

RECENT_URLS = deque(maxlen=20)
PREDEFINED_URLS = {
    "Microsoft Graph": [
        "https://graph.microsoft.com/v1.0/users",
        "https://graph.microsoft.com/v1.0/me",
        "https://graph.microsoft.com/v1.0/groups",
        "https://graph.microsoft.com/v1.0/applications",
        "https://graph.microsoft.com/v1.0/sites",
        "https://graph.microsoft.com/v1.0/drives",
        "https://graph.microsoft.com/v1.0/me/drive/root/children"
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
    <title>Microsoft Graph API Explorer</title>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            padding: 20px;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        
        h1 {
            color: #333;
            text-align: center;
            margin-bottom: 30px;
        }
        
        .section {
            margin-bottom: 20px;
            padding: 15px;
            border: 1px solid #ddd;
            border-radius: 3px;
        }
        
        .section h3 {
            margin-top: 0;
            color: #555;
        }
        
        .form-group {
            margin-bottom: 15px;
        }
        
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
        }
        
        input[type="text"], textarea, select {
            width: 100%;
            padding: 8px;
            border: 1px solid #ccc;
            border-radius: 3px;
            box-sizing: border-box;
        }
        
        .radio-group {
            display: flex;
            gap: 20px;
            flex-wrap: wrap;
        }
        
        .radio-item {
            display: flex;
            align-items: center;
            gap: 5px;
        }
        
        button {
            background-color: #007cba;
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 3px;
            cursor: pointer;
            font-size: 14px;
        }
        
        button:hover {
            background-color: #005a87;
        }
        
        .url-list {
            max-height: 200px;
            overflow-y: auto;
            border: 1px solid #ddd;
            padding: 10px;
            background-color: #f9f9f9;
        }
        
        .url-item {
            padding: 5px;
            cursor: pointer;
            border-bottom: 1px solid #eee;
            font-size: 12px;
        }
        
        .url-item:hover {
            background-color: #e9e9e9;
        }
        
        .status {
            padding: 10px;
            margin: 10px 0;
            border-radius: 3px;
            font-weight: bold;
        }
        
        .status-success {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        
        .status-error {
            background-color: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        
        .status-loading {
            background-color: #fff3cd;
            color: #856404;
            border: 1px solid #ffeaa7;
        }
        
        .response-area {
            background-color: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 3px;
            padding: 15px;
            min-height: 300px;
            font-family: monospace;
            font-size: 12px;
            white-space: pre-wrap;
            word-wrap: break-word;
            overflow: auto;
        }
        
        .tabs {
            display: flex;
            gap: 5px;
            margin-bottom: 10px;
        }
        
        .tab-btn {
            padding: 8px 16px;
            background-color: #e9ecef;
            border: 1px solid #ccc;
            cursor: pointer;
            border-radius: 3px 3px 0 0;
        }
        
        .tab-btn.active {
            background-color: #007cba;
            color: white;
        }
        
        .history-list {
            max-height: 150px;
            overflow-y: auto;
            border: 1px solid #ddd;
            background-color: #f9f9f9;
        }
        
        .history-item {
            padding: 5px 10px;
            cursor: pointer;
            border-bottom: 1px solid #eee;
            font-size: 12px;
        }
        
        .history-item:hover {
            background-color: #e9e9e9;
        }
        
        .payload-section {
            display: none;
        }
        
        .payload-section.show {
            display: block;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Microsoft Graph API Explorer</h1>
        
        <div class="section">
            <h3>Configuration</h3>
            
            <div class="form-group">
                <label>Authentication Method:</label>
                <div class="radio-group">
                    <div class="radio-item">
                        <input type="radio" id="auth-certificate" name="auth-method" value="certificate" checked>
                        <label for="auth-certificate">Certificate</label>
                    </div>
                    <div class="radio-item">
                        <input type="radio" id="auth-secret" name="auth-method" value="secret">
                        <label for="auth-secret">Client Secret</label>
                    </div>
                </div>
            </div>
            
            <div class="form-group">
                <label>Scope:</label>
                <div class="radio-group">
                    <div class="radio-item">
                        <input type="radio" id="scope-graph" name="scope" value="graph" checked>
                        <label for="scope-graph">Microsoft Graph</label>
                    </div>
                    <div class="radio-item">
                        <input type="radio" id="scope-sharepoint" name="scope" value="sharepoint">
                        <label for="scope-sharepoint">SharePoint</label>
                    </div>
                </div>
            </div>
            
            <div class="form-group">
                <label>HTTP Method:</label>
                <div class="radio-group">
                    <div class="radio-item">
                        <input type="radio" id="method-get" name="http-method" value="GET" checked onchange="togglePayloadSection()">
                        <label for="method-get">GET</label>
                    </div>
                    <div class="radio-item">
                        <input type="radio" id="method-post" name="http-method" value="POST" onchange="togglePayloadSection()">
                        <label for="method-post">POST</label>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="section">
            <h3>Predefined URLs</h3>
            {% for category, urls in predefined_urls.items() %}
            <h4>{{ category }}</h4>
            <div class="url-list">
                {% for url in urls %}
                <div class="url-item" onclick="setUrl('{{ url }}')">
                    {{ url }}
                </div>
                {% endfor %}
            </div>
            {% endfor %}
        </div>
        
        <div class="section">
            <h3>API Request</h3>
            
            <div class="form-group">
                <label>API URL:</label>
                <input type="text" id="api-url" value="https://graph.microsoft.com/v1.0/users" placeholder="Enter Graph API URL">
            </div>
            
            <div class="form-group payload-section" id="payload-section">
                <label>Request Payload (JSON):</label>
                <textarea id="request-payload" rows="8" placeholder='{"key": "value"}'></textarea>
            </div>
            
            <button onclick="callApi()">Send Request</button>
        </div>
        
        <div id="status" class="status" style="display: none;"></div>
        
        <div class="section">
            <h3>Response</h3>
            <div class="tabs">
                <button class="tab-btn active" onclick="switchTab('formatted')">Formatted</button>
                <button class="tab-btn" onclick="switchTab('raw')">Raw</button>
            </div>
            <div id="response-formatted" class="response-area">Response will appear here...</div>
            <div id="response-raw" class="response-area" style="display: none;">Raw response will appear here...</div>
        </div>
        
        <div class="section">
            <h3>Recent URLs</h3>
            <div id="history-container" class="history-list">
                <div style="padding: 20px; text-align: center; color: #666;">No history yet</div>
            </div>
        </div>
    </div>

    <script>
        let currentResponse = '';
        let currentRawResponse = '';
        let activeTab = 'formatted';
        
        function setUrl(url) {
            document.getElementById('api-url').value = url;
        }
        
        function togglePayloadSection() {
            const method = document.querySelector('input[name="http-method"]:checked').value;
            const payloadSection = document.getElementById('payload-section');
            
            if (method === 'POST') {
                payloadSection.classList.add('show');
            } else {
                payloadSection.classList.remove('show');
            }
        }
        
        function switchTab(tabName) {
            const tabs = document.querySelectorAll('.tab-btn');
            const formattedDiv = document.getElementById('response-formatted');
            const rawDiv = document.getElementById('response-raw');
            
            tabs.forEach(tab => tab.classList.remove('active'));
            
            if (tabName === 'formatted') {
                document.querySelectorAll('.tab-btn')[0].classList.add('active');
                formattedDiv.style.display = 'block';
                rawDiv.style.display = 'none';
                activeTab = 'formatted';
            } else {
                document.querySelectorAll('.tab-btn')[1].classList.add('active');
                formattedDiv.style.display = 'none';
                rawDiv.style.display = 'block';
                activeTab = 'raw';
            }
        }
        
        function callApi() {
            const url = document.getElementById('api-url').value.trim();
            const authMethod = document.querySelector('input[name="auth-method"]:checked').value;
            const scope = document.querySelector('input[name="scope"]:checked').value;
            const httpMethod = document.querySelector('input[name="http-method"]:checked').value;
            const payload = document.getElementById('request-payload').value.trim();
            const statusDiv = document.getElementById('status');
            const formattedResponseDiv = document.getElementById('response-formatted');
            const rawResponseDiv = document.getElementById('response-raw');
            
            if (!url) {
                setStatus('error', 'Please enter a URL');
                return;
            }

            // Validate JSON payload for POST requests
            let parsedPayload = null;
            if (httpMethod === 'POST' && payload) {
                try {
                    parsedPayload = JSON.parse(payload);
                } catch (e) {
                    setStatus('error', 'Invalid JSON payload');
                    return;
                }
            }

            setStatus('loading', 'Sending request...');
            formattedResponseDiv.textContent = 'Loading...';
            rawResponseDiv.textContent = 'Loading...';

            const requestData = { 
                url: url,
                auth_method: authMethod,
                scope: scope,
                method: httpMethod
            };

            if (httpMethod === 'POST' && parsedPayload) {
                requestData.payload = parsedPayload;
            }

            fetch('/call_graph', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(requestData)
            })
            .then(res => res.json())
            .then(data => {
                if (data.error) {
                    setStatus('error', `Error: ${data.error}`);
                    const errorResponse = JSON.stringify(data, null, 2);
                    formattedResponseDiv.textContent = errorResponse;
                    rawResponseDiv.textContent = errorResponse;
                    currentResponse = errorResponse;
                    currentRawResponse = errorResponse;
                } else {
                    setStatus('success', `Success (Status: ${data.status_code})`);
                    
                    currentRawResponse = data.raw_response || JSON.stringify(data.response, null, 2);
                    currentResponse = data.formatted_response || JSON.stringify(data.response, null, 2);
                    
                    formattedResponseDiv.textContent = currentResponse;
                    rawResponseDiv.textContent = currentRawResponse;
                    
                    addToHistory(url);
                }
            })
            .catch(err => {
                setStatus('error', 'Request failed');
                const errorMsg = `Error: ${err.message}`;
                formattedResponseDiv.textContent = errorMsg;
                rawResponseDiv.textContent = errorMsg;
                currentResponse = errorMsg;
                currentRawResponse = errorMsg;
            });
        }
        
        function setStatus(type, message) {
            const statusDiv = document.getElementById('status');
            statusDiv.style.display = 'block';
            statusDiv.className = `status status-${type}`;
            statusDiv.textContent = message;
        }
        
        function addToHistory(url) {
            let history = JSON.parse(localStorage.getItem('graphApiHistory') || '[]');
            history = history.filter(item => item !== url);
            history.unshift(url);
            history = history.slice(0, 20);
            localStorage.setItem('graphApiHistory', JSON.stringify(history));
            renderHistory();
        }
        
        function renderHistory() {
            const container = document.getElementById('history-container');
            const history = JSON.parse(localStorage.getItem('graphApiHistory') || '[]');
            
            if (history.length === 0) {
                container.innerHTML = '<div style="padding: 20px; text-align: center; color: #666;">No history yet</div>';
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
        
        document.addEventListener('DOMContentLoaded', function() {
            renderHistory();
        });
        
        // Allow Enter key to submit
        document.getElementById('api-url').addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                callApi();
            }
        });
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

@app.route('/call_graph', methods=['POST'])
def handle_call_graph():
    try:
        data = request.json
        api_url = data.get('url')
        auth_method = data.get('auth_method', 'certificate')
        scope_type = data.get('scope', 'graph')
        http_method = data.get('method', 'GET')
        payload = data.get('payload')
        
        # Select the appropriate scope
        scope = CONFIG['scopes'][scope_type]
        
        # Get access token based on authentication method
        if auth_method == 'certificate':
            access_token = get_token_with_certificate(scope)
        else:
            access_token = get_token_with_secret(scope)
        
        if not access_token:
            return jsonify({"error": "Failed to obtain access token"}), 500
        
        # Make the Graph API request
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        if http_method == 'POST':
            graph_response = requests.post(api_url, headers=headers, json=payload)
        else:
            graph_response = requests.get(api_url, headers=headers)
        
        # Get content type
        content_type = graph_response.headers.get('content-type', '')
        raw_response = graph_response.text
        
        try:
            # Try to parse as JSON first
            response_data = graph_response.json()
            formatted_response = json.dumps(response_data, indent=2, ensure_ascii=False)
        except ValueError:
            # If not JSON, check if it's XML
            if is_xml_content(content_type, raw_response):
                response_data = {"content_type": "XML", "raw": raw_response}
                formatted_response = format_xml_response(raw_response)
            else:
                # Plain text or other format
                response_data = {"content_type": content_type, "raw": raw_response}
                formatted_response = raw_response

        return jsonify({
            "status_code": graph_response.status_code,
            "response": response_data,
            "formatted_response": formatted_response,
            "raw_response": raw_response,
            "content_type": content_type,
            "auth_method": auth_method,
            "scope": scope_type,
            "method": http_method
        })

    except Exception as e:
        logger.exception("Error occurred")
        return jsonify({"error": str(e)}), 500

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

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)