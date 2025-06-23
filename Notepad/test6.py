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
import re

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
            max-width: 1400px;
            margin: 0 auto;
            background: rgba(255, 255, 255, 0.95);
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            backdrop-filter: blur(10px);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%);
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
            display: flex;
            min-height: 600px;
        }
        
        .sidebar {
            background: #f8f9fa;
            padding: 30px;
            border-right: 1px solid #e9ecef;
            width: 350px;
            overflow-y: auto;
        }
        
        .content-area {
            padding: 30px;
            flex: 1;
            display: flex;
            flex-direction: column;
        }
        
        .config-section {
            background: white;
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        }
        
        .config-section h3 {
            color: #0078d4;
            margin-bottom: 15px;
            font-size: 1.1rem;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .form-group {
            margin-bottom: 15px;
        }
        
        .form-group label {
            display: block;
            margin-bottom: 5px;
            font-weight: 500;
            color: #495057;
        }
        
        .form-control {
            width: 100%;
            padding: 12px;
            border: 2px solid #e9ecef;
            border-radius: 8px;
            font-size: 14px;
            transition: all 0.3s ease;
        }
        
        .form-control:focus {
            outline: none;
            border-color: #0078d4;
            box-shadow: 0 0 0 3px rgba(0, 120, 212, 0.1);
        }
        
        .radio-group {
            display: flex;
            gap: 15px;
            flex-wrap: wrap;
        }
        
        .radio-item {
            display: flex;
            align-items: center;
            gap: 8px;
            cursor: pointer;
            padding: 8px 12px;
            border-radius: 6px;
            transition: background-color 0.2s;
        }
        
        .radio-item:hover {
            background-color: #f8f9fa;
        }
        
        .radio-item input[type="radio"] {
            margin: 0;
        }
        
        .btn {
            padding: 12px 24px;
            border: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.3s ease;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            text-decoration: none;
        }
        
        .btn-primary {
            background: linear-gradient(135deg, #0078d4 0%, #106ebe 100%);
            color: white;
        }
        
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 16px rgba(0, 120, 212, 0.3);
        }
        
        .btn-block {
            width: 100%;
            justify-content: center;
        }
        
        .url-input-section {
            background: white;
            border-radius: 12px;
            padding: 25px;
            margin-bottom: 25px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        }
        
        .predefined-urls {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 15px;
            margin-bottom: 20px;
        }
        
        .url-category {
            background: #f8f9fa;
            border-radius: 8px;
            padding: 15px;
        }
        
        .url-category h4 {
            color: #0078d4;
            margin-bottom: 10px;
            font-size: 1rem;
        }
        
        .url-item {
            padding: 8px 12px;
            background: white;
            border: 1px solid #e9ecef;
            border-radius: 6px;
            margin-bottom: 5px;
            cursor: pointer;
            transition: all 0.2s;
            font-size: 12px;
            word-break: break-all;
        }
        
        .url-item:hover {
            background: #e3f2fd;
            border-color: #0078d4;
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
        
        .response-section {
            background: white;
            border-radius: 12px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
            flex: 1;
            display: flex;
            flex-direction: column;
        }
        
        .response-header {
            background: #f8f9fa;
            padding: 15px 20px;
            border-bottom: 1px solid #e9ecef;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        
        .response-tabs {
            display: flex;
            gap: 10px;
        }
        
        .tab-btn {
            padding: 8px 16px;
            border: none;
            background: #e9ecef;
            border-radius: 6px;
            cursor: pointer;
            font-size: 13px;
            transition: all 0.2s;
        }
        
        .tab-btn.active {
            background: #0078d4;
            color: white;
        }
        
        .response-body {
            padding: 20px;
            overflow-y: auto;
            flex: 1;
            min-height: 400px;
        }
        
        .response-content {
            background: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 6px;
            padding: 15px;
            font-family: 'Courier New', monospace;
            font-size: 13px;
            line-height: 1.4;
            white-space: pre-wrap;
            word-break: break-word;
            min-height: 300px;
        }
        
        /* Syntax highlighting styles */
        .json-key {
            color: #d63384;
            font-weight: bold;
        }
        .json-string {
            color: #20a8d8;
        }
        .json-number {
            color: #4dbd74;
        }
        .json-boolean {
            color: #f86c6b;
        }
        .json-null {
            color: #f86c6b;
        }
        .json-punctuation {
            color: #5c6873;
        }
        
        /* XML syntax highlighting */
        .xml-tag {
            color: #2f6f9f;
        }
        .xml-attribute {
            color: #4f9fcf;
        }
        .xml-attribute-value {
            color: #d44950;
        }
        .xml-text {
            color: #333;
        }
        .xml-comment {
            color: #999;
            font-style: italic;
        }
        .xml-punctuation {
            color: #5c6873;
        }
        
        .history-section {
            background: white;
            border-radius: 12px;
            padding: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        }
        
        .history-item {
            padding: 10px;
            background: #f8f9fa;
            border: 1px solid #e9ecef;
            border-radius: 6px;
            margin-bottom: 8px;
            cursor: pointer;
            transition: all 0.2s;
            font-size: 13px;
            word-break: break-all;
        }
        
        .history-item:hover {
            background: #e3f2fd;
            border-color: #0078d4;
        }
        
        .copy-btn {
            background: none;
            border: none;
            color: #6c757d;
            cursor: pointer;
            padding: 5px 10px;
            border-radius: 4px;
            transition: all 0.2s;
            font-size: 12px;
        }
        
        .copy-btn:hover {
            color: #0078d4;
            background: #f8f9fa;
        }
        
        @media (max-width: 768px) {
            .main-content {
                flex-direction: column;
            }
            
            .sidebar {
                width: 100%;
            }
            
            .predefined-urls {
                grid-template-columns: 1fr;
            }
            
            .header h1 {
                font-size: 2rem;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i class="fas fa-chart-network"></i> Microsoft Graph API Explorer</h1>
            <p>Advanced tool for exploring Microsoft Graph and SharePoint APIs</p>
        </div>
        
        <div class="main-content">
            <div class="sidebar">
                <div class="config-section">
                    <h3><i class="fas fa-cog"></i> Authentication</h3>
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
                </div>
                
                <div class="config-section">
                    <h3><i class="fas fa-key"></i> Scope</h3>
                    <div class="form-group">
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
                </div>
                
                <div class="history-section">
                    <h3><i class="fas fa-history"></i> Recent URLs</h3>
                    <div id="history-container">
                        <div style="color: #6c757d; text-align: center; padding: 20px;">
                            No history yet
                        </div>
                    </div>
                </div>
            </div>
            
            <div class="content-area">
                <div class="url-input-section">
                    <h3 style="margin-bottom: 20px; color: #0078d4;"><i class="fas fa-link"></i> API Endpoint</h3>
                    
                    <div class="predefined-urls">
                        {% for category, urls in predefined_urls.items() %}
                        <div class="url-category">
                            <h4>{{ category }}</h4>
                            {% for url in urls %}
                            <div class="url-item" onclick="setUrl('{{ url }}')">
                                {{ url }}
                            </div>
                            {% endfor %}
                        </div>
                        {% endfor %}
                    </div>
                    
                    <div class="form-group">
                        <label>Custom API URL:</label>
                        <input type="text" id="api-url" class="form-control" 
                               value="https://graph.microsoft.com/v1.0/users" 
                               placeholder="Enter Graph API URL">
                    </div>
                    
                    <button class="btn btn-primary btn-block" onclick="callApi()">
                        <i class="fas fa-paper-plane"></i> Send Request
                    </button>
                </div>
                
                <div id="status" class="status-section status-ready">
                    <i class="fas fa-info-circle"></i> Ready to send request
                </div>
                
                <div class="response-section">
                    <div class="response-header">
                        <div class="response-tabs">
                            <button class="tab-btn active" onclick="switchTab('formatted')">Formatted</button>
                            <button class="tab-btn" onclick="switchTab('raw')">Raw</button>
                        </div>
                        <button class="copy-btn" onclick="copyResponse()" title="Copy Response">
                            <i class="fas fa-copy"></i> Copy
                        </button>
                    </div>
                    <div class="response-body">
                        <div id="response-formatted" class="response-content">Response will appear here...</div>
                        <div id="response-raw" class="response-content" style="display: none;">Raw response will appear here...</div>
                    </div>
                </div>
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
        
        function switchTab(tabName) {
            const tabs = document.querySelectorAll('.tab-btn');
            const formattedDiv = document.getElementById('response-formatted');
            const rawDiv = document.getElementById('response-raw');
            
            tabs.forEach(tab => tab.classList.remove('active'));
            
            if (tabName === 'formatted') {
                document.querySelector('.tab-btn').classList.add('active');
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
            const statusDiv = document.getElementById('status');
            const formattedResponseDiv = document.getElementById('response-formatted');
            const rawResponseDiv = document.getElementById('response-raw');
            
            if (!url) {
                setStatus('error', 'Please enter a URL');
                return;
            }

            setStatus('loading', 'Sending request...');
            formattedResponseDiv.textContent = '';
            rawResponseDiv.textContent = '';

            fetch('/call_graph', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    url: url,
                    auth_method: authMethod,
                    scope: scope
                })
            })
            .then(res => res.json())
            .then(data => {
                if (data.error) {
                    setStatus('error', `Error: ${data.error}`);
                    const errorResponse = JSON.stringify(data, null, 2);
                    formattedResponseDiv.innerHTML = syntaxHighlight(errorResponse);
                    rawResponseDiv.textContent = errorResponse;
                    currentResponse = errorResponse;
                    currentRawResponse = errorResponse;
                } else {
                    setStatus('success', `Success (Status: ${data.status_code})`);
                    
                    // Format the response
                    currentRawResponse = data.raw_response || JSON.stringify(data.response, null, 2);
                    
                    if (data.content_type && data.content_type.includes('xml')) {
                        currentResponse = syntaxHighlightXML(data.raw_response);
                    } else {
                        try {
                            const jsonResponse = JSON.parse(currentRawResponse);
                            currentResponse = syntaxHighlight(JSON.stringify(jsonResponse, null, 2));
                        } catch {
                            currentResponse = data.raw_response;
                        }
                    }
                    
                    formattedResponseDiv.innerHTML = currentResponse;
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
        
        function syntaxHighlight(json) {
            if (typeof json != 'string') {
                json = JSON.stringify(json, undefined, 2);
            }
            
            json = json.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
            
            return json.replace(
                /("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?)/g, 
                function (match) {
                    let cls = 'json-number';
                    if (/^"/.test(match)) {
                        if (/:$/.test(match)) {
                            cls = 'json-key';
                        } else {
                            cls = 'json-string';
                        }
                    } else if (/true|false/.test(match)) {
                        cls = 'json-boolean';
                    } else if (/null/.test(match)) {
                        cls = 'json-null';
                    }
                    return '<span class="' + cls + '">' + match + '</span>';
                }
            ).replace(/{|}|\[|\]|,|:/g, function(match) {
                return '<span class="json-punctuation">' + match + '</span>';
            });
        }
        
        function syntaxHighlightXML(xml) {
            if (typeof xml != 'string') {
                xml = xml.toString();
            }
            
            // Escape HTML
            xml = xml.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
            
            // Highlight tags
            xml = xml.replace(/&lt;([\/\!]?)([a-zA-Z][a-zA-Z0-9:-]*)([^&]*?)&gt;/g, 
                function(match, slash, tagName, rest) {
                    let result = '<span class="xml-punctuation">&lt;</span>';
                    if (slash) result += '<span class="xml-punctuation">' + slash + '</span>';
                    
                    result += '<span class="xml-tag">' + tagName + '</span>';
                    
                    // Highlight attributes
                    rest = rest.replace(/([a-zA-Z-]+)=("[^"]*"|'[^']*')/g, 
                        '<span class="xml-attribute">$1</span>=<span class="xml-attribute-value">$2</span>');
                    
                    result += rest;
                    result += '<span class="xml-punctuation">&gt;</span>';
                    return result;
                });
            
            // Highlight comments
            xml = xml.replace(/&lt;!--([\s\S]*?)--&gt;/g, 
                '<span class="xml-punctuation">&lt;!--</span><span class="xml-comment">$1</span><span class="xml-punctuation">--&gt;</span>');
            
            // Highlight text content
            xml = xml.replace(/(&gt;)([^&]*?)(&lt;)/g, 
                function(match, gt, text, lt) {
                    return gt + '<span class="xml-text">' + text + '</span>' + lt;
                });
            
            return xml;
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
                container.innerHTML = '<div style="color: #6c757d; text-align: center; padding: 20px;">No history yet</div>';
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
            const responseText = activeTab === 'formatted' ? 
                document.getElementById('response-formatted').textContent : 
                currentRawResponse;
                
            navigator.clipboard.writeText(responseText).then(() => {
                setStatus('success', 'Response copied to clipboard!');
                setTimeout(() => setStatus('ready', 'Ready to send request'), 2000);
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
            "scope": scope_type
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
    app.run(host='0.0.0.0', port=5002, debug=True)