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
import xml.dom.minidom
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import deque

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
    <title>Microsoft 365 Admin Tools</title>
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
            background-color: #f5f5f5;
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
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
        
        .tab-container {
            display: flex;
            background: #f8f9fa;
            border-bottom: 1px solid #dee2e6;
        }
        
        .tab-button {
            padding: 15px 25px;
            background: transparent;
            border: none;
            cursor: pointer;
            font-size: 16px;
            font-weight: 500;
            color: #495057;
            transition: all 0.3s ease;
            position: relative;
        }
        
        .tab-button:hover {
            color: #28a745;
        }
        
        .tab-button.active {
            color: #28a745;
            font-weight: 600;
        }
        
        .tab-button.active::after {
            content: '';
            position: absolute;
            bottom: 0;
            left: 0;
            right: 0;
            height: 3px;
            background: #28a745;
        }
        
        .tab-content {
            display: none;
            padding: 30px;
        }
        
        .tab-content.active {
            display: block;
        }
        
        /* OneDrive Repair Tool Styles */
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
        
        /* SharePoint User Remover Styles */
        .warning {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            padding: 15px;
            border-radius: 4px;
            margin-bottom: 20px;
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
        
        .results-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 20px;
        }
        
        .results-table th, .results-table td {
            padding: 10px;
            border: 1px solid #ccc;
            text-align: left;
        }
        
        .results-table th {
            background: #f9f9f9;
        }
        
        .status-success {
            color: green;
            font-weight: bold;
        }
        
        .status-error {
            color: red;
            font-weight: bold;
        }
        
        .status-pending {
            color: orange;
            font-weight: bold;
        }
        
        .progress-bar {
            width: 100%;
            height: 20px;
            background: #f0f0f0;
            border-radius: 10px;
            margin: 10px 0;
        }
        
        .progress-fill {
            height: 100%;
            background: #007cba;
            border-radius: 10px;
            transition: width 0.3s ease;
        }
        
        /* Graph Explorer Styles */
        .section {
            margin-bottom: 20px;
            padding: 15px;
            border: 1px solid #ddd;
            border-radius: 3px;
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
        
        /* Responsive styles */
        @media (max-width: 768px) {
            .container {
                margin: 10px;
                border-radius: 15px;
            }
            
            .tab-content {
                padding: 20px;
            }
            
            .header h1 {
                font-size: 2rem;
            }
            
            .btn-row {
                grid-template-columns: 1fr;
            }
            
            .tab-container {
                overflow-x: auto;
                white-space: nowrap;
            }
            
            .tab-button {
                padding: 12px 20px;
                font-size: 14px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i class="fas fa-tools"></i> Microsoft 365 Admin Tools</h1>
            <p>Tools for managing OneDrive, SharePoint, and Microsoft Graph</p>
        </div>
        
        <div class="tab-container">
            <button class="tab-button active" onclick="showTab('onedrive-repair')">OneDrive Repair</button>
            <button class="tab-button" onclick="showTab('user-remover')">User Remover</button>
            <button class="tab-button" onclick="showTab('graph-explorer')">Graph Explorer</button>
        </div>
        
        <!-- OneDrive Repair Tab -->
        <div id="onedrive-repair-tab" class="tab-content active">
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
        
        <!-- User Remover Tab -->
        <div id="user-remover-tab" class="tab-content">
            <div class="warning">
                <strong>Warning:</strong> This will remove the specified user from SharePoint sites or OneDrives. This action cannot be undone!
            </div>
            
            <div class="form-group">
                <label for="single-removal-type">Removal Type:</label>
                <select id="single-removal-type" class="form-control" onchange="toggleSingleRemovalType()">
                    <option value="site">Remove from SharePoint Site</option>
                    <option value="onedrive">Remove from User's OneDrive</option>
                </select>
            </div>
            
            <div class="form-group">
                <label for="single-user-upn">User to Remove (Email/UPN):</label>
                <input type="email" id="single-user-upn" class="form-control" placeholder="user@geekbyteonline.onmicrosoft.com" required>
            </div>
            
            <div id="single-site-url-group" class="form-group">
                <label for="single-site-url">SharePoint Site URL:</label>
                <input type="url" id="single-site-url" class="form-control" placeholder="https://geekbyteonline.sharepoint.com/sites/sitename">
            </div>
            
            <div id="single-onedrive-owner-group" class="form-group" style="display: none;">
                <label for="single-onedrive-owner-upn">OneDrive Owner (Email/UPN):</label>
                <input type="email" id="single-onedrive-owner-upn" class="form-control" placeholder="owner@geekbyteonline.onmicrosoft.com">
            </div>
            
            <div class="btn-row">
                <button id="single-find-user-btn" class="btn btn-info" onclick="findSingleUser()">
                    <i class="fas fa-search"></i> Find User
                </button>
                <button id="single-remove-user-btn" class="btn btn-success" onclick="removeSingleUser()" disabled>
                    <i class="fas fa-user-minus"></i> Remove User
                </button>
            </div>
            
            <div id="single-status" class="status-section status-ready">
                <i class="fas fa-info-circle"></i> Enter user details to start
            </div>
            
            <div id="single-results-section" class="results-section" style="display: none;">
                <h3>User Information</h3>
                <div id="single-user-info"></div>
                <div id="single-confirmation-section" style="display: none;"></div>
            </div>
            
            <div class="warning">
                <strong>Bulk Removal:</strong> Remove user from multiple OneDrives
            </div>
            
            <div class="form-group">
                <label for="bulk-user-upn">User to Remove from Multiple OneDrives (Email/UPN):</label>
                <input type="email" id="bulk-user-upn" class="form-control" placeholder="user@geekbyteonline.onmicrosoft.com" required>
            </div>
            
            <div class="form-group">
                <label for="onedrive-owners">OneDrive Owners (one email per line):</label>
                <textarea id="onedrive-owners" class="form-control" rows="5" placeholder="owner1@geekbyteonline.onmicrosoft.com&#10;owner2@geekbyteonline.onmicrosoft.com&#10;owner3@geekbyteonline.onmicrosoft.com" required></textarea>
                <div class="help-text">Enter one email address per line for each OneDrive owner</div>
            </div>
            
            <div class="btn-row">
                <button id="bulk-find-btn" class="btn btn-info" onclick="findBulkUsers()">
                    <i class="fas fa-search"></i> Find User on All OneDrives
                </button>
                <button id="bulk-remove-btn" class="btn btn-success" onclick="removeBulkUsers()" disabled>
                    <i class="fas fa-user-minus"></i> Remove from All Found OneDrives
                </button>
            </div>
            
            <div id="bulk-status" class="status-section status-ready">
                <i class="fas fa-info-circle"></i> Enter user details to start bulk removal
            </div>
            
            <div id="bulk-progress" class="progress-container" style="display: none;">
                <h4 style="margin-bottom: 15px; color: #28a745;">
                    <i class="fas fa-spinner"></i> Progress
                </h4>
                <div class="progress-bar">
                    <div id="progress-fill" class="progress-fill" style="width: 0%;"></div>
                </div>
                <p id="progress-text">0 / 0 completed</p>
            </div>
            
            <div id="bulk-results-section" class="results-section" style="display: none;">
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
        
        <!-- Graph Explorer Tab -->
        <div id="graph-explorer-tab" class="tab-content">
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
                    <input type="text" id="api-url" class="form-control" value="https://graph.microsoft.com/v1.0/users" placeholder="Enter Graph API URL">
                </div>
                
                <div class="form-group payload-section" id="payload-section">
                    <label>Request Payload (JSON):</label>
                    <textarea id="request-payload" class="form-control" rows="8" placeholder='{"key": "value"}'></textarea>
                </div>
                
                <button class="btn btn-success" onclick="callApi()">
                    <i class="fas fa-paper-plane"></i> Send Request
                </button>
            </div>
            
            <div id="graph-status" class="status-section status-ready">
                <i class="fas fa-info-circle"></i> Ready to send request
            </div>
            
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
    </div>

    <script>
        // Tab management
        function showTab(tabName) {
            // Hide all tabs
            document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
            document.querySelectorAll('.tab-button').forEach(btn => btn.classList.remove('active'));
            
            // Show selected tab
            document.getElementById(tabName + '-tab').classList.add('active');
            event.target.classList.add('active');
            
            // Reset forms when switching tabs
            if (tabName === 'onedrive-repair') {
                resetOneDriveForm();
            } else if (tabName === 'user-remover') {
                resetUserRemoverForm();
            } else if (tabName === 'graph-explorer') {
                resetGraphExplorerForm();
            }
        }
        
        // OneDrive Repair Tool Functions
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
            
            // Create EventSource for server-sent events
            const eventSource = new EventSource(`/repair_onedrive_stream?upn=${encodeURIComponent(currentUserUpn)}&onedrive_url=${encodeURIComponent(currentOneDriveUrl)}`);
            
            eventSource.onmessage = function(e) {
                const data = JSON.parse(e.data);
                
                if (data.step !== undefined && data.status) {
                    updateProgressStep(data.step, data.status, data.message || '');
                }
                
                if (data.error) {
                    setStatus('error', `Error: ${data.error}`);
                    eventSource.close();
                    repairBtn.disabled = false;
                    getBtn.disabled = false;
                    repairBtn.innerHTML = '<i class="fas fa-tools"></i> Repair OneDrive';
                }
                
                if (data.success) {
                    setStatus('success', 'OneDrive repair completed successfully');
                    eventSource.close();
                    repairBtn.disabled = false;
                    getBtn.disabled = false;
                    repairBtn.innerHTML = '<i class="fas fa-tools"></i> Repair OneDrive';
                }
            };
            
            eventSource.onerror = function() {
                setStatus('error', 'Connection to server failed');
                eventSource.close();
                repairBtn.disabled = false;
                getBtn.disabled = false;
                repairBtn.innerHTML = '<i class="fas fa-tools"></i> Repair OneDrive';
            };
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
            const steps = document.querySelectorAll('.progress-step');
            steps.forEach((step, index) => {
                updateProgressStep(index, 'success');
            });
        }
        
        function resetOneDriveForm() {
            currentOneDriveUrl = null;
            currentUserUpn = null;
            document.getElementById('user-upn').value = '';
            document.getElementById('repair-onedrive-btn').disabled = true;
            document.getElementById('results-section').style.display = 'none';
            document.getElementById('progress-container').style.display = 'none';
            setStatus('ready', 'Enter user UPN to start OneDrive repair');
        }
        
        // User Remover Functions
        let singleCurrentUserId = null;
        let singleCurrentSiteUrl = null;
        let bulkResults = [];
        
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
        
        function resetUserRemoverForm() {
            resetSingleForm();
            resetBulkForm();
        }
        
        function setSingleStatus(message) {
            const statusDiv = document.getElementById('single-status');
            statusDiv.innerHTML = `<i class="fas fa-info-circle"></i> <p>${message}</p>`;
        }
        
        function setBulkStatus(message) {
            const statusDiv = document.getElementById('bulk-status');
            statusDiv.innerHTML = `<i class="fas fa-info-circle"></i> <p>${message}</p>`;
        }
        
        function updateProgress(completed, total) {
            const progressFill = document.getElementById('progress-fill');
            const progressText = document.getElementById('progress-text');
            const percentage = total > 0 ? (completed / total) * 100 : 0;
            
            progressFill.style.width = percentage + '%';
            progressText.textContent = `${completed} / ${total} completed`;
        }
        
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
            findBtn.innerHTML = '<div class="spinner"></div> Finding User...';
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
                findBtn.innerHTML = '<i class="fas fa-search"></i> Find User';
                
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
                findBtn.innerHTML = '<i class="fas fa-search"></i> Find User';
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
            removeBtn.innerHTML = '<div class="spinner"></div> Removing User...';
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
                removeBtn.innerHTML = '<i class="fas fa-user-minus"></i> Remove User';
                
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
                removeBtn.innerHTML = '<i class="fas fa-user-minus"></i> Remove User';
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
            findBtn.innerHTML = '<div class="spinner"></div> Finding Users...';
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
                findBtn.innerHTML = '<i class="fas fa-search"></i> Find User on All OneDrives';
                
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
                findBtn.innerHTML = '<i class="fas fa-search"></i> Find User on All OneDrives';
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
            removeBtn.innerHTML = '<div class="spinner"></div> Removing Users...';
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
                removeBtn.innerHTML = '<i class="fas fa-user-minus"></i> Remove from All Found OneDrives';
                
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
                removeBtn.innerHTML = '<i class="fas fa-user-minus"></i> Remove from All Found OneDrives';
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
        
        // Graph Explorer Functions
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
        
        function setGraphStatus(type, message) {
            const statusDiv = document.getElementById('graph-status');
            statusDiv.style.display = 'block';
            statusDiv.className = `status-section status-${type}`;
            
            let icon = 'fas fa-info-circle';
            if (type === 'loading') icon = 'fas fa-spinner fa-spin';
            else if (type === 'success') icon = 'fas fa-check-circle';
            else if (type === 'error') icon = 'fas fa-exclamation-circle';
            
            statusDiv.innerHTML = `<i class="${icon}"></i> ${message}`;
        }
        
        function callApi() {
            const url = document.getElementById('api-url').value.trim();
            const authMethod = document.querySelector('input[name="auth-method"]:checked').value;
            const scope = document.querySelector('input[name="scope"]:checked').value;
            const httpMethod = document.querySelector('input[name="http-method"]:checked').value;
            const payload = document.getElementById('request-payload').value.trim();
            const formattedResponseDiv = document.getElementById('response-formatted');
            const rawResponseDiv = document.getElementById('response-raw');
            
            if (!url) {
                setGraphStatus('error', 'Please enter a URL');
                return;
            }

            // Validate JSON payload for POST requests
            let parsedPayload = null;
            if (httpMethod === 'POST' && payload) {
                try {
                    parsedPayload = JSON.parse(payload);
                } catch (e) {
                    setGraphStatus('error', 'Invalid JSON payload');
                    return;
                }
            }

            setGraphStatus('loading', 'Sending request...');
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
                    setGraphStatus('error', `Error: ${data.error}`);
                    const errorResponse = JSON.stringify(data, null, 2);
                    formattedResponseDiv.textContent = errorResponse;
                    rawResponseDiv.textContent = errorResponse;
                    currentResponse = errorResponse;
                    currentRawResponse = errorResponse;
                } else {
                    setGraphStatus('success', `Success (Status: ${data.status_code})`);
                    
                    currentRawResponse = data.raw_response || JSON.stringify(data.response, null, 2);
                    currentResponse = data.formatted_response || JSON.stringify(data.response, null, 2);
                    
                    formattedResponseDiv.textContent = currentResponse;
                    rawResponseDiv.textContent = currentRawResponse;
                    
                    addToHistory(url);
                }
            })
            .catch(err => {
                setGraphStatus('error', 'Request failed');
                const errorMsg = `Error: ${err.message}`;
                formattedResponseDiv.textContent = errorMsg;
                rawResponseDiv.textContent = errorMsg;
                currentResponse = errorMsg;
                currentRawResponse = errorMsg;
            });
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
        
        function resetGraphExplorerForm() {
            document.getElementById('api-url').value = 'https://graph.microsoft.com/v1.0/users';
            document.getElementById('request-payload').value = '';
            document.getElementById('payload-section').classList.remove('show');
            document.getElementById('response-formatted').textContent = 'Response will appear here...';
            document.getElementById('response-raw').textContent = 'Raw response will appear here...';
            setGraphStatus('ready', 'Ready to send request');
        }
        
        // Initialize the application
        document.addEventListener('DOMContentLoaded', function() {
            renderHistory();
            
            // Allow Enter key to submit
            document.getElementById('user-upn').addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    getOneDriveInfo();
                }
            });
            
            document.getElementById('api-url').addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    callApi();
                }
            });
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
    return render_template_string(HTML_TEMPLATE, predefined_urls=PREDEFINED_URLS)

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
            return jsonify({"error": "User UPN and OneDrive URL are required", "failed_step": 0}), 400
        
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
        result = {"steps": []}
        
        try:
            # Step 1: Add repair account as user
            result['steps'].append({"step": 0, "status": "running", "message": "Adding repair account as user"})
            repair_user_id = ensure_user(site_url, sp_token, request_digest, CONFIG['repair_account'])
            if not repair_user_id:
                raise Exception("Failed to add repair account as user")
            result['steps'].append({"step": 0, "status": "success"})
            
            # Step 2: Make repair account site admin
            result['steps'].append({"step": 1, "status": "running", "message": "Making repair account site admin"})
            if not set_site_admin(site_url, sp_token, request_digest, repair_user_id, True):
                raise Exception("Failed to make repair account site admin")
            result['steps'].append({"step": 1, "status": "success"})
            
            # Step 3: Get original user ID and remove admin rights
            result['steps'].append({"step": 2, "status": "running", "message": "Getting original user and removing admin rights"})
            original_user_id = get_user_id_from_site(site_url, sp_token, upn)
            if original_user_id:
                if not set_site_admin(site_url, sp_token, request_digest, original_user_id, False):
                    logger.warning("Failed to remove original user admin rights, continuing...")
                    result['steps'].append({"step": 2, "status": "warning", "message": "Failed to remove original user admin rights, continuing..."})
            result['steps'].append({"step": 2, "status": "success"})
            
            # Step 4: Remove original user from site
            result['steps'].append({"step": 3, "status": "running", "message": "Removing original user from site"})
            if original_user_id:
                if not remove_user_by_id(site_url, sp_token, request_digest, original_user_id):
                    logger.warning("Failed to remove original user, continuing...")
                    result['steps'].append({"step": 3, "status": "warning", "message": "Failed to remove original user, continuing..."})
            result['steps'].append({"step": 3, "status": "success"})
            
            # Wait a moment for changes to propagate
            time.sleep(2)
            
            # Step 5: Re-add original user to site
            result['steps'].append({"step": 4, "status": "running", "message": "Re-adding original user to site"})
            new_user_id = ensure_user(site_url, sp_token, request_digest, upn)
            if not new_user_id:
                raise Exception("Failed to re-add original user to site")
            result['steps'].append({"step": 4, "status": "success"})
            
            # Step 6: Make original user site admin
            result['steps'].append({"step": 5, "status": "running", "message": "Making original user site admin"})
            if not set_site_admin(site_url, sp_token, request_digest, new_user_id, True):
                raise Exception("Failed to make original user site admin")
            result['steps'].append({"step": 5, "status": "success"})
            
            # Step 7: Remove repair account admin rights
            result['steps'].append({"step": 6, "status": "running", "message": "Removing repair account admin rights"})
            if not set_site_admin(site_url, sp_token, request_digest, repair_user_id, False):
                logger.warning("Failed to remove repair account admin rights, continuing...")
                result['steps'].append({"step": 6, "status": "warning", "message": "Failed to remove repair account admin rights, continuing..."})
            result['steps'].append({"step": 6, "status": "success"})
            
            # Step 8: Remove repair account from site
            result['steps'].append({"step": 7, "status": "running", "message": "Removing repair account from site"})
            if not remove_user_by_id(site_url, sp_token, request_digest, repair_user_id):
                logger.warning("Failed to remove repair account from site")
                result['steps'].append({"step": 7, "status": "warning", "message": "Failed to remove repair account from site"})
            result['steps'].append({"step": 7, "status": "success"})
            
            logger.info(f"OneDrive repair completed successfully for {upn}")
            result["success"] = True
            result["message"] = "OneDrive repair completed successfully"
            return jsonify(result)
            
        except Exception as e:
            logger.exception("Error during OneDrive repair")
            result["error"] = str(e)
            return jsonify(result), 500
            
    except Exception as e:
        logger.exception("Error during OneDrive repair initialization")
        return jsonify({"error": str(e), "failed_step": 0}), 500

        

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

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5500, debug=True)