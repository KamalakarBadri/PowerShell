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
    "new_id_site_url": "https://geekbyteonline.sharepoint.com/sites/2DayRetention",
    "scopes": {
        "graph": "https://graph.microsoft.com/.default",
        "sharepoint": "https://geekbyteonline.sharepoint.com/.default"
    }
}

RECENT_URLS = deque(maxlen=20)

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
        
        .btn-warning {
            background: linear-gradient(135deg, #ffc107 0%, #e0a800 100%);
            color: #212529;
        }
        
        .btn-warning:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 16px rgba(255, 193, 7, 0.3);
        }
        
        .btn-danger {
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
            color: white;
        }
        
        .btn-danger:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 16px rgba(220, 53, 69, 0.3);
        }
        
        .btn-sm {
            padding: 8px 15px;
            font-size: 14px;
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
        
        .btn-row-three {
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
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
        
        /* NameId comparison styles */
        .nameid-match {
            color: green;
            font-weight: bold;
        }
        
        .nameid-mismatch {
            color: red;
            font-weight: bold;
        }
        
        .nameid-warning {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            padding: 15px;
            border-radius: 8px;
            margin: 15px 0;
        }
        
        .nameid-comparison-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 15px;
            margin-bottom: 15px;
        }
        
        .nameid-box {
            background: #f8f9fa;
            padding: 15px;
            border-radius: 8px;
            border: 2px solid #e9ecef;
        }
        
        .nameid-box.mismatch {
            border-color: #dc3545;
            background: #f8d7da;
        }
        
        .nameid-box.match {
            border-color: #28a745;
            background: #d4edda;
        }
        
        .action-buttons {
            display: flex;
            gap: 10px;
            margin-top: 10px;
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
            
            .btn-row, .btn-row-three {
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
            
            .nameid-comparison-grid {
                grid-template-columns: 1fr;
            }
            
            .action-buttons {
                flex-direction: column;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1><i class="fas fa-tools"></i> Microsoft 365 Admin Tools</h1>
            <p>Tools for managing OneDrive and SharePoint user permissions</p>
        </div>
        
        <div class="tab-container">
            <button class="tab-button active" onclick="showTab('onedrive-repair')">OneDrive Repair</button>
            <button class="tab-button" onclick="showTab('user-remover')">User Remover</button>
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
            
            <div class="info-banner">
                <i class="fas fa-info-circle"></i>
                <strong>NameId Comparison:</strong> The tool compares NameIds between the current site and new ID site. 
                A mismatch indicates PUID issues, but removal works regardless of NameId status.
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
            
            <div class="btn-row-three">
                <button id="single-find-user-btn" class="btn btn-info" onclick="findSingleUser()">
                    <i class="fas fa-search"></i> Find User
                </button>
                <button id="single-remove-user-btn" class="btn btn-success" onclick="removeSingleUser()" disabled>
                    <i class="fas fa-user-minus"></i> Remove User
                </button>
                <button id="single-remove-mismatch-btn" class="btn btn-warning" onclick="removeUserForMismatch()" disabled style="display: none;">
                    <i class="fas fa-exclamation-triangle"></i> Remove for Mismatch
                </button>
            </div>
            
            <div id="single-status" class="status-section status-ready">
                <i class="fas fa-info-circle"></i> Enter user details to start
            </div>
            
            <div id="single-results-section" class="results-section" style="display: none;">
                <h3>User Information</h3>
                <div id="single-user-info"></div>
                <div id="nameid-comparison" style="margin-top: 20px; display: none;">
                    <h4>NameId Comparison</h4>
                    <div id="nameid-comparison-content"></div>
                </div>
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
                <button id="bulk-remove-all-btn" class="btn btn-danger" onclick="removeAllBulkUsers()" disabled>
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
                            <th>NameId Comparison</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody id="bulk-results-body">
                    </tbody>
                </table>
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
            document.getElementById('single-remove-mismatch-btn').disabled = true;
            document.getElementById('single-remove-mismatch-btn').style.display = 'none';
            document.getElementById('single-results-section').style.display = 'none';
            document.getElementById('nameid-comparison').style.display = 'none';
            setSingleStatus('Enter user details to start');
        }
        
        function resetBulkForm() {
            bulkResults = [];
            document.getElementById('bulk-remove-all-btn').disabled = true;
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
            currentUserUpn = upn;
            const removalType = document.getElementById('single-removal-type').value;
            const findBtn = document.getElementById('single-find-user-btn');
            const removeBtn = document.getElementById('single-remove-user-btn');
            const mismatchBtn = document.getElementById('single-remove-mismatch-btn');
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
            mismatchBtn.disabled = true;
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
                    mismatchBtn.disabled = true;
                } else {
                    setSingleStatus(removalType === 'site' ? 'User found on site' : 'User found on OneDrive');
                    displaySingleUserInfo(data);
                    resultsSection.style.display = 'block';
                    removeBtn.disabled = false;
                    singleCurrentUserId = data.user_id;
                    singleCurrentSiteUrl = data.site_url;
                    
                    // Show NameId comparison if available
                    if (data.nameid_comparison) {
                        displayNameIdComparison(data.nameid_comparison);
                    }
                }
            })
            .catch(err => {
                findBtn.disabled = false;
                findBtn.innerHTML = '<i class="fas fa-search"></i> Find User';
                setSingleStatus('Request failed: ' + err.message);
                resultsSection.style.display = 'none';
                removeBtn.disabled = true;
                mismatchBtn.disabled = true;
            });
        }
        
        function removeSingleUser() {
            if (!singleCurrentUserId || !singleCurrentSiteUrl) {
                setSingleStatus('Error: Please find a user first');
                return;
            }
            
            let confirmationMessage = 'Are you sure you want to remove this user? This action cannot be undone.';
            
            // Check if NameId comparison exists and show mismatch warning
            const nameidComparison = document.getElementById('nameid-comparison');
            if (nameidComparison.style.display !== 'none') {
                const nameidContent = document.getElementById('nameid-comparison-content');
                if (nameidContent.innerHTML.includes('nameid-mismatch')) {
                    confirmationMessage = 'WARNING: NameId mismatch detected! This indicates a PUID issue. Are you sure you want to remove this user? The removal will work, but the user may need OneDrive repair.';
                }
            }
            
            if (!confirm(confirmationMessage)) {
                return;
            }
            
            const removeBtn = document.getElementById('single-remove-user-btn');
            const findBtn = document.getElementById('single-find-user-btn');
            const mismatchBtn = document.getElementById('single-remove-mismatch-btn');
            
            removeBtn.disabled = true;
            findBtn.disabled = true;
            mismatchBtn.disabled = true;
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
                mismatchBtn.disabled = false;
                removeBtn.innerHTML = '<i class="fas fa-user-minus"></i> Remove User';
                
                if (data.error) {
                    setSingleStatus('Error: ' + data.error);
                } else {
                    setSingleStatus('User successfully removed from site');
                    displaySingleRemovalConfirmation(data);
                    singleCurrentUserId = null;
                    removeBtn.disabled = true;
                    mismatchBtn.disabled = true;
                }
            })
            .catch(err => {
                removeBtn.disabled = false;
                findBtn.disabled = false;
                mismatchBtn.disabled = false;
                removeBtn.innerHTML = '<i class="fas fa-user-minus"></i> Remove User';
                setSingleStatus('Request failed: ' + err.message);
            });
        }
        
        function removeUserForMismatch() {
            if (!singleCurrentUserId || !singleCurrentSiteUrl || !currentUserUpn) {
                setSingleStatus('Error: Please find a user first');
                return;
            }

            if (!confirm('WARNING: This will remove the user from the current site because of NameId mismatch. The user will need to be re-added with the correct permissions. Continue?')) {
                return;
            }

            const removeBtn = document.getElementById('single-remove-user-btn');
            const findBtn = document.getElementById('single-find-user-btn');
            const mismatchBtn = document.getElementById('single-remove-mismatch-btn');
            
            removeBtn.disabled = true;
            findBtn.disabled = true;
            mismatchBtn.disabled = true;
            mismatchBtn.innerHTML = '<div class="spinner"></div> Removing User for Mismatch...';
            setSingleStatus('Removing user due to NameId mismatch...');

            fetch('/remove_user_mismatch', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    upn: currentUserUpn,
                    current_site_url: singleCurrentSiteUrl
                })
            })
            .then(res => res.json())
            .then(data => {
                removeBtn.disabled = false;
                findBtn.disabled = false;
                mismatchBtn.disabled = false;
                mismatchBtn.innerHTML = '<i class="fas fa-exclamation-triangle"></i> Remove for Mismatch';
                
                if (data.error) {
                    setSingleStatus('Error: ' + data.error);
                } else if (data.warning) {
                    setSingleStatus('Warning: ' + data.warning);
                } else {
                    setSingleStatus('User successfully removed due to NameId mismatch');
                    displayMismatchRemovalResult(data);
                    singleCurrentUserId = null;
                    removeBtn.disabled = true;
                    mismatchBtn.disabled = true;
                }
            })
            .catch(err => {
                removeBtn.disabled = false;
                findBtn.disabled = false;
                mismatchBtn.disabled = false;
                mismatchBtn.innerHTML = '<i class="fas fa-exclamation-triangle"></i> Remove for Mismatch';
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
                <p><strong>User Principal Name:</strong> ${data.user_principal_name || 'N/A'}</p>
                <p><strong>Site Type:</strong> ${data.site_type || 'SharePoint Site'}</p>
                <p><strong>Site URL:</strong> ${data.site_url || 'N/A'}</p>
                <p><strong>OneDrive Owner:</strong> ${data.onedrive_owner || 'N/A'}</p>
                <p><strong>Is Site Admin:</strong> ${data.is_site_admin ? 'Yes' : 'No'}</p>
                <p><strong>Current NameId:</strong> ${data.current_nameid || 'N/A'}</p>
            `;
        }
        
        function displayNameIdComparison(comparison) {
            const comparisonDiv = document.getElementById('nameid-comparison');
            const contentDiv = document.getElementById('nameid-comparison-content');
            const mismatchBtn = document.getElementById('single-remove-mismatch-btn');

            const matchStatus = comparison.match ? 'nameid-match' : 'nameid-mismatch';
            const statusText = comparison.match ? 'MATCH' : 'MISMATCH';
            
            // Show/hide mismatch removal button
            if (comparison.match) {
                mismatchBtn.style.display = 'none';
            } else {
                mismatchBtn.style.display = 'inline-block';
                mismatchBtn.disabled = false;
            }

            let warningMessage = '';
            if (!comparison.match) {
                warningMessage = `
                    <div class="nameid-warning">
                        <strong><i class="fas fa-exclamation-triangle"></i> NameId Mismatch Detected</strong>
                        <p style="margin-top: 8px; margin-bottom: 0;">
                            This indicates a PUID mismatch issue. You can remove the user from this site 
                            and re-add them to fix the NameId. The removal button above will work regardless, 
                            but for automatic mismatch handling use the "Remove for Mismatch" button.
                        </p>
                    </div>
                `;
            }
            
            contentDiv.innerHTML = `
                <div class="nameid-comparison-grid">
                    <div class="nameid-box ${comparison.match ? 'match' : 'mismatch'}">
                        <h5 style="margin-bottom: 10px; color: #495057;">Current Site NameId</h5>
                        <p style="font-family: monospace; word-break: break-all; font-size: 0.9rem;">${comparison.current_nameid || 'Not found'}</p>
                    </div>
                    <div class="nameid-box ${comparison.match ? 'match' : 'mismatch'}">
                        <h5 style="margin-bottom: 10px; color: #495057;">New ID Site NameId</h5>
                        <p style="font-family: monospace; word-break: break-all; font-size: 0.9rem;">${comparison.new_nameid || 'Not found'}</p>
                    </div>
                </div>
                <div style="padding: 10px; background: ${comparison.match ? '#d4edda' : '#f8d7da'}; border-radius: 4px;">
                    <strong class="${matchStatus}">NameId Status: ${statusText}</strong>
                </div>
                ${warningMessage}
            `;
            
            comparisonDiv.style.display = 'block';
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
        
        function displayMismatchRemovalResult(data) {
            const confirmationDiv = document.getElementById('single-confirmation-section');
            
            confirmationDiv.innerHTML = `
                <div style="background: #f8d7da; padding: 15px; border-radius: 8px; margin-top: 15px;">
                    <h4 style="color: #721c24; margin-bottom: 10px;">
                        <i class="fas fa-exclamation-triangle"></i> User Removed Due to NameId Mismatch
                    </h4>
                    <p><strong>User:</strong> ${data.upn}</p>
                    <p><strong>Removed From:</strong> ${data.current_site_url}</p>
                    <p><strong>Current NameId:</strong> ${data.nameid_comparison.current_nameid || 'N/A'}</p>
                    <p><strong>New NameId:</strong> ${data.nameid_comparison.new_nameid || 'N/A'}</p>
                    <p><strong>Removal Time:</strong> ${data.removal_time}</p>
                    <p style="margin-top: 10px; font-weight: bold;">
                        The user should now be re-added to the site to get the correct NameId.
                    </p>
                </div>
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
            const removeAllBtn = document.getElementById('bulk-remove-all-btn');
            
            findBtn.disabled = true;
            removeAllBtn.disabled = true;
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
                    removeAllBtn.disabled = foundCount === 0;
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
        
        function removeSingleBulkUser(owner, userId, siteUrl) {
            if (!confirm(`Are you sure you want to remove the user from ${owner}'s OneDrive?`)) {
                return;
            }
            
            const button = event.target;
            const originalText = button.innerHTML;
            
            button.disabled = true;
            button.innerHTML = '<div class="spinner" style="width: 12px; height: 12px; border-width: 2px;"></div> Removing...';
            
            fetch('/remove_user', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    user_id: userId, 
                    site_url: siteUrl 
                })
            })
            .then(res => res.json())
            .then(data => {
                if (data.error) {
                    button.innerHTML = '<span class="status-error">Failed</span>';
                    button.style.background = '#dc3545';
                } else {
                    button.innerHTML = '<span class="status-success">Removed</span>';
                    button.style.background = '#28a745';
                    button.disabled = true;
                    
                    // Update the bulk results array
                    const resultIndex = bulkResults.findIndex(r => r.onedrive_owner === owner);
                    if (resultIndex !== -1) {
                        bulkResults[resultIndex].removal_success = true;
                    }
                }
            })
            .catch(err => {
                button.innerHTML = '<span class="status-error">Failed</span>';
                button.style.background = '#dc3545';
            });
        }
        
        function removeAllBulkUsers() {
            const foundUsers = bulkResults.filter(r => r.found);
            
            if (foundUsers.length === 0) {
                setBulkStatus('Error: No users found to remove');
                return;
            }
            
            let confirmationMessage = `Are you sure you want to remove the user from ${foundUsers.length} OneDrive sites? This action cannot be undone.`;
            
            // Check for any NameId mismatches in the results
            const hasMismatches = bulkResults.some(r => r.found && r.current_nameid && !r.nameid_match);
            if (hasMismatches) {
                confirmationMessage = `WARNING: NameId mismatches detected in some OneDrives! This indicates PUID issues. Are you sure you want to remove the user from ${foundUsers.length} OneDrive sites? The removal will work, but users may need OneDrive repair.`;
            }
            
            if (!confirm(confirmationMessage)) {
                return;
            }
            
            const findBtn = document.getElementById('bulk-find-btn');
            const removeAllBtn = document.getElementById('bulk-remove-all-btn');
            
            findBtn.disabled = true;
            removeAllBtn.disabled = true;
            removeAllBtn.innerHTML = '<div class="spinner"></div> Removing Users...';
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
                removeAllBtn.disabled = false;
                removeAllBtn.innerHTML = '<i class="fas fa-user-minus"></i> Remove from All Found OneDrives';
                
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
                removeAllBtn.disabled = false;
                removeAllBtn.innerHTML = '<i class="fas fa-user-minus"></i> Remove from All Found OneDrives';
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
                
                // NameId Comparison
                const nameidCell = row.insertCell(2);
                if (result.found && result.current_nameid) {
                    const matchStatus = result.nameid_match ? 'match' : 'mismatch';
                    const statusText = result.nameid_match ? 'MATCH' : 'MISMATCH';
                    
                    nameidCell.innerHTML = `
                        <div style="font-size: 0.8rem;">
                            <div><strong>Current:</strong> ${result.current_nameid || 'N/A'}</div>
                            <div><strong>New ID:</strong> ${result.new_nameid || 'N/A'}</div>
                            <div style="margin-top: 5px; padding: 2px 5px; background: ${result.nameid_match ? '#d4edda' : '#f8d7da'}; border-radius: 3px;">
                                <strong class="${result.nameid_match ? 'nameid-match' : 'nameid-mismatch'}">${statusText}</strong>
                            </div>
                        </div>
                    `;
                } else {
                    nameidCell.innerHTML = '<span class="status-pending">N/A</span>';
                }
                
                // Actions
                const actionsCell = row.insertCell(3);
                if (result.found) {
                    if (result.removal_success) {
                        actionsCell.innerHTML = '<span class="status-success">✓ Removed</span>';
                    } else if (result.removal_success === false) {
                        actionsCell.innerHTML = '<span class="status-error">✗ Failed</span>';
                    } else {
                        actionsCell.innerHTML = `
                            <div class="action-buttons">
                                <button class="btn btn-success btn-sm" onclick="removeSingleBulkUser('${result.onedrive_owner}', '${result.user_id}', '${result.site_url}')">
                                    <i class="fas fa-user-minus"></i> Remove
                                </button>
                                ${!result.nameid_match ? `
                                <button class="btn btn-warning btn-sm" onclick="repairOneDriveForBulk('${result.onedrive_owner}', '${result.site_url}')">
                                    <i class="fas fa-tools"></i> Repair
                                </button>
                                ` : ''}
                            </div>
                        `;
                    }
                } else {
                    actionsCell.innerHTML = '<span class="status-pending">N/A</span>';
                }
            });
        }
        
        function repairOneDriveForBulk(owner, siteUrl) {
            if (!confirm(`Are you sure you want to repair OneDrive for ${owner}? This will reset permissions and may take several minutes.`)) {
                return;
            }
            
            const button = event.target;
            const originalText = button.innerHTML;
            
            button.disabled = true;
            button.innerHTML = '<div class="spinner" style="width: 12px; height: 12px; border-width: 2px;"></div> Repairing...';
            
            // Extract UPN from owner (assuming owner is the UPN)
            const upn = owner;
            
            fetch('/repair_onedrive', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ 
                    upn: upn,
                    onedrive_url: siteUrl + '/Documents' // Convert site URL back to OneDrive URL
                })
            })
            .then(res => res.json())
            .then(data => {
                if (data.error) {
                    button.innerHTML = '<span class="status-error">Failed</span>';
                    button.style.background = '#dc3545';
                } else {
                    button.innerHTML = '<span class="status-success">Repaired</span>';
                    button.style.background = '#28a745';
                    button.disabled = true;
                }
            })
            .catch(err => {
                button.innerHTML = '<span class="status-error">Failed</span>';
                button.style.background = '#dc3545';
            });
        }
        
        // Initialize the application
        document.addEventListener('DOMContentLoaded', function() {
            // Allow Enter key to submit
            document.getElementById('user-upn').addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    getOneDriveInfo();
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

def parse_site_users_json(json_content, target_upn):
    """Parse SharePoint site users JSON response and find specific user"""
    try:
        logger.debug(f"Parsing JSON for user: {target_upn}")
        
        # Handle both direct JSON object and response.json() structure
        if isinstance(json_content, dict):
            data = json_content
        else:
            data = json_content.json() if hasattr(json_content, 'json') else json_content
        
        # Handle different response formats
        if 'd' in data:
            # OData format with 'd' wrapper
            if 'results' in data['d']:
                users = data['d']['results']
            else:
                users = [data['d']]  # Single user
        elif 'value' in data:
            # Direct value array
            users = data['value']
        else:
            # Assume it's already a list or single user
            users = data if isinstance(data, list) else [data]
        
        logger.debug(f"Processing {len(users)} users")
        
        for user in users:
            # Extract user properties with safe access
            user_id = user.get('Id')
            title = user.get('Title')
            email = user.get('Email')
            login_name = user.get('LoginName')
            user_principal_name = user.get('UserPrincipalName')
            is_site_admin = user.get('IsSiteAdmin', False)
            
            # Extract NameId from UserId if available
            current_nameid = None
            user_id_obj = user.get('UserId')
            if user_id_obj and isinstance(user_id_obj, dict):
                current_nameid = user_id_obj.get('NameId')
            
            logger.debug(f"Checking user - ID: {user_id}, Email: {email}, Login: {login_name}, UPN: {user_principal_name}, NameId: {current_nameid}")
            
            # Check if this is the target user using multiple identifiers
            email_match = email and email.lower() == target_upn.lower()
            login_match = login_name and target_upn.lower() in login_name.lower()
            upn_match = user_principal_name and user_principal_name.lower() == target_upn.lower()
            
            if email_match or login_match or upn_match:
                user_info = {
                    'user_id': user_id,
                    'title': title,
                    'email': email,
                    'login_name': login_name,
                    'user_principal_name': user_principal_name,
                    'is_site_admin': is_site_admin,
                    'current_nameid': current_nameid
                }
                logger.info(f"Found target user: {user_info}")
                return user_info
        
        logger.warning(f"User {target_upn} not found in site users")
        return None
        
    except Exception as e:
        logger.exception("Failed to parse site users JSON")
        raise Exception(f"Failed to parse JSON response: {str(e)}")

def get_new_site_nameid(target_upn):
    """Get NameId for user from new ID site by ensuring user exists there"""
    try:
        new_site_url = CONFIG['new_id_site_url']
        if not new_site_url.endswith('/'):
            new_site_url += '/'
        
        # Get SharePoint token
        sharepoint_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sharepoint_token:
            sharepoint_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
        if not sharepoint_token:
            logger.error("Failed to obtain SharePoint access token for new ID site")
            return None
        
        # Get request digest
        request_digest = get_request_digest(new_site_url, sharepoint_token)
        if not request_digest:
            logger.error("Failed to get request digest for new ID site")
            return None
        
        # Ensure user exists on new ID site
        ensure_url = f"{new_site_url}_api/web/ensureuser"
        headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": request_digest
        }
        
        body = {
            "logonName": target_upn
        }
        
        logger.info(f"Ensuring user {target_upn} on new ID site")
        response = requests.post(ensure_url, headers=headers, json=body)
        
        logger.info(f"Ensure user response status: {response.status_code}")
        
        if response.status_code == 200:
            user_data = response.json()
            logger.debug(f"Ensure user response: {json.dumps(user_data, indent=2)}")
            
            user_info = user_data.get('d', {})
            
            # Extract NameId from UserId
            user_id_obj = user_info.get('UserId')
            if user_id_obj and isinstance(user_id_obj, dict):
                nameid = user_id_obj.get('NameId')
                logger.info(f"Retrieved NameId from new ID site: {nameid}")
                return nameid
            else:
                logger.warning("UserId object not found or invalid in new ID site response")
                # Try alternative extraction
                if 'Id' in user_info:
                    logger.info(f"User ensured but no NameId found. User ID: {user_info.get('Id')}")
                return None
        else:
            logger.error(f"Failed to ensure user on new ID site: {response.status_code} - {response.text}")
            return None
            
    except Exception as e:
        logger.exception(f"Error getting NameId from new ID site for user {target_upn}")
        return None

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
        
        # Call SharePoint API to get site users with detailed logging
        sharepoint_headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/json;odata=verbose"
        }
        
        logger.info("Calling SharePoint API to get site users")
        site_users_response = requests.get(site_users_url, headers=sharepoint_headers)
        
        logger.info(f"Response status: {site_users_response.status_code}")
        logger.info(f"Response headers: {site_users_response.headers}")
        
        if site_users_response.status_code != 200:
            logger.error(f"SharePoint API failed: {site_users_response.status_code} - {site_users_response.text}")
            raise Exception(f"Failed to get site users: {site_users_response.text}")
        
        # Log the raw response for debugging
        logger.debug(f"Raw response: {site_users_response.text}")
        
        # Parse JSON and find the specific user
        user_info = parse_site_users_json(site_users_response, target_upn)
        
        if not user_info:
            # Try alternative approach - get user by login name directly
            logger.info(f"User not found in site users list, trying direct lookup...")
            user_info = get_user_by_login_name(site_url, sharepoint_token, target_upn)
        
        return user_info
        
    except Exception as e:
        logger.exception(f"Failed to find user {target_upn} on site {site_url}")
        raise

def get_user_by_login_name(site_url, token, target_upn):
    """Try to get user directly by login name"""
    try:
        # Try different login name formats
        login_formats = [
            target_upn,
            f"i:0#.f|membership|{target_upn}",
            f"i:0#.f|membership|{target_upn.lower()}",
            f"i:0%23.f|membership|{target_upn}",
            f"i:0%23.f|membership|{target_upn.lower()}"
        ]
        
        for login_format in login_formats:
            try:
                user_url = f"{site_url}_api/web/siteusers('{login_format}')"
                headers = {
                    "Authorization": f"Bearer {token}",
                    "Accept": "application/json;odata=verbose"
                }
                
                logger.info(f"Trying direct user lookup with: {login_format}")
                response = requests.get(user_url, headers=headers)
                
                if response.status_code == 200:
                    user_data = response.json()
                    user_info = parse_site_users_json(user_data, target_upn)
                    if user_info:
                        logger.info(f"Found user using direct lookup with format: {login_format}")
                        return user_info
                
            except Exception as e:
                logger.debug(f"Direct lookup failed for format {login_format}: {str(e)}")
                continue
        
        return None
        
    except Exception as e:
        logger.exception("Error in direct user lookup")
        return None

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
        ensure_url = f"{site_url}_api/web/ensureuser"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": request_digest
        }
        
        body = {
            "logonName": user_upn
        }
        
        response = requests.post(ensure_url, headers=headers, json=body)
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
        admin_url = f"{site_url}_api/web/getuserbyid({user_id})"
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
        remove_url = f"{site_url}_api/web/siteusers/removebyid({user_id})"
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

def is_user_site_owner(site_url, token, user_upn):
    """Check if user is the site owner"""
    try:
        # Get site information to check owner
        site_info_url = f"{site_url}_api/web"
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json;odata=verbose"
        }
        
        response = requests.get(site_info_url, headers=headers)
        if response.status_code == 200:
            site_data = response.json()
            # Check various owner properties
            owner_login = site_data.get('d', {}).get('Owner', {}).get('LoginName', '')
            owner_email = site_data.get('d', {}).get('Owner', {}).get('Email', '')
            
            # Check if the user is the owner
            return (user_upn.lower() in owner_login.lower() or 
                   user_upn.lower() == owner_email.lower())
        
        return False
    except Exception as e:
        logger.exception("Error checking site owner")
        return False

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
            # Check if user is site owner
            is_owner = is_user_site_owner(site_url, sp_token, upn)
            logger.info(f"User {upn} is site owner: {is_owner}")
            
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
            
            # If user is site owner, we need to handle this specially
            if is_owner:
                logger.info(f"User {upn} is site owner, using special handling")
                result['steps'].append({"step": 2, "status": "running", "message": "User is site owner - using special handling"})
                # For site owners, we don't remove them completely, just remove admin rights temporarily
                original_user_id = get_user_id_from_site(site_url, sp_token, upn)
                if original_user_id:
                    if not set_site_admin(site_url, sp_token, request_digest, original_user_id, False):
                        logger.warning("Failed to remove original user admin rights, continuing...")
                result['steps'].append({"step": 2, "status": "success"})
            else:
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
            
            try:
                site_url = get_onedrive_url(onedrive_owner_upn)
                site_type = 'OneDrive'
                onedrive_owner = onedrive_owner_upn
            except Exception as e:
                logger.error(f"Failed to get OneDrive URL for {onedrive_owner_upn}: {str(e)}")
                return jsonify({"error": f"Failed to get OneDrive URL: {str(e)}"}), 400
        
        logger.info(f"Searching for user {upn} on {site_url}")
        
        # Find user on site
        try:
            user_info = find_user_on_site(upn, site_url)
        except Exception as e:
            logger.error(f"Error finding user: {str(e)}")
            return jsonify({"error": f"Failed to find user: {str(e)}"}), 500
        
        if not user_info:
            return jsonify({"error": f"User '{upn}' not found on the specified {site_type.lower()}. The user may not have explicit permissions or may be accessing the site through a group."}), 404
        
        # Get NameId from new ID site for comparison
        logger.info(f"Getting NameId from new ID site for {upn}")
        new_nameid = get_new_site_nameid(upn)
        
        # Prepare NameId comparison
        nameid_comparison = {
            'current_nameid': user_info.get('current_nameid'),
            'new_nameid': new_nameid,
            'match': user_info.get('current_nameid') == new_nameid if user_info.get('current_nameid') and new_nameid else False
        }
        
        # Add additional info to response
        user_info['site_url'] = site_url
        user_info['site_type'] = site_type
        user_info['onedrive_owner'] = onedrive_owner
        user_info['nameid_comparison'] = nameid_comparison
        
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

@app.route('/remove_user_mismatch', methods=['POST'])
def remove_user_mismatch():
    """Remove user from current site when NameId doesn't match new ID site"""
    try:
        data = request.json
        upn = data.get('upn', '').strip()
        current_site_url = data.get('current_site_url', '').strip()
        
        if not upn or not current_site_url:
            return jsonify({"error": "User UPN and current site URL are required"}), 400
        
        # Ensure site URL ends properly
        if not current_site_url.endswith('/'):
            current_site_url += '/'
        
        logger.info(f"Processing mismatch removal for {upn} from {current_site_url}")
        
        # Get SharePoint token
        sp_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sp_token:
            sp_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
        if not sp_token:
            return jsonify({"error": "Failed to get SharePoint token"}), 500
        
        # Step 1: Find user on current site
        user_info = find_user_on_site(upn, current_site_url)
        if not user_info:
            return jsonify({"error": f"User '{upn}' not found on current site"}), 404
        
        current_nameid = user_info.get('current_nameid')
        user_id = user_info['user_id']
        
        # Step 2: Get new NameId from new ID site
        new_nameid = get_new_site_nameid(upn)
        
        if not new_nameid:
            return jsonify({"error": "Could not retrieve new NameId from reference site"}), 400
        
        # Step 3: Check if mismatch exists
        nameid_match = current_nameid == new_nameid if current_nameid and new_nameid else False
        
        if nameid_match:
            return jsonify({
                "warning": "NameIds match - no removal needed",
                "current_nameid": current_nameid,
                "new_nameid": new_nameid,
                "match": True
            }), 200
        
        # Step 4: Remove user from current site (mismatch confirmed)
        logger.info(f"NameId mismatch detected. Removing user {upn} from current site")
        remove_user_from_site(user_id, current_site_url)
        
        return jsonify({
            "success": True,
            "message": "User successfully removed due to NameId mismatch",
            "user_id": user_id,
            "upn": upn,
            "current_site_url": current_site_url,
            "nameid_comparison": {
                "current_nameid": current_nameid,
                "new_nameid": new_nameid,
                "match": False
            },
            "removal_time": time.strftime('%Y-%m-%d %H:%M:%S')
        })
        
    except Exception as e:
        logger.exception("Error during mismatch removal")
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
                'user_principal_name': None,
                'is_site_admin': False,
                'current_nameid': None,
                'new_nameid': None,
                'nameid_match': False
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
                    result['user_principal_name'] = user_info.get('user_principal_name')
                    result['is_site_admin'] = user_info['is_site_admin']
                    result['current_nameid'] = user_info.get('current_nameid')
                    
                    # Get NameId from new ID site for comparison
                    new_nameid = get_new_site_nameid(target_upn)
                    result['new_nameid'] = new_nameid
                    result['nameid_match'] = result['current_nameid'] == new_nameid if result['current_nameid'] and new_nameid else False
                    
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

@app.route('/debug_site_users', methods=['POST'])
def debug_site_users():
    """Debug endpoint to see all users on a site"""
    try:
        data = request.json
        site_url = data.get('site_url', '').strip()
        
        if not site_url:
            return jsonify({"error": "Site URL is required"}), 400
        
        if not site_url.endswith('/'):
            site_url += '/'
        
        # Get SharePoint token
        sharepoint_token = get_token_with_certificate(CONFIG['scopes']['sharepoint'])
        if not sharepoint_token:
            sharepoint_token = get_token_with_secret(CONFIG['scopes']['sharepoint'])
        
        if not sharepoint_token:
            return jsonify({"error": "Failed to obtain SharePoint access token"}), 500
        
        # Call SharePoint API to get site users
        site_users_url = f"{site_url}_api/web/siteusers"
        headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/json;odata=verbose"
        }
        
        response = requests.get(site_users_url, headers=headers)
        
        if response.status_code != 200:
            return jsonify({"error": f"Failed to get site users: {response.text}"}), 400
        
        users_data = response.json()
        
        # Extract user information
        users = []
        if 'd' in users_data and 'results' in users_data['d']:
            for user in users_data['d']['results']:
                users.append({
                    'Id': user.get('Id'),
                    'Title': user.get('Title'),
                    'Email': user.get('Email'),
                    'LoginName': user.get('LoginName'),
                    'UserPrincipalName': user.get('UserPrincipalName'),
                    'IsSiteAdmin': user.get('IsSiteAdmin', False),
                    'UserId': user.get('UserId')
                })
        
        return jsonify({
            "site_url": site_url,
            "total_users": len(users),
            "users": users
        })
        
    except Exception as e:
        logger.exception("Error in debug_site_users")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5500, debug=True)
