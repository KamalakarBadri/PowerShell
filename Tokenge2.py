from flask import Flask, render_template_string, request, jsonify
import requests
import json
import os
import time
import uuid
import base64
import logging
from datetime import datetime, timedelta
from cryptography.hazmat.primitives import serialization, hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.hazmat.backends import default_backend

app = Flask(__name__)

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuration file - Update these with your tenant details
TENANT_CONFIG = {
    "tenant1": {
        "name": "Geekbyte Online (Primary)",
        "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
        "tenant_name": "TestT.onmicrosoft.com",
        "app_id": "73efa35d-6188-42d4-b258-838a977eb149",
        "client_secret": "CyG8Q~FYHuCMSyVmt4sNxt5IejrMc2c24Ziz4a.t",
        "certificate_path": "certificate.pem",
        "private_key_path": "private_key.pem",
        "repair_account": "edit@geekbyte.online",
        "scopes": {
            "graph": "https://graph.microsoft.com/.default",
            "sharepoint": "https://TestT.sharepoint.com/.default",
            "forms": "https://forms.office.com/.default"
        }
    },
    "tenant2": {
        "name": "Tenant 2 (Secondary)",
        "tenant_id": "your-tenant2-id",  # Replace with your second tenant ID
        "tenant_name": "your-tenant2.onmicrosoft.com",  # Replace with your second tenant name
        "app_id": "your-app-id-2",  # Replace with your second app ID
        "client_secret": "your-client-secret-2",  # Replace with your second client secret
        "certificate_path": "certificate2.pem",  # Replace with your second certificate path
        "private_key_path": "private_key2.pem",  # Replace with your second private key path
        "repair_account": "admin@tenant2.com",  # Replace with your second repair account
        "scopes": {
            "graph": "https://graph.microsoft.com/.default",
            "sharepoint": "https://tenant2.sharepoint.com/.default",  # Replace with your second SharePoint URL
            "forms": "https://forms.cloud.microsoft/.default"
        }
    }
}

def get_token_with_certificate(tenant_config, scope):
    """Get access token using certificate-based authentication"""
    try:
        cert_path = tenant_config['certificate_path']
        key_path = tenant_config['private_key_path']
        
        if not os.path.exists(cert_path) or not os.path.exists(key_path):
            logger.warning(f"Certificate files not found: {cert_path} or {key_path}")
            return None
            
        with open(cert_path, "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
        with open(key_path, "rb") as key_file:
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
            logger.info("Successfully obtained token using certificate")
            return token_response.json()
        else:
            logger.error(f"Certificate token request failed: {token_response.text}")
            return None
            
    except Exception as e:
        logger.exception("Certificate authentication failed")
        return None

def get_token_with_client_secret(tenant_config, scope):
    """Get access token using client secret authentication"""
    try:
        tenant_id = tenant_config['tenant_id']
        if '.onmicrosoft.com' in tenant_id or len(tenant_id.split('-')) == 5:
            token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
        else:
            token_url = f"https://login.microsoftonline.com/{tenant_id}.onmicrosoft.com/oauth2/v2.0/token"
        
        token_data = {
            'grant_type': 'client_credentials',
            'client_id': tenant_config['app_id'],
            'client_secret': tenant_config['client_secret'],
            'scope': scope
        }
        
        response = requests.post(token_url, data=token_data)
        
        if response.status_code == 200:
            logger.info("Successfully obtained token using client secret")
            return response.json()
        else:
            logger.error(f"Client secret token request failed: {response.text}")
            return None
            
    except Exception as e:
        logger.exception("Client secret authentication failed")
        return None

# HTML template for the web interface
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Multi-Tenant SharePoint Access Token Generator</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            max-width: 1000px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f5f5f5;
        }
        .container {
            background: white;
            padding: 30px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        h1 {
            color: #0078d4;
            text-align: center;
            margin-bottom: 30px;
        }
        .form-group {
            margin-bottom: 20px;
        }
        label {
            display: block;
            margin-bottom: 8px;
            font-weight: 600;
            color: #323130;
        }
        select {
            width: 100%;
            padding: 12px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
            box-sizing: border-box;
            background-color: white;
        }
        .tenant-info {
            background-color: #e7f3ff;
            border: 1px solid #b3d9ff;
            padding: 15px;
            border-radius: 4px;
            margin-top: 10px;
            font-size: 13px;
        }
        .auth-method-container {
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 15px;
            background-color: #fafafa;
            margin-bottom: 20px;
        }
        .auth-method-item {
            display: flex;
            align-items: center;
            margin-bottom: 15px;
            padding: 10px;
            background-color: white;
            border-radius: 4px;
            border: 1px solid #e0e0e0;
        }
        .auth-method-item input[type="radio"] {
            margin-right: 15px;
            transform: scale(1.3);
        }
        .auth-method-item label {
            margin-bottom: 0;
            font-weight: normal;
            cursor: pointer;
            flex: 1;
        }
        .scopes-container {
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 15px;
            background-color: #fafafa;
        }
        .scope-item {
            display: flex;
            align-items: center;
            margin-bottom: 15px;
            padding: 10px;
            background-color: white;
            border-radius: 4px;
            border: 1px solid #e0e0e0;
        }
        .scope-item input[type="checkbox"] {
            margin-right: 15px;
            transform: scale(1.3);
        }
        .scope-item label {
            margin-bottom: 0;
            font-weight: normal;
            cursor: pointer;
            flex: 1;
        }
        .scope-description {
            font-size: 12px;
            color: #666;
            margin-top: 5px;
        }
        .btn {
            background-color: #0078d4;
            color: white;
            padding: 12px 30px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
            width: 100%;
            margin-top: 20px;
        }
        .btn:hover {
            background-color: #106ebe;
        }
        .btn:disabled {
            background-color: #ccc;
            cursor: not-allowed;
        }
        .results-container {
            margin-top: 30px;
        }
        .token-result {
            margin-bottom: 25px;
            padding: 20px;
            border-radius: 8px;
            font-family: 'Courier New', monospace;
        }
        .token-success {
            background-color: #d4edda;
            border: 1px solid #c3e6cb;
            color: #155724;
        }
        .token-error {
            background-color: #f8d7da;
            border: 1px solid #f5c6cb;
            color: #721c24;
        }
        .loading {
            text-align: center;
            color: #0078d4;
            padding: 20px;
        }
        .token-info {
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 4px;
            margin-top: 15px;
        }
        .copy-btn {
            background-color: #28a745;
            color: white;
            border: none;
            padding: 8px 15px;
            border-radius: 3px;
            cursor: pointer;
            font-size: 12px;
            margin-left: 10px;
        }
        .copy-btn:hover {
            background-color: #218838;
        }
        .token-header {
            font-size: 16px;
            font-weight: bold;
            color: #0078d4;
            margin-bottom: 10px;
            font-family: 'Segoe UI', sans-serif;
        }
        .warning {
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            padding: 15px;
            border-radius: 4px;
            margin-bottom: 20px;
            color: #856404;
        }
        .config-info {
            background-color: #d1ecf1;
            border: 1px solid #bee5eb;
            padding: 15px;
            border-radius: 4px;
            margin-bottom: 20px;
            font-size: 13px;
        }
        .cert-status {
            font-size: 12px;
            margin-top: 5px;
            padding: 5px 8px;
            border-radius: 3px;
        }
        .cert-available {
            background-color: #d4edda;
            color: #155724;
            border: 1px solid #c3e6cb;
        }
        .cert-unavailable {
            background-color: #f8d7da;
            color: #721c24;
            border: 1px solid #f5c6cb;
        }
        .auth-method-description {
            font-size: 12px;
            color: #666;
            margin-top: 5px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🔐 Multi-Tenant SharePoint Access Token Generator</h1>
        
        <div class="config-info">
            <strong>📋 Configuration Status:</strong> Update the TENANT_CONFIG in the Python file with your actual tenant details and certificate paths.
            <br><strong>📦 Required Dependencies:</strong> pip install flask requests cryptography
        </div>
        
        <form id="tokenForm">
            <div class="form-group">
                <label for="tenant_select">Select Tenant:</label>
                <select id="tenant_select" name="tenant_select" required>
                    <option value="">Choose a tenant...</option>
                    {% for key, tenant in tenants.items() %}
                    <option value="{{ key }}">{{ tenant.name }}</option>
                    {% endfor %}
                </select>
                <div id="tenant_info" class="tenant-info" style="display: none;"></div>
            </div>
            
            <div class="form-group">
                <label>Select Authentication Method:</label>
                <div class="auth-method-container">
                    <div class="auth-method-item">
                        <input type="radio" id="auth_auto" name="auth_method" value="auto" checked>
                        <div>
                            <label for="auth_auto">🔄 Auto (Certificate → Client Secret)</label>
                            <div class="auth-method-description">Try certificate first, fallback to client secret if certificate fails</div>
                        </div>
                    </div>
                    <div class="auth-method-item">
                        <input type="radio" id="auth_certificate" name="auth_method" value="certificate">
                        <div>
                            <label for="auth_certificate">🔒 Certificate Only</label>
                            <div class="auth-method-description">Use certificate-based authentication (more secure)</div>
                        </div>
                    </div>
                    <div class="auth-method-item">
                        <input type="radio" id="auth_secret" name="auth_method" value="secret">
                        <div>
                            <label for="auth_secret">🔑 Client Secret Only</label>
                            <div class="auth-method-description">Use client secret authentication (simpler setup)</div>
                        </div>
                    </div>
                </div>
            </div>
            
            <div class="form-group">
                <label>Select Required Scopes:</label>
                <div class="scopes-container" id="scopes_container">
                    <!-- Scopes will be populated based on selected tenant -->
                </div>
                <div class="warning">
                    <strong>⚠️ Important:</strong> Each selected scope will generate a separate access token. This allows you to use scope-specific tokens for different API calls.
                </div>
            </div>
            
            <button type="submit" class="btn" id="generateBtn">Generate Access Tokens</button>
        </form>
        
        <div id="results"></div>
    </div>

    <script>
        // Show tenant info and scopes when tenant is selected
        document.getElementById('tenant_select').addEventListener('change', function() {
            const tenantInfo = document.getElementById('tenant_info');
            const scopesContainer = document.getElementById('scopes_container');
            const selectedTenant = this.value;
            
            if (selectedTenant) {
                const tenants = {{ tenants | tojson }};
                const tenant = tenants[selectedTenant];
                
                // Show tenant info
                tenantInfo.innerHTML = `
                    <strong>Tenant ID:</strong> ${tenant.tenant_id}<br>
                    <strong>Tenant Name:</strong> ${tenant.tenant_name}<br>
                    <strong>App ID:</strong> ${tenant.app_id}<br>
                    <strong>Client Secret:</strong> ${'*'.repeat(20)}...<br>
                    <strong>Certificate Path:</strong> ${tenant.certificate_path}
                    <div class="cert-status" id="cert_status_${selectedTenant}">📁 Certificate status will be checked when generating tokens</div>
                `;
                tenantInfo.style.display = 'block';
                
                // Populate scopes for selected tenant
                let scopesHTML = '';
                Object.keys(tenant.scopes).forEach(scopeKey => {
                    const scopeUrl = tenant.scopes[scopeKey];
                    scopesHTML += `
                        <div class="scope-item">
                            <input type="checkbox" id="scope_${scopeKey}" name="scopes" value="${scopeKey}">
                            <div>
                                <label for="scope_${scopeKey}">${scopeKey.charAt(0).toUpperCase() + scopeKey.slice(1)} API</label>
                                <div class="scope-description">${scopeUrl}</div>
                            </div>
                        </div>
                    `;
                });
                scopesContainer.innerHTML = scopesHTML;
                
            } else {
                tenantInfo.style.display = 'none';
                scopesContainer.innerHTML = '';
            }
        });

        document.getElementById('tokenForm').addEventListener('submit', async function(e) {
            e.preventDefault();
            
            const formData = new FormData(e.target);
            const data = {
                tenant: formData.get('tenant_select'),
                auth_method: formData.get('auth_method'),
                scopes: formData.getAll('scopes')
            };
            
            if (!data.tenant) {
                alert('Please select a tenant.');
                return;
            }
            
            if (data.scopes.length === 0) {
                alert('Please select at least one scope.');
                return;
            }
            
            const resultsDiv = document.getElementById('results');
            const generateBtn = document.getElementById('generateBtn');
            
            generateBtn.disabled = true;
            generateBtn.textContent = 'Generating Tokens...';
            resultsDiv.innerHTML = '<div class="loading">🔄 Generating access tokens using selected authentication method...</div>';
            
            try {
                const response = await fetch('/generate_tokens', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify(data)
                });
                
                const result = await response.json();
                
                let resultHTML = '<div class="results-container">';
                
                if (result.success) {
                    resultHTML += `<h3>✅ Access Tokens Generated Successfully!</h3>`;
                    
                    result.tokens.forEach(tokenResult => {
                        if (tokenResult.success) {
                            resultHTML += `
                                <div class="token-result token-success">
                                    <div class="token-header">${tokenResult.scope_name} Token (${tokenResult.auth_method})</div>
                                    <div class="token-info">
                                        <strong>Access Token:</strong>
                                        <button class="copy-btn" onclick="copyToClipboard('${tokenResult.access_token}')">Copy Token</button>
                                        <br><br>
                                        <div style="word-break: break-all; background: #e9ecef; padding: 10px; border-radius: 3px; font-size: 10px; max-height: 100px; overflow-y: auto;">
                                            ${tokenResult.access_token}
                                        </div>
                                        <br>
                                        <strong>Scope:</strong> ${tokenResult.scope_url}<br>
                                        <strong>Auth Method:</strong> ${tokenResult.auth_method}<br>
                                        <strong>Token Type:</strong> ${tokenResult.token_type}<br>
                                        <strong>Expires In:</strong> ${tokenResult.expires_in} seconds<br>
                                        <strong>Expires At:</strong> ${tokenResult.expires_at}
                                    </div>
                                </div>
                            `;
                        } else {
                            resultHTML += `
                                <div class="token-result token-error">
                                    <div class="token-header">${tokenResult.scope_name} Token - FAILED</div>
                                    <strong>Auth Method Attempted:</strong> ${tokenResult.auth_method || 'Unknown'}<br>
                                    <strong>Error:</strong> ${tokenResult.error}<br>
                                    <strong>Description:</strong> ${tokenResult.error_description || 'No additional description provided'}
                                </div>
                            `;
                        }
                    });
                } else {
                    resultHTML += `
                        <div class="token-result token-error">
                            <h3>❌ Error Generating Tokens</h3>
                            <strong>Error:</strong> ${result.error}
                        </div>
                    `;
                }
                
                resultHTML += '</div>';
                resultsDiv.innerHTML = resultHTML;
                
            } catch (error) {
                resultsDiv.innerHTML = `
                    <div class="results-container">
                        <div class="token-result token-error">
                            <h3>❌ Request Failed</h3>
                            <strong>Error:</strong> ${error.message}
                        </div>
                    </div>
                `;
            } finally {
                generateBtn.disabled = false;
                generateBtn.textContent = 'Generate Access Tokens';
            }
        });
        
        function copyToClipboard(text) {
            navigator.clipboard.writeText(text).then(function() {
                alert('Access token copied to clipboard!');
            }, function(err) {
                alert('Failed to copy token: ' + err);
            });
        }
    </script>
</body>
</html>
"""

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE, tenants=TENANT_CONFIG)

@app.route('/generate_tokens', methods=['POST'])
def generate_tokens():
    try:
        data = request.get_json()
        
        tenant_key = data.get('tenant')
        auth_method = data.get('auth_method', 'auto')
        selected_scopes = data.get('scopes', [])
        
        if not tenant_key or tenant_key not in TENANT_CONFIG:
            return jsonify({
                'success': False,
                'error': 'Invalid tenant selection'
            })
        
        if not selected_scopes:
            return jsonify({
                'success': False,
                'error': 'No scopes selected'
            })
        
        tenant_config = TENANT_CONFIG[tenant_key]
        token_results = []
        
        # Generate separate token for each selected scope
        for scope_key in selected_scopes:
            if scope_key not in tenant_config['scopes']:
                token_results.append({
                    'success': False,
                    'scope_name': scope_key,
                    'error': 'Invalid scope',
                    'error_description': f'Scope {scope_key} not found for this tenant'
                })
                continue
                
            try:
                scope_url = tenant_config['scopes'][scope_key]
                token_data = None
                auth_used = None
                
                # Try authentication based on selected method
                if auth_method == 'certificate':
                    # Certificate only
                    token_data = get_token_with_certificate(tenant_config, scope_url)
                    auth_used = 'Certificate'
                    if not token_data:
                        token_results.append({
                            'success': False,
                            'scope_name': scope_key.title(),
                            'scope_url': scope_url,
                            'auth_method': 'Certificate',
                            'error': 'Certificate authentication failed',
                            'error_description': 'Certificate files not found or authentication failed'
                        })
                        continue
                        
                elif auth_method == 'secret':
                    # Client secret only
                    token_data = get_token_with_client_secret(tenant_config, scope_url)
                    auth_used = 'Client Secret'
                    if not token_data:
                        token_results.append({
                            'success': False,
                            'scope_name': scope_key.title(),
                            'scope_url': scope_url,
                            'auth_method': 'Client Secret',
                            'error': 'Client secret authentication failed',
                            'error_description': 'Client secret authentication failed'
                        })
                        continue
                        
                else:  # auto
                    # Try certificate first, then client secret
                    token_data = get_token_with_certificate(tenant_config, scope_url)
                    if token_data:
                        auth_used = 'Certificate'
                    else:
                        token_data = get_token_with_client_secret(tenant_config, scope_url)
                        if token_data:
                            auth_used = 'Client Secret (Fallback)'
                        else:
                            token_results.append({
                                'success': False,
                                'scope_name': scope_key.title(),
                                'scope_url': scope_url,
                                'auth_method': 'Auto (Both Failed)',
                                'error': 'Both authentication methods failed',
                                'error_description': 'Neither certificate nor client secret authentication succeeded'
                            })
                            continue
                
                if token_data and 'access_token' in token_data:
                    # Calculate expiration time
                    expires_in = token_data.get('expires_in', 3600)
                    expires_at = datetime.now() + timedelta(seconds=expires_in)
                    
                    token_results.append({
                        'success': True,
                        'scope_name': scope_key.title(),
                        'scope_url': scope_url,
                        'auth_method': auth_used,
                        'access_token': token_data['access_token'],
                        'token_type': token_data.get('token_type', 'Bearer'),
                        'expires_in': expires_in,
                        'expires_at': expires_at.strftime('%Y-%m-%d %H:%M:%S')
                    })
                else:
                    token_results.append({
                        'success': False,
                        'scope_name': scope_key.title(),
                        'scope_url': scope_url,
                        'auth_method': auth_used,
                        'error': token_data.get('error', 'Unknown error') if token_data else 'No token data received',
                        'error_description': token_data.get('error_description', 'Token generation failed') if token_data else 'Authentication method returned no data'
                    })
                    
            except Exception as e:
                token_results.append({
                    'success': False,
                    'scope_name': scope_key.title(),
                    'auth_method': auth_method,
                    'error': 'Request failed',
                    'error_description': str(e)
                })
        
        # Check if at least one token was generated successfully
        success_count = sum(1 for result in token_results if result['success'])
        
        return jsonify({
            'success': success_count > 0,
            'tokens': token_results,
            'summary': f"{success_count}/{len(token_results)} tokens generated successfully"
        })
            
    except Exception as e:
        return jsonify({
            'success': False,
            'error': 'Server error',
            'error_description': str(e)
        })

if __name__ == '__main__':
    print("🚀 Starting Multi-Tenant SharePoint Access Token Generator with Certificate Support...")
    print("📦 Required Dependencies: pip install flask requests cryptography")
    print("📝 Configuration Instructions:")
    print("   1. Update TENANT_CONFIG in this file with your actual tenant details")
    print("   2. Place certificate files (certificate.pem, private_key.pem) in the correct paths")
    print("   3. Ensure certificate files have proper permissions")
    print("\n🔧 Current Configuration:")
    for key, tenant in TENANT_CONFIG.items():
        print(f"   {key}: {tenant['name']}")
        print(f"      Tenant: {tenant['tenant_name']}")
        print(f"      App ID: {tenant['app_id']}")
        print(f"      Certificate: {tenant['certificate_path']}")
        print(f"      Private Key: {tenant['private_key_path']}")
        cert_exists = os.path.exists(tenant['certificate_path'])
        key_exists = os.path.exists(tenant['private_key_path'])
        print(f"      Cert Status: {'✅' if cert_exists else '❌'} Certificate, {'✅' if key_exists else '❌'} Private Key")
        print(f"      Scopes: {list(tenant['scopes'].keys())}")
        print()
    
    print(f"🌐 Server will start at: http://localhost:5000")
    print("💡 Press Ctrl+C to stop the server")
    
    app.run(debug=True, host='0.0.0.0', port=5555)