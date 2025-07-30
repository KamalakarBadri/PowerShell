import csv
import requests
import xml.etree.ElementTree as ET
import logging
import uuid
import base64
import time
import json
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Configuration
CONFIG = {
    # Authentication
    "tenant_id": "0e439a1f-a46b-4e582e203607",
    "client_id": "73efa35d-61-b258-838a977eb149",
    "client_secret": "sNxt5IejrMc2c24Ziz4a.t",
    "certificate_path": "certificate.pem",
    "private_key_path": "private_key.pem",
    
    # SharePoint
    "sharepoint_site": "https://TestT.sharepoint.com/sites/New365",
    "sharepoint_scope": "https://TestT.sharepoint.com/.default",
    
    # Graph
    "graph_scope": "https://graph.microsoft.com/.default",
    "graph_site_id": "your-site-id",
    "graph_list_id": "your-list-id"
}

def get_sharepoint_token():
    """Get SharePoint access token using JWT with private key/certificate"""
    try:
        # Load certificate and private key
        with open(CONFIG['certificate_path'], "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
        with open(CONFIG['private_key_path'], "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())

        # Prepare JWT claims
        now = int(time.time())
        jwt_header = {
            "alg": "RS256",
            "typ": "JWT",
            "x5t": base64.urlsafe_b64encode(certificate.fingerprint(hashes.SHA1())).decode().rstrip('=')
        }
        jwt_payload = {
            "aud": f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            "exp": now + 300,
            "iss": CONFIG['client_id'],
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": CONFIG['client_id']
        }

        # Encode and sign JWT
        encoded_header = base64.urlsafe_b64encode(json.dumps(jwt_header).encode()).decode().rstrip('=')
        encoded_payload = base64.urlsafe_b64encode(json.dumps(jwt_payload).encode()).decode().rstrip('=')
        jwt_unsigned = f"{encoded_header}.{encoded_payload}"
        signature = private_key.sign(jwt_unsigned.encode(), padding.PKCS1v15(), hashes.SHA256())
        encoded_signature = base64.urlsafe_b64encode(signature).decode().rstrip('=')
        jwt = f"{jwt_unsigned}.{encoded_signature}"

        # Request token
        token_response = requests.post(
            f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            data={
                "client_id": CONFIG['client_id'],
                "client_assertion": jwt,
                "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                "scope": CONFIG['sharepoint_scope'],
                "grant_type": "client_credentials"
            }
        )

        if token_response.status_code == 200:
            return token_response.json()["access_token"]
        else:
            logger.error(f"SharePoint token request failed: {token_response.text}")
            return None
            
    except Exception as e:
        logger.exception("SharePoint authentication failed")
        return None

def get_graph_token():
    """Get Graph API access token using client secret"""
    try:
        token_response = requests.post(
            f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            data={
                "client_id": CONFIG['client_id'],
                "client_secret": CONFIG['client_secret'],
                "scope": CONFIG['graph_scope'],
                "grant_type": "client_credentials"
            }
        )

        if token_response.status_code == 200:
            return token_response.json()["access_token"]
        else:
            logger.error(f"Graph token request failed: {token_response.text}")
            return None
            
    except Exception as e:
        logger.exception("Graph authentication failed")
        return None

def ensure_user(email, sp_token):
    """Ensure user exists in SharePoint using direct endpoint"""
    try:
        endpoint = f"{CONFIG['sharepoint_site']}/_api/web/ensureuser('{email}')"
        headers = {
            "Authorization": f"Bearer {sp_token}",
            "Accept": "application/xml",
            "Content-Type": "application/xml"
        }
        
        # First get request digest
        digest_response = requests.post(
            f"{CONFIG['sharepoint_site']}/_api/contextinfo",
            headers=headers
        )
        
        if digest_response.status_code != 200:
            logger.error(f"Failed to get request digest: {digest_response.text}")
            return None
            
        request_digest = digest_response.json()['d']['GetContextWebInformation']['FormDigestValue']
        headers['X-RequestDigest'] = request_digest
        
        # Call ensureuser
        response = requests.post(endpoint, headers=headers)
        
        if response.status_code == 200:
            # Parse XML response
            root = ET.fromstring(response.text)
            ns = {
                'd': 'http://schemas.microsoft.com/ado/2007/08/dataservices',
                'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'
            }
            return root.find('.//d:Id', ns).text
        else:
            logger.error(f"Failed to ensure user {email}: {response.text}")
            return None
            
    except Exception as e:
        logger.exception(f"Error ensuring user {email}")
        return None

def create_list_item(user_id, full_url_path, graph_token):
    """Create list item using Graph API"""
    try:
        endpoint = f"https://graph.microsoft.com/v1.0/sites/{CONFIG['graph_site_id']}/lists/{CONFIG['graph_list_id']}/items"
        headers = {
            "Authorization": f"Bearer {graph_token}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "fields": {
                "Emaillookup": int(user_id),
                "FULLURLPATH": full_url_path
            }
        }
        
        response = requests.post(endpoint, headers=headers, json=payload)
        if response.status_code == 201:
            logger.info(f"Created item for user ID {user_id}")
            return True
        else:
            logger.error(f"Failed to create item: {response.text}")
            return False
            
    except Exception as e:
        logger.exception("Error creating list item")
        return False

def process_csv(csv_path):
    """Process CSV file and upload data"""
    # Get tokens
    sp_token = get_sharepoint_token()
    graph_token = get_graph_token()
    
    if not sp_token or not graph_token:
        logger.error("Authentication failed")
        return False
    
    with open(csv_path, mode='r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            email = row['Email'].strip()
            path = row['FULLURLPATH'].strip()
            
            if not email or not path:
                logger.warning(f"Skipping row with missing data: {row}")
                continue
                
            user_id = ensure_user(f"i:0#.f|membership|{email}", sp_token)
            if user_id:
                if not create_list_item(user_id, path, graph_token):
                    logger.warning(f"Failed to create item for {email}")
            else:
                logger.warning(f"Failed to ensure user {email}")

    return True

if __name__ == '__main__':
    if process_csv("data.csv"):
        logger.info("CSV processing completed successfully")
    else:
        logger.error("CSV processing failed")
