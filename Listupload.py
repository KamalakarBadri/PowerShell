import csv
import requests
import time
import xml.etree.ElementTree as ET
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend
import base64
import json
import uuid

class SharePointCSVUploader:
    def __init__(self):
        # Configuration - Update these with your actual values
        self.config = {
            "sharepoint": {
                "site_url": "https://geekbyteonline.sharepoint.com/sites/New365",
                "tenant_id": "0e439a1f-a497-462b-9e6b-4e582e203607",
                "client_id": "73e-838a977eb149",
                "cert_path": "certificate.pem",
                "key_path": "private_key.pem",
                "resource": "https://geekbyteonline.sharepoint.com",
                "scope": "https://geekbyteonline.sharepoint.com/.default"
            },
            "graph": {
                "site_id": "e0578a400-baa0-0ad7a1bf76dd",
                "list_id": "f9b0f9-a6dc-7c9e8678219c",
                "client_id": "73efa38-42d4-b258-838a977eb149",
                "client_secret": "CyG8Q4sNxt5IejrMc2c24Ziz4a.t",
                "scope": "https://graph.microsoft.com/.default"
            }
        }

        # Token management
        self.sp_token = None
        self.graph_token = None
        self.token_expiry = 0

    def _generate_sp_token(self):
        """Generate SharePoint access token using certificate auth"""
        try:
            # Load certificate and private key
            with open(self.config["sharepoint"]["cert_path"], "rb") as cert_file:
                cert = load_pem_x509_certificate(cert_file.read(), default_backend())
            with open(self.config["sharepoint"]["key_path"], "rb") as key_file:
                private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())

            # Prepare JWT claims
            now = int(time.time())
            header = {
                "alg": "RS256",
                "typ": "JWT",
                "x5t": base64.urlsafe_b64encode(cert.fingerprint(hashes.SHA1())).decode().rstrip('=')
            }
            payload = {
                "aud": f"https://login.microsoftonline.com/{self.config['sharepoint']['tenant_id']}/oauth2/v2.0/token",
                "exp": now + 3600,  # 1 hour expiration
                "iss": self.config["sharepoint"]["client_id"],
                "jti": str(uuid.uuid4()),
                "nbf": now,
                "sub": self.config["sharepoint"]["client_id"]
            }

            # Encode and sign JWT
            encoded_header = base64.urlsafe_b64encode(json.dumps(header).encode()).decode().rstrip('=')
            encoded_payload = base64.urlsafe_b64encode(json.dumps(payload).encode()).decode().rstrip('=')
            jwt = f"{encoded_header}.{encoded_payload}"
            signature = private_key.sign(jwt.encode(), padding.PKCS1v15(), hashes.SHA256())
            encoded_signature = base64.urlsafe_b64encode(signature).decode().rstrip('=')
            signed_jwt = f"{jwt}.{encoded_signature}"

            # Request token
            response = requests.post(
                f"https://login.microsoftonline.com/{self.config['sharepoint']['tenant_id']}/oauth2/v2.0/token",
                data={
                    "client_id": self.config["sharepoint"]["client_id"],
                    "client_assertion": signed_jwt,
                    "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                    "scope": self.config["sharepoint"]["scope"],
                    "grant_type": "client_credentials"
                },
                timeout=30
            )

            if response.status_code != 200:
                raise Exception(f"Token request failed: {response.status_code} - {response.text}")

            return response.json()["access_token"]
        except Exception as e:
            print(f"ERROR generating SharePoint token: {str(e)}")
            return None

    def _generate_graph_token(self):
        """Generate Graph API token using client secret"""
        try:
            response = requests.post(
                f"https://login.microsoftonline.com/{self.config['sharepoint']['tenant_id']}/oauth2/v2.0/token",
                data={
                    "client_id": self.config["graph"]["client_id"],
                    "client_secret": self.config["graph"]["client_secret"],
                    "scope": self.config["graph"]["scope"],
                    "grant_type": "client_credentials"
                },
                timeout=30
            )

            if response.status_code != 200:
                raise Exception(f"Graph token request failed: {response.status_code} - {response.text}")

            return response.json()["access_token"]
        except Exception as e:
            print(f"ERROR generating Graph token: {str(e)}")
            return None

    def _refresh_tokens(self):
        """Refresh tokens if they're expired or about to expire"""
        current_time = time.time()
        if current_time > self.token_expiry - 900:  # Refresh if less than 15 minutes remaining
            new_sp_token = self._generate_sp_token()
            new_graph_token = self._generate_graph_token()
            
            if new_sp_token and new_graph_token:
                self.sp_token = new_sp_token
                self.graph_token = new_graph_token
                self.token_expiry = current_time + 2700  # 45 minutes from now
                print("Tokens refreshed successfully")
                return True
            else:
                raise Exception("Failed to refresh one or more tokens")
        return True

    def ensure_user(self, email):
        """Get user ID from SharePoint using exact ensureuser endpoint"""
        try:
            if not self._refresh_tokens():
                return None

            endpoint = f"{self.config['sharepoint']['site_url']}/_api/web/ensureuser('{email}')"
            headers = {
                "Authorization": f"Bearer {self.sp_token}",
                "Accept": "application/xml",
            }
            
            response = requests.post(endpoint, headers=headers, timeout=30)
            
            if response.status_code != 200:
                raise Exception(f"EnsureUser failed: {response.status_code} - {response.text}")

            # Parse XML response
            root = ET.fromstring(response.text)
            ns = {
                'd': 'http://schemas.microsoft.com/ado/2007/08/dataservices',
                'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'
            }
            user_id = root.find('.//d:Id', ns).text
            return user_id
            
        except Exception as e:
            print(f"ERROR ensuring user {email}: {str(e)}")
            return None

    def create_list_item(self, user_id, row_data):
        """Create list item using Graph API with exact specified format"""
        try:
            if not self._refresh_tokens():
                return False

            # Prepare payload with EmaillookupID and all other fields
            payload = {
                "fields": {
                    "EmailLookupId": user_id
                }
            }
            
            # Add all other fields from CSV
            for field_name, field_value in row_data.items():
                if field_name.lower() != "email":
                    payload["fields"][field_name] = field_value

            response = requests.post(
                f"https://graph.microsoft.com/v1.0/sites/{self.config['graph']['site_id']}/lists/{self.config['graph']['list_id']}/items",
                headers={
                    "Authorization": f"Bearer {self.graph_token}",
                    "Content-Type": "application/json"
                },
                json=payload,
                timeout=30
            )

            if response.status_code != 201:
                raise Exception(f"Create item failed: {response.status_code} - {response.text}")

            return True
        except Exception as e:
            print(f"ERROR creating list item: {str(e)}")
            return False

    def process_csv(self, file_path):
        """Process CSV file and upload to SharePoint list"""
        try:
            print(f"Starting CSV processing: {file_path}")
            if not self._refresh_tokens():
                raise Exception("Failed to initialize authentication")

            with open(file_path, mode='r', encoding='utf-8-sig') as csv_file:
                reader = csv.DictReader(csv_file)
                for row_num, row in enumerate(reader, 1):
                    email = row.get('Email', '').strip()
                    if not email:
                        print(f"Row {row_num}: Skipping - missing email")
                        continue

                    print(f"Processing row {row_num}: {email}")
                    user_id = self.ensure_user(email)
                    if not user_id:
                        print(f"Row {row_num}: Failed to get user ID")
                        continue

                    # Create a copy of row without Email field
                    item_data = {k: v for k, v in row.items() if k.lower() != "email"}
                    
                    if not self.create_list_item(user_id, item_data):
                        print(f"Row {row_num}: Failed to create list item")
                        continue

                    print(f"Row {row_num}: Successfully processed")

            print("CSV processing completed successfully")
            return True

        except Exception as e:
            print(f"FATAL ERROR processing CSV: {str(e)}")
            return False

if __name__ == "__main__":
    print("Starting SharePoint CSV Uploader")
    uploader = SharePointCSVUploader()
    success = uploader.process_csv("list.csv")
    
    if success:
        print("Upload completed successfully")
    else:
        print("Upload failed with errors")
    print("Script finished")
