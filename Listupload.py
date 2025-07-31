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
import os
from datetime import datetime

class SharePointCSVUploader:
    def __init__(self):
        # Configuration - Update these with your actual values
        self.config = {
            "sharepoint": {
                "site_url": "https://geekbyteonline.sharepoint.com/sites/New365",
                "tenant_id": "0e41f-a497-462b-9e6b-4e582e203607",
                "client_id": "73ef6188-42d4-b258-838a977eb149",
                "cert_path": "certificate.pem",
                "key_path": "private_key.pem",
                "resource": "https://geekbyteonline.sharepoint.com",
                "scope": "https://geekbyteonline.sharepoint.com/.default"
            },
            "graph": {
                "site_id": "e057-9d0d-4400-baa0-0ad7a1bf76dd",
                "list_id": "f9b0f4e4f-41fc-a6dc-7c9e8678219c",
                "client_id": "73ef88-42d4-b258-838a977eb149",
                "client_secret": "CYHuCMSyVmt4rMc2c24Ziz4a.t",
                "scope": "https://graph.microsoft.com/.default"
            }
        }

        # Token management
        self.sp_token = None
        self.graph_token = None
        self.token_expiry = 0

        # Logging setup
        self.log_file = f"upload_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        self._init_log_file()

    def _init_log_file(self):
        """Initialize the log CSV file with headers"""
        with open(self.log_file, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow([
                'Timestamp', 'Email', 'Status', 'UserID', 
                'ListItemID', 'Error', 'ActionTaken'
            ])

    def _log_entry(self, email, status, user_id=None, item_id=None, error=None, action=None):
        """Add an entry to the log CSV"""
        with open(self.log_file, mode='a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow([
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                email,
                status,
                user_id if user_id else 'N/A',
                item_id if item_id else 'N/A',
                error if error else 'N/A',
                action if action else 'N/A'
            ])

    def _generate_sp_token(self):
        """Generate SharePoint access token using certificate auth"""
        try:
            with open(self.config["sharepoint"]["cert_path"], "rb") as cert_file:
                cert = load_pem_x509_certificate(cert_file.read(), default_backend())
            with open(self.config["sharepoint"]["key_path"], "rb") as key_file:
                private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())

            now = int(time.time())
            header = {
                "alg": "RS256",
                "typ": "JWT",
                "x5t": base64.urlsafe_b64encode(cert.fingerprint(hashes.SHA1())).decode().rstrip('=')
            }
            payload = {
                "aud": f"https://login.microsoftonline.com/{self.config['sharepoint']['tenant_id']}/oauth2/v2.0/token",
                "exp": now + 3600,
                "iss": self.config["sharepoint"]["client_id"],
                "jti": str(uuid.uuid4()),
                "nbf": now,
                "sub": self.config["sharepoint"]["client_id"]
            }

            encoded_header = base64.urlsafe_b64encode(json.dumps(header).encode()).decode().rstrip('=')
            encoded_payload = base64.urlsafe_b64encode(json.dumps(payload).encode()).decode().rstrip('=')
            jwt = f"{encoded_header}.{encoded_payload}"
            signature = private_key.sign(jwt.encode(), padding.PKCS1v15(), hashes.SHA256())
            encoded_signature = base64.urlsafe_b64encode(signature).decode().rstrip('=')
            signed_jwt = f"{jwt}.{encoded_signature}"

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
        """Get user ID from SharePoint or return None if not found"""
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
                # User not found - we'll handle this gracefully
                return None

            # Parse XML response
            root = ET.fromstring(response.text)
            ns = {
                'd': 'http://schemas.microsoft.com/ado/2007/08/dataservices',
                'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'
            }
            return root.find('.//d:Id', ns).text
            
        except Exception as e:
            print(f"ERROR ensuring user {email}: {str(e)}")
            return None

    def create_list_item(self, email, row_data):
        """Create list item with fallback to EmailText when user not found"""
        try:
            if not self._refresh_tokens():
                return None

            # Try to get user ID first
            user_id = self.ensure_user(email)
            action_taken = "Used EmaillookupID" if user_id else "Used EmailText fallback"
            
            # Prepare payload based on whether user was found
            if user_id:
                payload = {
                    "fields": {
                        "EmailLookupId": int(user_id),
                        **{k: v for k, v in row_data.items() if k.lower() != "email"}
                    }
                }
            else:
                payload = {
                    "fields": {
                        "EmailText": email,
                        **{k: v for k, v in row_data.items() if k.lower() != "email"}
                    }
                }

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
                error_msg = f"Create item failed: {response.status_code} - {response.text}"
                self._log_entry(email, "Failed", user_id, None, error_msg, action_taken)
                return None

            item_id = response.json().get('id', 'Unknown')
            self._log_entry(email, "Success", user_id, item_id, None, action_taken)
            return item_id

        except Exception as e:
            error_msg = str(e)
            print(f"ERROR creating list item: {error_msg}")
            self._log_entry(email, "Failed", user_id, None, error_msg, action_taken)
            return None

    def process_csv(self, file_path):
        """Process CSV file and upload to SharePoint list with fallback handling"""
        try:
            print(f"Starting CSV processing: {file_path}")
            print(f"Log file will be created at: {os.path.abspath(self.log_file)}")
            
            if not self._refresh_tokens():
                raise Exception("Failed to initialize authentication")

            with open(file_path, mode='r', encoding='utf-8-sig') as csv_file:
                reader = csv.DictReader(csv_file)
                total_rows = 0
                success_count = 0
                fail_count = 0

                for row_num, row in enumerate(reader, 1):
                    total_rows += 1
                    email = row.get('Email', '').strip()
                    if not email:
                        print(f"Row {row_num}: Skipping - missing email")
                        self._log_entry("N/A", "Skipped", None, None, "Missing email", "Skipped row")
                        continue

                    print(f"Processing row {row_num}: {email}")
                    
                    # Create a copy of row without Email field
                    item_data = {k: v for k, v in row.items() if k.lower() != "email"}
                    
                    item_id = self.create_list_item(email, item_data)
                    if item_id:
                        success_count += 1
                        print(f"Row {row_num}: Successfully created item {item_id}")
                    else:
                        fail_count += 1
                        print(f"Row {row_num}: Failed to create list item")

            print("\nProcessing complete:")
            print(f"Total rows processed: {total_rows}")
            print(f"Successfully created: {success_count}")
            print(f"Failed to create: {fail_count}")
            print(f"Detailed log saved to: {os.path.abspath(self.log_file)}")
            
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
