import requests
import os

# ===== CONFIGURATION =====
TENANT_ID = "0e439a1f-a497-462b-9e6b-4e582e203607"
CLIENT_ID = "73efa35d-6188-42d4-b258-838a977eb149"
CLIENT_SECRET = "CyG8Q~FYHuCMSyVmt4sNxt5IejrMc2c24Ziz4a.t"

# SharePoint site where you want to upload
SITE_URL = "https://geekbyteonline.sharepoint.com/sites/2DayRetention"
DOCUMENT_LIBRARY = "Site Analytics"  # The library/folder name

# File to upload
FILE_PATH = "C:/path/to/your/file.csv"  # Change this to your file path
# =========================

def get_access_token():
    """Get access token"""
    auth_url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    body = {
        "client_id": CLIENT_ID,
        "scope": "https://graph.microsoft.com/.default",
        "client_secret": CLIENT_SECRET,
        "grant_type": "client_credentials"
    }
    
    response = requests.post(auth_url, data=body)
    response.raise_for_status()
    return response.json().get("access_token")

def get_site_id(token, site_url):
    """Get SharePoint site ID"""
    parsed = requests.utils.urlparse(site_url)
    hostname = parsed.netloc
    site_path = parsed.path.rstrip('/')
    
    url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"
    headers = {"Authorization": f"Bearer {token}"}
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json().get("id")

def get_drive_id(token, site_id, library_name):
    """Get drive ID for the document library"""
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    headers = {"Authorization": f"Bearer {token}"}
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    
    drives = response.json().get("value", [])
    for drive in drives:
        if drive.get("name") == library_name:
            return drive.get("id")
    
    raise Exception(f"Library '{library_name}' not found")

def upload_file(token, site_id, drive_id, file_path):
    """Upload file to SharePoint"""
    file_name = os.path.basename(file_path)
    upload_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/root:/{file_name}:/content"
    
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/octet-stream"
    }
    
    with open(file_path, 'rb') as file:
        response = requests.put(upload_url, headers=headers, data=file.read())
    
    response.raise_for_status()
    return response.json()

def main():
    try:
        print("📤 Getting access token...")
        token = get_access_token()
        
        print(f"📍 Getting site info for: {SITE_URL}")
        site_id = get_site_id(token, SITE_URL)
        
        print(f"📁 Getting drive ID for: {DOCUMENT_LIBRARY}")
        drive_id = get_drive_id(token, site_id, DOCUMENT_LIBRARY)
        
        print(f"⬆️ Uploading {FILE_PATH}...")
        result = upload_file(token, site_id, drive_id, FILE_PATH)
        
        print("✅ File uploaded successfully!")
        print(f"   File name: {result.get('name')}")
        print(f"   File size: {result.get('size')} bytes")
        print(f"   Web URL: {result.get('webUrl')}")
        
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    main()
