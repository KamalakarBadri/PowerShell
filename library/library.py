import requests
import json
import uuid

# Load tokens from the JSON file
with open('tokens.json', 'r') as f:
    tokens = json.load(f)

sharepoint_token = tokens['sharepoint_token']
site_url = "https://geekbyteonline.sharepoint.com/sites/Newwww2"

# Fixed SharePoint groups (modify these as needed)
SHAREPOINT_GROUPS = [
    "FAX Site Owners",
    "FAX Site Admins"
]

def get_request_digest(token, site_url):
    contextinfo_url = f"{site_url}/_api/contextinfo"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata=verbose"
    }
    response = requests.post(contextinfo_url, headers=headers)
    response.raise_for_status()
    return response.json()['d']['GetContextWebInformation']['FormDigestValue']

def ensure_principal(token, site_url, principal_identifier, is_group=False):
    endpoint = f"{site_url}/_api/web/ensureuser"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": get_request_digest(token, site_url)
    }
    
    if is_group and "@" not in principal_identifier and "|" not in principal_identifier:
        # Try both domain group formats
        for format_str in [principal_identifier, f"c:0t.c|tenant|{principal_identifier}"]:
            try:
                data = {'logonName': format_str}
                response = requests.post(endpoint, headers=headers, json=data)
                response.raise_for_status()
                return response.json()['d']
            except requests.exceptions.HTTPError:
                continue
        raise Exception(f"Could not resolve group: {principal_identifier}")
    else:
        login_name = principal_identifier if ("@" in principal_identifier or "|" in principal_identifier) else f"i:0#.f|membership|{principal_identifier}"
        data = {'logonName': login_name}
        response = requests.post(endpoint, headers=headers, json=data)
        response.raise_for_status()
        return response.json()['d']

def get_group_id(token, site_url, group_name):
    endpoint = f"{site_url}/_api/web/sitegroups/getbyname('{group_name}')"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata=verbose"
    }
    response = requests.get(endpoint, headers=headers)
    response.raise_for_status()
    return response.json()['d']['Id']

def add_role_assignment(token, site_url, list_id, principal_id, role_def_id):
    endpoint = f"{site_url}/_api/web/lists('{list_id}')/roleassignments/addroleassignment(principalid={principal_id},roledefid={role_def_id})"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": get_request_digest(token, site_url)
    }
    response = requests.post(endpoint, headers=headers)
    response.raise_for_status()
    return response.status_code

def add_user_to_library(token, site_url, list_id, user_email, role_def_id):
    try:
        user_principal = ensure_principal(token, site_url, user_email)
        add_role_assignment(token, site_url, list_id, user_principal['Id'], role_def_id)
        print(f"Added user {user_email} with access")
        return True
    except Exception as e:
        print(f"Could not add user {user_email}: {str(e)}")
        return False

def main():
    try:
        # Get project name from user
        project_name = input("Enter the project name: ").strip()
        library_name = f"FAX_{project_name}_Site"
        
        # Get request digest
        request_digest = get_request_digest(sharepoint_token, site_url)
        
        # Create headers
        headers = {
            "Authorization": f"Bearer {sharepoint_token}",
            "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose",
            "X-RequestDigest": request_digest
        }
        
        # 1. Create document library
        library_data = {
            "__metadata": {"type": "SP.List"},
            "AllowContentTypes": True,
            "BaseTemplate": 101,
            "ContentTypesEnabled": True,
            "Description": f"Document library for {project_name} project",
            "Title": library_name
        }
        
        create_response = requests.post(f"{site_url}/_api/web/lists", headers=headers, json=library_data)
        create_response.raise_for_status()
        list_id = create_response.json()['d']['Id']
        print(f"Created library '{library_name}' (ID: {list_id})")
        
        # 2. Break role inheritance
        break_response = requests.post(
            f"{site_url}/_api/web/lists('{list_id}')/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)",
            headers=headers
        )
        break_response.raise_for_status()
        print("Broke role inheritance successfully")
        
        # 3. Get role definitions
        roles_response = requests.get(
            f"{site_url}/_api/web/roledefinitions",
            headers={"Authorization": f"Bearer {sharepoint_token}", "Accept": "application/json;odata=verbose"}
        )
        roles = {r['Name']: r['Id'] for r in roles_response.json()['d']['results']}
        
        # 4. Add domain groups with Contribute access
        domain_groups = [
            f"FAX_{project_name}_AZG",
            f"FAX_{project_name}_ZZZ",
            f"DHCP Administrators"
        ]
        
        for group_name in domain_groups:
            try:
                group_principal = ensure_principal(sharepoint_token, site_url, group_name, is_group=True)
                add_role_assignment(sharepoint_token, site_url, list_id, group_principal['Id'], roles['Contribute'])
                print(f"Added domain group '{group_principal['Title']}' with Contribute access")
            except Exception as e:
                print(f"Could not add domain group {group_name}: {str(e)}")
        
        # 5. Add SharePoint groups with Full Control
        for group_name in SHAREPOINT_GROUPS:
            try:
                sp_group_id = get_group_id(sharepoint_token, site_url, group_name)
                add_role_assignment(sharepoint_token, site_url, list_id, sp_group_id, roles['Full Control'])
                print(f"Added SharePoint group '{group_name}' with Full Control access")
            except Exception as e:
                print(f"Could not add SharePoint group {group_name}: {str(e)}")
        
        # 6. Optional user provisioning
        users_input = input("Enter user emails to add (comma separated, or leave blank): ").strip()
        if users_input:
            user_emails = [email.strip() for email in users_input.split(",")]
            for email in user_emails:
                add_user_to_library(sharepoint_token, site_url, list_id, email, roles['Contribute'])
        
        print("\nDocument library setup complete!")
        print(f"Library Name: {library_name}")
        print(f"Library ID: {list_id}")
        
    except requests.exceptions.HTTPError as errh:
        print(f"HTTP Error: {str(errh)}")
        if hasattr(errh, 'response') and errh.response.text:
            print(f"Response: {errh.response.text}")
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()