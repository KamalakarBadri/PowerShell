from flask import Flask, render_template_string, request, jsonify
import requests
import json
import uuid
import base64
import time
import os
import logging
import xml.etree.ElementTree as ET
from datetime import datetime
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
from typing import List, Dict, Any, Optional, Tuple

app = Flask(__name__)
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

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
    },
    "max_workers": 5
}

group_owner_cache: Dict[str, Tuple[List[Dict], str]] = {}
cache_lock = Lock()
token_cache: Dict[str, Tuple[str, float]] = {}
token_lock = Lock()


# ============ AUTH ============
def get_token_with_certificate(scope: str) -> Optional[str]:
    try:
        if not os.path.exists(CONFIG['certificate_path']) or not os.path.exists(CONFIG['private_key_path']):
            return None
        with open(CONFIG['certificate_path'], "rb") as cf:
            cert = load_pem_x509_certificate(cf.read(), default_backend())
        with open(CONFIG['private_key_path'], "rb") as kf:
            key = load_pem_private_key(kf.read(), password=None, backend=default_backend())

        now = int(time.time())
        header = {"alg": "RS256", "typ": "JWT",
                  "x5t": base64.urlsafe_b64encode(cert.fingerprint(hashes.SHA1())).decode().rstrip('=')}
        payload = {"aud": f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
                   "exp": now + 300, "iss": CONFIG['app_id'], "jti": str(uuid.uuid4()),
                   "nbf": now, "sub": CONFIG['app_id']}
        eh = base64.urlsafe_b64encode(json.dumps(header).encode()).decode().rstrip('=')
        ep = base64.urlsafe_b64encode(json.dumps(payload).encode()).decode().rstrip('=')
        unsigned = f"{eh}.{ep}"
        sig = key.sign(unsigned.encode(), padding.PKCS1v15(), hashes.SHA256())
        es = base64.urlsafe_b64encode(sig).decode().rstrip('=')
        jwt = f"{unsigned}.{es}"

        r = requests.post(
            f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            data={"client_id": CONFIG['app_id'], "client_assertion": jwt,
                  "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
                  "scope": scope, "grant_type": "client_credentials"}, timeout=30)
        return r.json()["access_token"] if r.status_code == 200 else None
    except Exception:
        logger.exception("Cert auth failed")
        return None


def get_token_with_secret(scope: str) -> Optional[str]:
    try:
        r = requests.post(
            f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            data={"client_id": CONFIG['app_id'], "client_secret": CONFIG['client_secret'],
                  "scope": scope, "grant_type": "client_credentials"}, timeout=30)
        return r.json()["access_token"] if r.status_code == 200 else None
    except Exception:
        logger.exception("Secret auth failed")
        return None


def get_cached_token(scope_key: str) -> Optional[str]:
    scope = CONFIG['scopes'][scope_key]
    with token_lock:
        cached = token_cache.get(scope_key)
        if cached and cached[1] > time.time() + 60:
            return cached[0]
    tok = get_token_with_certificate(scope) or get_token_with_secret(scope)
    if tok:
        with token_lock:
            token_cache[scope_key] = (tok, time.time() + 3300)
    return tok


# ============ GRAPH / GROUP HELPERS ============
def extract_group_id_from_loginname(login_name: str) -> Optional[str]:
    if not login_name:
        return None
    try:
        parts = login_name.split('|')
        last = parts[-1]
        if last.endswith('_o'):
            last = last[:-2]
        return last.strip() or None
    except Exception:
        return None


def get_group_owners(group_id: str) -> Tuple[List[Dict[str, Any]], str]:
    with cache_lock:
        if group_id in group_owner_cache:
            return group_owner_cache[group_id]

    token = get_cached_token('graph')
    if not token:
        return [], group_id

    headers = {"Authorization": f"Bearer {token}"}
    group_title = group_id
    owners = []

    try:
        gr = requests.get(f"https://graph.microsoft.com/v1.0/groups/{group_id}", headers=headers, timeout=30)
        if gr.status_code == 200:
            group_title = gr.json().get('displayName', group_id)

        or_ = requests.get(f"https://graph.microsoft.com/v1.0/groups/{group_id}/owners", headers=headers, timeout=30)
        if or_.status_code == 200:
            for o in or_.json().get('value', []):
                owners.append({
                    'user_id': o.get('id'),
                    'title': o.get('displayName', ''),
                    'email': o.get('userPrincipalName', ''),
                    'login_name': o.get('userPrincipalName', ''),
                    'is_site_admin': True,
                    'is_group_member': True,
                    'group_id': group_id,
                    'group_title': group_title,
                    'principal_type': 1
                })
    except Exception:
        logger.exception(f"Failed owners for {group_id}")

    with cache_lock:
        group_owner_cache[group_id] = (owners, group_title)
    return owners, group_title


# ============ SHAREPOINT HELPERS ============
def normalize_site_url(s: str) -> str:
    s = s.strip()
    if not s:
        return s
    if s.startswith('http'):
        return s.rstrip('/')
    if not s.startswith('sites/'):
        s = f"sites/{s}"
    tenant = CONFIG['tenant_name'].split('.')[0]
    return f"https://{tenant}.sharepoint.com/{s}".rstrip('/')


def resolve_input_to_site_url(input_str: str) -> str:
    input_str = input_str.strip()
    if '@' in input_str and not input_str.startswith('http'):
        token = get_cached_token('graph')
        if not token:
            raise Exception("No Graph token for OneDrive lookup")
        r = requests.get(
            f"https://graph.microsoft.com/v1.0/users/{input_str}/drive?$select=webUrl",
            headers={"Authorization": f"Bearer {token}"}, timeout=30)
        if r.status_code != 200:
            raise Exception(f"OneDrive not found: {r.text}")
        url = r.json().get('webUrl', '')
        return url[:-10] if url.endswith('/Documents') else url.rstrip('/')
    return normalize_site_url(input_str)


def get_request_digest(site_url: str, token: str) -> str:
    if not site_url.endswith('/'):
        site_url += '/'
    r = requests.post(f"{site_url}_api/contextinfo",
                      headers={"Authorization": f"Bearer {token}",
                               "Accept": "application/json;odata=verbose"}, timeout=30)
    if r.status_code != 200:
        raise Exception(f"Digest failed: {r.text}")
    return r.json()['d']['GetContextWebInformation']['FormDigestValue']


def parse_site_users_xml(xml_content: str) -> Tuple[List[Dict], List[Dict], List[Dict]]:
    root = ET.fromstring(xml_content)
    ns = {'d': 'http://schemas.microsoft.com/ado/2007/08/dataservices',
          'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata',
          'atom': 'http://www.w3.org/2005/Atom'}

    direct_users, groups_to_expand = [], []

    for entry in root.findall('.//atom:entry', ns):
        content = entry.find('atom:content', ns)
        if content is None:
            continue
        props = content.find('m:properties', ns)
        if props is None:
            continue

        def get(tag):
            e = props.find(f'd:{tag}', ns)
            return e.text if e is not None else None

        user_id = get('Id')
        title = get('Title')
        email = get('Email')
        login = get('LoginName')
        is_admin = get('IsSiteAdmin') == 'true'
        ptype_raw = get('PrincipalType')
        ptype = int(ptype_raw) if ptype_raw else None

        if not is_admin:
            continue

        if ptype == 4:
            gid = extract_group_id_from_loginname(login) if login else None
            if gid:
                groups_to_expand.append({'group_id': gid, 'group_name': title})
        else:
            direct_users.append({
                'user_id': user_id, 'title': title, 'email': email,
                'login_name': login, 'is_site_admin': True,
                'principal_type': ptype, 'is_group_member': False,
                'group_id': None, 'group_title': None
            })

    group_owners, groups_info = [], []
    if groups_to_expand:
        with ThreadPoolExecutor(max_workers=min(CONFIG['max_workers'], max(1, len(groups_to_expand)))) as ex:
            futures = {ex.submit(get_group_owners, g['group_id']): g for g in groups_to_expand}
            for fut in as_completed(futures):
                g = futures[fut]
                try:
                    owners, gtitle = fut.result()
                except Exception:
                    owners, gtitle = [], g.get('group_name', '')
                display = gtitle if gtitle and gtitle != g['group_id'] else g.get('group_name', '')
                if owners:
                    for o in owners:
                        o['group_title'] = display
                    group_owners.extend(owners)
                else:
                    group_owners.append({
                        'user_id': g['group_id'], 'title': display, 'email': '',
                        'login_name': '', 'is_site_admin': True,
                        'is_group_member': False, 'is_group_itself': True,
                        'group_id': g['group_id'], 'group_title': display,
                        'principal_type': 4
                    })
                groups_info.append({
                    'group_id': g['group_id'], 'group_title': display,
                    'owner_count': len(owners)
                })

    return direct_users, group_owners, groups_info


def get_site_admins_detailed(site_url: str) -> Dict[str, Any]:
    if not site_url.endswith('/'):
        site_url += '/'
    token = get_cached_token('sharepoint')
    if not token:
        raise Exception("No SharePoint token")

    r = requests.get(f"{site_url}_api/web/siteusers",
                     headers={"Authorization": f"Bearer {token}", "Accept": "application/xml"}, timeout=60)
    if r.status_code != 200:
        raise Exception(f"SharePoint API error {r.status_code}: {r.text[:300]}")

    direct, owners, groups = parse_site_users_xml(r.text)
    return {
        'site_url': site_url.rstrip('/'),
        'direct_users': direct,
        'group_owners': owners,
        'groups': groups,
        'total_admins': len(direct) + len(owners)
    }


def load_one_site(input_str: str) -> Dict[str, Any]:
    """Load details for a single site — used by thread pool."""
    result = {'site_input': input_str, 'success': False, 'data': None, 'error': None}
    try:
        site_url = resolve_input_to_site_url(input_str)
        data = get_site_admins_detailed(site_url)
        result['success'] = True
        result['data'] = data
    except Exception as e:
        logger.exception(f"Load failed for {input_str}")
        result['error'] = str(e)
    return result


def ensure_user(site_url: str, token: str, digest: str, upn: str) -> Optional[str]:
    url = f"{site_url}/_api/web/ensureuser('{upn}')"
    try:
        r = requests.post(url, headers={
            "Authorization": f"Bearer {token}", "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose", "X-RequestDigest": digest
        }, timeout=30)
        if r.status_code == 200:
            return r.json()['d']['Id']
    except Exception:
        logger.exception(f"ensureuser failed for {upn}")
    return None


def set_site_admin_flag(site_url: str, token: str, digest: str, user_id: str, flag: bool) -> bool:
    url = f"{site_url}/_api/web/siteusers/getbyid({user_id})"
    try:
        r = requests.post(url, headers={
            "Authorization": f"Bearer {token}", "Accept": "application/json;odata=verbose",
            "Content-Type": "application/json;odata=verbose", "X-RequestDigest": digest,
            "IF-MATCH": "*", "X-HTTP-Method": "MERGE"
        }, json={"__metadata": {"type": "SP.User"}, "IsSiteAdmin": flag}, timeout=30)
        return r.status_code in (200, 204)
    except Exception:
        logger.exception(f"set_site_admin failed for {user_id}")
        return False


def find_user_id_on_site(site_url: str, token: str, upn: str) -> Optional[str]:
    """Resolve a user's site-user ID by UPN (via ensureuser — creates if missing)."""
    try:
        digest = get_request_digest(site_url, token)
        uid = ensure_user(site_url, token, digest, upn)
        return str(uid) if uid else None
    except Exception:
        logger.exception(f"find_user_id_on_site failed: {upn}")
        return None


def add_admin_to_site(site_url: str, upn: str) -> Dict[str, Any]:
    site_url = site_url.rstrip('/')
    token = get_cached_token('sharepoint')
    if not token:
        return {'success': False, 'site_url': site_url, 'upn': upn, 'error': 'No token'}
    try:
        digest = get_request_digest(site_url, token)
        uid = ensure_user(site_url, token, digest, upn)
        if not uid:
            return {'success': False, 'site_url': site_url, 'upn': upn, 'error': 'User not resolvable'}
        if not set_site_admin_flag(site_url, token, digest, uid, True):
            return {'success': False, 'site_url': site_url, 'upn': upn, 'error': 'Failed to set admin flag'}
        return {'success': True, 'site_url': site_url, 'upn': upn, 'user_id': uid}
    except Exception as e:
        return {'success': False, 'site_url': site_url, 'upn': upn, 'error': str(e)}


def remove_admin_from_site(site_url: str, identifier: str) -> Dict[str, Any]:
    """identifier can be either a site-user ID (numeric) or a UPN."""
    site_url = site_url.rstrip('/')
    token = get_cached_token('sharepoint')
    if not token:
        return {'success': False, 'site_url': site_url, 'identifier': identifier, 'error': 'No token'}
    try:
        digest = get_request_digest(site_url, token)
        if '@' in identifier:
            uid = ensure_user(site_url, token, digest, identifier)
            if not uid:
                return {'success': False, 'site_url': site_url, 'identifier': identifier,
                        'error': 'User not found on site'}
        else:
            uid = identifier
        if not set_site_admin_flag(site_url, token, digest, str(uid), False):
            return {'success': False, 'site_url': site_url, 'identifier': identifier,
                    'error': 'Failed to remove admin flag'}
        return {'success': True, 'site_url': site_url, 'identifier': identifier, 'user_id': str(uid)}
    except Exception as e:
        return {'success': False, 'site_url': site_url, 'identifier': identifier, 'error': str(e)}


# ============ HTML ============
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html>
<head>
<title>Site Admin Manager - Bulk Edition</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
:root{--p:#0078d4;--p2:#106ebe;--d:#d83b01;--s:#107c10;--w:#ffaa44;--g:#f3f2f1;--m:#e1dfdd;--dgr:#323130}
*{box-sizing:border-box}
body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;margin:0;padding:20px;background:#f9f9f9;color:var(--dgr)}
.container{max-width:1400px;margin:0 auto}
h1{color:var(--p);margin:0 0 20px;padding-bottom:10px;border-bottom:2px solid var(--m)}
h2{color:var(--dgr);font-size:1.2em;margin:0 0 10px}
.tabs{display:flex;gap:4px;margin-bottom:20px;border-bottom:2px solid var(--m);flex-wrap:wrap}
.tab{padding:10px 20px;cursor:pointer;background:transparent;border:none;font-size:15px;font-weight:600;color:#666;border-bottom:3px solid transparent;margin-bottom:-2px}
.tab.active{color:var(--p);border-bottom-color:var(--p)}
.tab-content{display:none}
.tab-content.active{display:block}
.card{background:#fff;padding:20px;border-radius:8px;box-shadow:0 2px 10px rgba(0,0,0,.08);margin-bottom:20px}
label{display:block;margin-bottom:6px;font-weight:600;font-size:14px}
textarea,input[type=text],input[type=email]{width:100%;padding:10px;border:1px solid var(--m);border-radius:4px;font-size:14px;font-family:inherit}
textarea{min-height:100px;resize:vertical}
.btn{padding:9px 18px;border:none;border-radius:4px;cursor:pointer;font-size:14px;font-weight:600;transition:background .2s}
.btn:disabled{opacity:.55;cursor:not-allowed}
.btn-primary{background:var(--p);color:#fff}.btn-primary:hover:not(:disabled){background:var(--p2)}
.btn-danger{background:var(--d);color:#fff}.btn-danger:hover:not(:disabled){background:#b32d00}
.btn-success{background:var(--s);color:#fff}.btn-success:hover:not(:disabled){background:#0c5c0c}
.btn-ghost{background:#fff;color:var(--p);border:1px solid var(--p)}
.btn-ghost:hover:not(:disabled){background:#f0f8ff}
.btn-sm{padding:6px 12px;font-size:13px}
.row{display:flex;gap:10px;flex-wrap:wrap;align-items:flex-end}
.row > div{flex:1;min-width:200px}
.status{padding:12px 16px;border-radius:4px;margin:15px 0;font-size:14px}
.status-success{background:#dff6dd;color:var(--s);border-left:4px solid var(--s)}
.status-error{background:#fde7e9;color:var(--d);border-left:4px solid var(--d)}
.status-info{background:#deecf9;color:var(--p);border-left:4px solid var(--p)}
.hidden{display:none}
.site-card{background:#fff;border:1px solid var(--m);border-radius:8px;margin-bottom:16px;overflow:hidden}
.site-head{background:var(--g);padding:12px 16px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid var(--m);gap:10px;flex-wrap:wrap}
.site-head h3{margin:0;font-size:15px;word-break:break-all}
.site-head .url{font-size:12px;color:#666;margin-top:2px}
.site-body{padding:0 16px 16px}
.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;margin-left:6px}
.badge-direct{background:#deecf9;color:var(--p)}
.badge-group{background:#fce8f3;color:#a4262c}
.section-title{font-weight:600;font-size:13px;color:#555;margin:14px 0 8px;text-transform:uppercase;letter-spacing:.5px}
.admin-row{display:flex;justify-content:space-between;align-items:center;padding:8px 10px;border-bottom:1px solid #f0f0f0;font-size:14px;gap:8px}
.admin-row:last-child{border-bottom:none}
.admin-info{flex:1;min-width:0}
.admin-name{font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.admin-sub{color:#666;font-size:12px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.group-tag{display:inline-block;background:#f0f0f0;color:#555;padding:1px 6px;border-radius:3px;font-size:11px;margin-left:6px}
.spinner{display:inline-block;width:16px;height:16px;border:2px solid rgba(255,255,255,.3);border-top-color:#fff;border-radius:50%;animation:spin .8s linear infinite;vertical-align:middle}
@keyframes spin{to{transform:rotate(360deg)}}
.pill{display:inline-block;background:var(--g);padding:4px 10px;border-radius:12px;font-size:12px;margin:2px}
.owners-table{width:100%;border-collapse:collapse;font-size:13px;margin-top:8px}
.owners-table th{text-align:left;padding:6px 8px;background:var(--g);font-weight:600;border-bottom:1px solid var(--m)}
.owners-table td{padding:6px 8px;border-bottom:1px solid #f0f0f0}
.bulk-result{margin-top:15px}
.bulk-result table{width:100%;border-collapse:collapse;font-size:13px}
.bulk-result th,.bulk-result td{padding:8px 10px;text-align:left;border-bottom:1px solid #eee}
.bulk-result th{background:var(--g);font-weight:600}
.ok{color:var(--s);font-weight:600}
.err{color:var(--d);font-weight:600}
</style>
</head>
<body>
<div class="container">
<h1>🛡️ Site Admin Manager — Bulk Edition</h1>

<div class="tabs">
  <button class="tab active" data-tab="load" onclick="switchTab('load',this)">📥 Load &amp; View</button>
  <button class="tab" data-tab="addbulk" onclick="switchTab('addbulk',this)">➕ Bulk Add Admins</button>
  <button class="tab" data-tab="removebulk" onclick="switchTab('removebulk',this)">➖ Bulk Remove Admins</button>
</div>

<!-- ========== TAB: LOAD ========== -->
<div id="tab-load" class="tab-content active">
  <div class="card">
    <label>Enter one site per line — supports Site URL, short name (e.g. <code>projectx</code>), or user UPN (resolves to OneDrive)</label>
    <textarea id="load-input" placeholder="https://geekbyteonline.sharepoint.com/sites/projectx&#10;marketing&#10;user@geekbyteonline.onmicrosoft.com"></textarea>
    <div class="row" style="margin-top:12px">
      <div style="flex:0"><button class="btn btn-primary" onclick="loadAll()" id="load-btn">Load All Sites</button></div>
      <div style="flex:0"><button class="btn btn-ghost" onclick="clearResults()">Clear Results</button></div>
    </div>
  </div>
  <div id="load-status"></div>
  <div id="load-results"></div>
</div>

<!-- ========== TAB: BULK ADD ========== -->
<div id="tab-addbulk" class="tab-content">
  <div class="card">
    <h2>Bulk Add Admins</h2>
    <p style="font-size:13px;color:#666">Paste multiple sites and multiple users. All users are added as admins to all sites.</p>
    <div class="row">
      <div><label>Sites (one per line)</label><textarea id="add-sites" placeholder="https://tenant.sharepoint.com/sites/a&#10;https://tenant.sharepoint.com/sites/b"></textarea></div>
      <div><label>Users to add (one UPN per line)</label><textarea id="add-users" placeholder="user1@tenant.com&#10;user2@tenant.com"></textarea></div>
    </div>
    <div style="margin-top:12px">
      <button class="btn btn-success" onclick="runBulkAdd()" id="bulk-add-btn">Add Users to All Sites</button>
    </div>
  </div>
  <div id="add-status"></div>
  <div id="add-results" class="bulk-result"></div>
</div>

<!-- ========== TAB: BULK REMOVE ========== -->
<div id="tab-removebulk" class="tab-content">
  <div class="card">
    <h2>Bulk Remove Admins</h2>
    <p style="font-size:13px;color:#666">Paste sites and users (UPNs or emails). Each user is removed from each site.</p>
    <div class="row">
      <div><label>Sites (one per line)</label><textarea id="rem-sites" placeholder="https://tenant.sharepoint.com/sites/a"></textarea></div>
      <div><label>Users to remove (UPNs or emails)</label><textarea id="rem-users" placeholder="user@tenant.com"></textarea></div>
    </div>
    <div style="margin-top:12px">
      <button class="btn btn-danger" onclick="runBulkRemove()" id="bulk-rem-btn">Remove Users from All Sites</button>
    </div>
  </div>
  <div id="rem-status"></div>
  <div id="rem-results" class="bulk-result"></div>
</div>

</div>

<script>
let loadedSites = {};

function switchTab(name, el) {
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
  el.classList.add('active');
  document.getElementById('tab-' + name).classList.add('active');
}

function status(id, msg, type) {
  const el = document.getElementById(id);
  el.innerHTML = msg ? `<div class="status status-${type}">${msg}</div>` : '';
}

function setBusy(btn, busy, text) {
  if (busy) {
    btn.dataset.orig = btn.textContent;
    btn.innerHTML = `<span class="spinner"></span> ${text || 'Working...'}`;
    btn.disabled = true;
  } else {
    btn.textContent = btn.dataset.orig || btn.textContent;
    btn.disabled = false;
  }
}

function esc(s) {
  if (s == null) return '';
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// ============ LOAD ============
async function loadAll() {
  const raw = document.getElementById('load-input').value.trim();
  if (!raw) { status('load-status', 'Please enter at least one site', 'error'); return; }
  const sites = raw.split('\n').map(s => s.trim()).filter(Boolean);
  const btn = document.getElementById('load-btn');
  setBusy(btn, true, 'Loading...');
  status('load-status', `Loading ${sites.length} site(s) in parallel...`, 'info');

  try {
    const res = await fetch('/api/load_sites', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ sites })
    });
    const data = await res.json();
    if (data.error) { status('load-status', data.error, 'error'); return; }

    const results = data.results || [];
    let ok = 0, fail = 0;
    results.forEach(r => {
      if (r.success) { ok++; loadedSites[r.data.site_url] = r.data; }
      else fail++;
    });

    renderResults(results);
    status('load-status', `✅ Loaded ${ok} site(s). ${fail ? '❌ ' + fail + ' failed.' : ''}`,
           fail ? 'info' : 'success');
  } catch (e) {
    status('load-status', 'Request failed: ' + e.message, 'error');
  } finally {
    setBusy(btn, false);
  }
}

function renderResults(results) {
  const container = document.getElementById('load-results');
  container.innerHTML = '';
  if (!results.length) return;

  results.forEach(r => {
    const card = document.createElement('div');
    card.className = 'site-card';
    if (!r.success) {
      card.innerHTML = `<div class="site-head"><div><h3>❌ ${esc(r.site_input || 'Unknown')}</h3><div class="url">${esc(r.error || '')}</div></div></div>`;
      container.appendChild(card);
      return;
    }

    const d = r.data;
    const direct = d.direct_users || [];
    const owners = d.group_owners || [];
    const groups = d.groups || [];

    const byGroup = {};
    owners.forEach(o => {
      const key = o.group_title || 'Unknown Group';
      if (!byGroup[key]) byGroup[key] = [];
      byGroup[key].push(o);
    });

    let groupsHtml = '';
    if (groups.length) {
      groupsHtml = `<div class="section-title">Groups with Admin Rights (${groups.length})</div>`;
      Object.keys(byGroup).forEach(gname => {
        const list = byGroup[gname];
        const gid = (list[0] && list[0].group_id) || '';
        groupsHtml += `
          <div style="margin:10px 0;padding:10px;background:#fafafa;border-radius:6px;border:1px solid #eee">
            <div style="font-weight:600;margin-bottom:6px">
              <span class="badge badge-group">GROUP</span> ${esc(gname)}
              <span class="group-tag">ID: ${esc(gid)}</span>
              <span class="group-tag">${list.length} owner(s)</span>
            </div>
            <table class="owners-table">
              <thead><tr><th>Owner Name</th><th>Email / UPN</th></tr></thead>
              <tbody>`;
        list.forEach(o => {
          const name = o.is_group_itself ? '(no owners resolved)' : (o.title || '—');
          const email = o.email || o.login_name || '';
          groupsHtml += `<tr><td>${esc(name)}</td><td>${esc(email)}</td></tr>`;
        });
        groupsHtml += `</tbody></table></div>`;
      });
    }

    let directHtml = '';
    if (direct.length) {
      directHtml = `<div class="section-title">Direct Site Administrators (${direct.length})</div>`;
      direct.forEach(u => {
        directHtml += `
          <div class="admin-row">
            <div class="admin-info">
              <div class="admin-name">${esc(u.title || u.email || u.login_name || '—')} <span class="badge badge-direct">DIRECT</span></div>
              <div class="admin-sub">${esc(u.email || u.login_name || '')}</div>
            </div>
            <button class="btn btn-danger btn-sm" data-site="${esc(d.site_url)}" data-id="${esc(u.user_id)}" data-label="${esc(u.email || u.login_name)}" onclick="removeOne(this)">Remove</button>
          </div>`;
      });
    } else {
      directHtml = `<div class="section-title">Direct Site Administrators</div><div style="padding:8px;color:#888;font-size:13px">None</div>`;
    }

    card.innerHTML = `
      <div class="site-head">
        <div>
          <h3>✅ ${esc(d.site_url)}</h3>
          <div class="url">Direct: ${direct.length} · Group Owners: ${owners.length} · Groups: ${groups.length} · Total: ${d.total_admins}</div>
        </div>
        <div><span class="pill" style="background:#deecf9">Total Admins: ${d.total_admins}</span></div>
      </div>
      <div class="site-body">${directHtml}${groupsHtml}</div>`;
    container.appendChild(card);
  });
}

function clearResults() {
  loadedSites = {};
  document.getElementById('load-results').innerHTML = '';
  document.getElementById('load-status').innerHTML = '';
}

async function removeOne(btn) {
  const siteUrl = btn.dataset.site;
  const userId = btn.dataset.id;
  const label = btn.dataset.label;
  if (!confirm(`Remove admin privileges for ${label}?`)) return;
  setBusy(btn, true, '...');
  try {
    const res = await fetch('/api/manage_admin', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ site_url: siteUrl, action: 'remove', user_id: userId })
    });
    const d = await res.json();
    if (d.error) { alert('Error: ' + d.error); return; }
    // Reload only this site
    const reload = await fetch('/api/load_sites', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ sites: [siteUrl] })
    });
    const rd = await reload.json();
    if (rd.results && rd.results[0] && rd.results[0].success) {
      loadedSites[rd.results[0].data.site_url] = rd.results[0].data;
      // Re-render only that card's data by replacing whole list (simple approach)
      const all = Object.values(loadedSites).map(x => ({success: true, data: x, site_input: x.site_url}));
      renderResults(all);
    }
  } catch (e) {
    alert('Request failed: ' + e.message);
  } finally {
    setBusy(btn, false);
  }
}

// ============ BULK ADD ============
async function runBulkAdd() {
  const sites = document.getElementById('add-sites').value.split('\n').map(s => s.trim()).filter(Boolean);
  const users = document.getElementById('add-users').value.split('\n').map(s => s.trim()).filter(Boolean);
  if (!sites.length || !users.length) {
    status('add-status', 'Please provide at least one site and one user', 'error');
    return;
  }
  const btn = document.getElementById('bulk-add-btn');
  setBusy(btn, true, 'Adding...');
  status('add-status', `Adding ${users.length} user(s) to ${sites.length} site(s)...`, 'info');

  try {
    const res = await fetch('/api/bulk_add', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ sites, users })
    });
    const data = await res.json();
    if (data.error) { status('add-status', data.error, 'error'); return; }
    renderBulkResults('add', data.results, data.summary);
    status('add-status', `✅ ${data.summary.success} succeeded · ❌ ${data.summary.failed} failed`,
           data.summary.failed ? 'info' : 'success');
  } catch (e) {
    status('add-status', 'Request failed: ' + e.message, 'error');
  } finally {
    setBusy(btn, false);
  }
}

// ============ BULK REMOVE ============
async function runBulkRemove() {
  const sites = document.getElementById('rem-sites').value.split('\n').map(s => s.trim()).filter(Boolean);
  const users = document.getElementById('rem-users').value.split('\n').map(s => s.trim()).filter(Boolean);
  if (!sites.length || !users.length) {
    status('rem-status', 'Please provide at least one site and one user', 'error');
    return;
  }
  if (!confirm(`Remove ${users.length} user(s) from ${sites.length} site(s)? This cannot be undone.`)) return;

  const btn = document.getElementById('bulk-rem-btn');
  setBusy(btn, true, 'Removing...');
  status('rem-status', `Removing ${users.length} user(s) from ${sites.length} site(s)...`, 'info');

  try {
    const res = await fetch('/api/bulk_remove', {
      method: 'POST', headers: {'Content-Type': 'application/json'},
      body: JSON.stringify({ sites, users })
    });
    const data = await res.json();
    if (data.error) { status('rem-status', data.error, 'error'); return; }
    renderBulkResults('rem', data.results, data.summary);
    status('rem-status', `✅ ${data.summary.success} succeeded · ❌ ${data.summary.failed} failed`,
           data.summary.failed ? 'info' : 'success');
  } catch (e) {
    status('rem-status', 'Request failed: ' + e.message, 'error');
  } finally {
    setBusy(btn, false);
  }
}

function renderBulkResults(prefix, results, summary) {
  const container = document.getElementById(prefix + '-results');
  if (!results || !results.length) { container.innerHTML = ''; return; }

  let html = `<div class="card" style="padding:0;overflow:hidden">
    <table>
      <thead><tr><th>Site</th><th>User</th><th>Status</th><th>Message</th></tr></thead>
      <tbody>`;
  results.forEach(r => {
    const ok = r.success;
    html += `<tr>
      <td style="font-size:12px;word-break:break-all">${esc(r.site_url || '')}</td>
      <td style="font-size:12px">${esc(r.upn || r.identifier || '')}</td>
      <td class="${ok ? 'ok' : 'err'}">${ok ? '✅ OK' : '❌ FAIL'}</td>
      <td style="font-size:12px;color:#666">${esc(ok ? ('user_id=' + (r.user_id || '?')) : (r.error || ''))}</td>
    </tr>`;
  });
  html += `</tbody></table></div>`;
  container.innerHTML = html;
}
</script>
</body>
</html>
"""


# ============ ROUTES ============
@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/api/load_sites', methods=['POST'])
def api_load_sites():
    try:
        data = request.json or {}
        sites = data.get('sites', [])
        if not isinstance(sites, list) or not sites:
            return jsonify({'error': 'No sites provided'}), 400

        results = []
        with ThreadPoolExecutor(max_workers=min(CONFIG['max_workers'], max(1, len(sites)))) as ex:
            futures = {ex.submit(load_one_site, s): s for s in sites}
            for fut in as_completed(futures):
                try:
                    results.append(fut.result())
                except Exception as e:
                    results.append({'site_input': futures[fut], 'success': False,
                                    'data': None, 'error': str(e)})

        # Preserve input order
        order = {s: i for i, s in enumerate(sites)}
        results.sort(key=lambda r: order.get(r.get('site_input', ''), 999))

        return jsonify({'results': results})
    except Exception as e:
        logger.exception("api_load_sites failed")
        return jsonify({'error': str(e)}), 500


@app.route('/api/manage_admin', methods=['POST'])
def api_manage_admin():
    try:
        data = request.json or {}
        site_url = (data.get('site_url') or '').strip()
        action = (data.get('action') or '').strip()
        upn = (data.get('upn') or '').strip()
        user_id = (data.get('user_id') or '').strip()

        if not site_url or action not in ('add', 'remove'):
            return jsonify({'error': 'Invalid parameters'}), 400

        if action == 'add':
            if not upn:
                return jsonify({'error': 'UPN required'}), 400
            r = add_admin_to_site(site_url, upn)
        else:
            if not user_id:
                return jsonify({'error': 'user_id required'}), 400
            r = remove_admin_from_site(site_url, user_id)

        if not r.get('success'):
            return jsonify({'error': r.get('error', 'Unknown error')}), 500
        return jsonify({'success': True, 'result': r})
    except Exception as e:
        logger.exception("api_manage_admin failed")
        return jsonify({'error': str(e)}), 500


@app.route('/api/bulk_add', methods=['POST'])
def api_bulk_add():
    try:
        data = request.json or {}
        sites = data.get('sites', [])
        users = data.get('users', [])
        if not sites or not users:
            return jsonify({'error': 'Sites and users required'}), 400

        # Resolve sites first (normalize / OneDrive)
        resolved = []
        for s in sites:
            try:
                resolved.append(resolve_input_to_site_url(s))
            except Exception as e:
                resolved.append({'error': str(e), 'input': s})

        tasks = []
        for s in resolved:
            if isinstance(s, dict):
                # failed resolution — record one failed row per user
                for u in users:
                    tasks.append((None, u, s.get('error')))
                continue
            for u in users:
                tasks.append((s, u, None))

        results = []

        def do_add(site_url, upn, pre_err):
            if pre_err:
                return {'success': False, 'site_url': site_url or '', 'upn': upn, 'error': pre_err}
            return add_admin_to_site(site_url, upn)

        with ThreadPoolExecutor(max_workers=CONFIG['max_workers']) as ex:
            futures = [ex.submit(do_add, s, u, e) for (s, u, e) in tasks]
            for fut in as_completed(futures):
                try:
                    results.append(fut.result())
                except Exception as e:
                    results.append({'success': False, 'error': str(e)})

        success = sum(1 for r in results if r.get('success'))
        return jsonify({
            'results': results,
            'summary': {'success': success, 'failed': len(results) - success, 'total': len(results)}
        })
    except Exception as e:
        logger.exception("api_bulk_add failed")
        return jsonify({'error': str(e)}), 500


@app.route('/api/bulk_remove', methods=['POST'])
def api_bulk_remove():
    try:
        data = request.json or {}
        sites = data.get('sites', [])
        users = data.get('users', [])
        if not sites or not users:
            return jsonify({'error': 'Sites and users required'}), 400

        resolved = []
        for s in sites:
            try:
                resolved.append(resolve_input_to_site_url(s))
            except Exception as e:
                resolved.append({'error': str(e), 'input': s})

        tasks = []
        for s in resolved:
            if isinstance(s, dict):
                for u in users:
                    tasks.append((None, u, s.get('error')))
                continue
            for u in users:
                tasks.append((s, u, None))

        results = []

        def do_remove(site_url, ident, pre_err):
            if pre_err:
                return {'success': False, 'site_url': site_url or '', 'identifier': ident, 'error': pre_err}
            return remove_admin_from_site(site_url, ident)

        with ThreadPoolExecutor(max_workers=CONFIG['max_workers']) as ex:
            futures = [ex.submit(do_remove, s, u, e) for (s, u, e) in tasks]
            for fut in as_completed(futures):
                try:
                    results.append(fut.result())
                except Exception as e:
                    results.append({'success': False, 'error': str(e)})

        success = sum(1 for r in results if r.get('success'))
        return jsonify({
            'results': results,
            'summary': {'success': success, 'failed': len(results) - success, 'total': len(results)}
        })
    except Exception as e:
        logger.exception("api_bulk_remove failed")
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5055, debug=False, threaded=True)