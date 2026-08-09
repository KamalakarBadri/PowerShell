import requests
import json
import csv
import uuid
import base64
import time
import os
from datetime import datetime
import re
import psutil
import gc
import threading
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives.serialization import load_pem_private_key
from cryptography.x509 import load_pem_x509_certificate
from cryptography.hazmat.backends import default_backend

# ============================================================
# LOAD CONFIGURATION
# ============================================================

def load_config():
    """Load configuration from JSON file"""
    config_file = "config_version.json"
    
    if not os.path.exists(config_file):
        print(f"❌ Config file '{config_file}' not found!")
        print("📌 Please create config_version.json file")
        sys.exit(1)
    
    with open(config_file, 'r') as f:
        config = json.load(f)
    
    print(f"✅ Configuration loaded from {config_file}")
    return config

CONFIG = load_config()

# ============================================================
# PERFORMANCE ANALYZER CLASS
# ============================================================

class PerformanceAnalyzer:
    def __init__(self, config_name, batch_size, max_workers):
        self.config_name = config_name
        self.batch_size = batch_size
        self.max_workers = max_workers
        self.start_time = None
        self.end_time = None
        self.total_items = 0
        self.processed_items = 0
        self.batch_count = 0
        self.worker_count = 0
        self.api_calls = 0
        self.retry_count = 0
        self.rate_limit_errors = 0
        self.other_errors = 0
        self.memory_usage = []
        self.batch_times = []
        self.file_times = []
        self.start_memory = 0
        self.peak_memory = 0
        self.current_batch_size = 0
        self.batch_stats = []
        
    def start(self):
        """Start performance monitoring"""
        self.start_time = time.time()
        self.start_memory = psutil.Process().memory_info().rss / 1024 / 1024
        self.peak_memory = self.start_memory
        print(f"\n{'='*80}")
        print(f"🔬 CONFIG: {self.config_name}")
        print(f"   Batch Size: {self.batch_size}, Max Workers: {self.max_workers}")
        print(f"   Start Memory: {self.start_memory:.2f} MB")
        print(f"{'='*80}")
        
    def log_memory(self, label=""):
        """Log current memory usage"""
        current_memory = psutil.Process().memory_info().rss / 1024 / 1024
        self.memory_usage.append(current_memory)
        if current_memory > self.peak_memory:
            self.peak_memory = current_memory
        return current_memory
    
    def log_batch_start(self, batch_num, items_in_batch):
        """Log batch start"""
        self.batch_count = batch_num
        self.current_batch_size = items_in_batch
        self.batch_start_time = time.time()
        memory = self.log_memory(f"Batch {batch_num} Start")
        print(f"\n  📦 Batch {batch_num}: {items_in_batch} files, Memory: {memory:.1f} MB")
        
    def log_batch_end(self):
        """Log batch end"""
        batch_time = time.time() - self.batch_start_time
        self.batch_times.append(batch_time)
        memory = self.log_memory(f"Batch {self.batch_count} End")
        print(f"  ✅ Batch {self.batch_count} done: {batch_time:.2f}s, Memory: {memory:.1f} MB")
        
    def log_file_processed(self):
        """Log file processing"""
        self.processed_items += 1
        
    def log_api_call(self):
        """Log API call"""
        self.api_calls += 1
        
    def log_rate_limit_error(self):
        """Log 429 error"""
        self.rate_limit_errors += 1
        
    def log_retry(self):
        """Log retry"""
        self.retry_count += 1
        
    def log_other_error(self):
        """Log other error"""
        self.other_errors += 1
        
    def finish(self):
        """Finish monitoring and return results"""
        self.end_time = time.time()
        self.total_items = self.processed_items
        total_time = self.end_time - self.start_time
        
        avg_batch_time = sum(self.batch_times) / len(self.batch_times) if self.batch_times else 0
        avg_file_time = total_time / self.processed_items if self.processed_items > 0 else 0
        avg_memory = sum(self.memory_usage) / len(self.memory_usage) if self.memory_usage else 0
        
        results = {
            'config': self.config_name,
            'batch_size': self.batch_size,
            'max_workers': self.max_workers,
            'total_items': self.total_items,
            'total_time': total_time,
            'avg_batch_time': avg_batch_time,
            'avg_file_time': avg_file_time,
            'start_memory': self.start_memory,
            'peak_memory': self.peak_memory,
            'avg_memory': avg_memory,
            'end_memory': self.log_memory("End"),
            'batch_count': self.batch_count,
            'api_calls': self.api_calls,
            'retry_count': self.retry_count,
            'rate_limit_errors': self.rate_limit_errors,
            'other_errors': self.other_errors,
            'throughput': self.processed_items / total_time if total_time > 0 else 0,
            'batches': self.batch_stats,
            'memory_usage': self.memory_usage
        }
        
        self.print_summary(results)
        return results
    
    def print_summary(self, results):
        """Print performance summary"""
        print(f"\n{'='*80}")
        print(f"📊 SUMMARY: {self.config_name}")
        print(f"{'='*80}")
        print(f"  📄 ITEMS:")
        print(f"    Total items: {results['total_items']}")
        print(f"    Batches: {results['batch_count']}")
        print(f"    Avg batch size: {results['total_items']/results['batch_count']:.1f}" if results['batch_count'] > 0 else "N/A")
        
        print(f"\n  ⏱️  TIME:")
        print(f"    Total time: {results['total_time']:.2f}s")
        print(f"    Avg batch time: {results['avg_batch_time']:.2f}s")
        print(f"    Avg file time: {results['avg_file_time']:.3f}s")
        print(f"    Throughput: {results['throughput']:.2f} files/sec")
        
        print(f"\n  💾 MEMORY:")
        print(f"    Start: {results['start_memory']:.1f} MB")
        print(f"    Peak: {results['peak_memory']:.1f} MB")
        print(f"    Avg: {results['avg_memory']:.1f} MB")
        print(f"    End: {results['end_memory']:.1f} MB")
        print(f"    Max spike: {results['peak_memory'] - results['start_memory']:.1f} MB")
        
        print(f"\n  🚀 API:")
        print(f"    API calls: {results['api_calls']}")
        print(f"    Retries: {results['retry_count']}")
        print(f"    429 errors: {results['rate_limit_errors']}")
        print(f"    Other errors: {results['other_errors']}")
        
        print(f"\n  📈 PERFORMANCE:")
        if results['throughput'] > 10:
            print(f"    ✅ FAST: {results['throughput']:.2f} files/sec")
        elif results['throughput'] > 5:
            print(f"    ⚠️ MEDIUM: {results['throughput']:.2f} files/sec")
        else:
            print(f"    ❌ SLOW: {results['throughput']:.2f} files/sec")
            
        if results['rate_limit_errors'] > 0:
            print(f"    ⚠️ Rate limit hit {results['rate_limit_errors']} times!")
            
        if results['peak_memory'] > 2048:
            print(f"    ❌ Memory too high: {results['peak_memory']:.1f} MB!")
            
        print(f"{'='*80}")

# ============================================================
# GLOBAL VARIABLES
# ============================================================

TOKEN_CACHE = {"token": None, "expires": 0}
ALLOWED_FILE_EXTENSIONS = None
global_analyzer = None
stats_lock = threading.Lock()

# ============================================================
# AUTHENTICATION FUNCTIONS
# ============================================================

def load_certificate_and_key():
    try:
        if not os.path.exists(CONFIG['certificate_path']) or not os.path.exists(CONFIG['private_key_path']):
            raise Exception(f"Certificate files not found.")
        
        with open(CONFIG['certificate_path'], "rb") as cert_file:
            certificate = load_pem_x509_certificate(cert_file.read(), default_backend())
        
        with open(CONFIG['private_key_path'], "rb") as key_file:
            private_key = load_pem_private_key(key_file.read(), password=None, backend=default_backend())
        
        return certificate, private_key
    except Exception as e:
        print(f"Error loading certificate: {str(e)}")
        raise

def get_jwt_token(certificate, private_key):
    try:
        now = int(time.time())
        expiration = now + 300
        
        thumbprint = certificate.fingerprint(hashes.SHA1())
        x5t = base64.urlsafe_b64encode(thumbprint).decode('utf-8').replace('=', '')
        
        jwt_header = {"alg": "RS256", "typ": "JWT", "x5t": x5t}
        
        jwt_payload = {
            "aud": f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token",
            "exp": expiration,
            "iss": CONFIG['app_id'],
            "jti": str(uuid.uuid4()),
            "nbf": now,
            "sub": CONFIG['app_id']
        }
        
        encoded_header = base64.urlsafe_b64encode(
            json.dumps(jwt_header, separators=(',', ':')).encode('utf-8')
        ).decode('utf-8').replace('=', '')
        
        encoded_payload = base64.urlsafe_b64encode(
            json.dumps(jwt_payload, separators=(',', ':')).encode('utf-8')
        ).decode('utf-8').replace('=', '')
        
        jwt_unsigned = f"{encoded_header}.{encoded_payload}"
        
        signature = private_key.sign(
            jwt_unsigned.encode('utf-8'),
            padding.PKCS1v15(),
            hashes.SHA256()
        )
        encoded_signature = base64.urlsafe_b64encode(signature).decode('utf-8').replace('=', '')
        
        return f"{jwt_unsigned}.{encoded_signature}"
    except Exception as e:
        print(f"Error generating JWT: {str(e)}")
        raise

def get_access_token(jwt):
    print("  🔑 Getting access token...")
    url = f"https://login.microsoftonline.com/{CONFIG['tenant_id']}/oauth2/v2.0/token"
    
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    
    data = {
        "client_id": CONFIG['app_id'],
        "client_assertion": jwt,
        "client_assertion_type": "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
        "scope": CONFIG['scope'],
        "grant_type": "client_credentials"
    }
    
    try:
        response = requests.post(url, headers=headers, data=data, timeout=CONFIG['version_settings']['request_timeout'])
        response.raise_for_status()
        result = response.json()
        print("  ✅ Token obtained")
        return result["access_token"]
    except Exception as e:
        print(f"  ❌ Error: {str(e)}")
        raise

def get_cached_token(force_refresh=False):
    cache = TOKEN_CACHE
    
    if not force_refresh and cache["token"] and cache["expires"] > time.time() + 300:
        return cache["token"]
    
    try:
        certificate, private_key = load_certificate_and_key()
        jwt = get_jwt_token(certificate, private_key)
        token = get_access_token(jwt)
        
        if token:
            cache["token"] = token
            cache["expires"] = time.time() + 3600
            return token
        return None
    except Exception as e:
        print(f"  ❌ Authentication failed: {str(e)}")
        return None

def get_current_token():
    return get_cached_token()

def make_sharepoint_request(site_url, url, max_retries=3):
    """Make request with retry logic"""
    global global_analyzer
    
    for attempt in range(max_retries + 1):
        try:
            token = get_current_token()
            if not token:
                return None
            
            headers = {
                "Authorization": f"Bearer {token}",
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json"
            }
            
            response = requests.get(url, headers=headers, timeout=CONFIG['version_settings']['request_timeout'])
            
            if global_analyzer:
                global_analyzer.log_api_call()
            
            if response.status_code == 429:
                if global_analyzer:
                    global_analyzer.log_rate_limit_error()
                
                wait_time = 2 ** (attempt + 1)
                print(f"    ⚠️ 429 Too Many Requests! Waiting {wait_time}s...")
                time.sleep(wait_time)
                
                TOKEN_CACHE["token"] = None
                TOKEN_CACHE["expires"] = 0
                continue
            
            if response.status_code == 401 and attempt < max_retries:
                print(f"    ⚠️ Token expired, refreshing...")
                TOKEN_CACHE["token"] = None
                TOKEN_CACHE["expires"] = 0
                if global_analyzer:
                    global_analyzer.log_retry()
                continue
            
            response.raise_for_status()
            return response.json()
            
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 429:
                continue
            if attempt < max_retries:
                time.sleep(2 ** attempt)
                if global_analyzer:
                    global_analyzer.log_retry()
                continue
            if global_analyzer:
                global_analyzer.log_other_error()
            return None
        except Exception as e:
            if attempt < max_retries:
                time.sleep(2 ** attempt)
                if global_analyzer:
                    global_analyzer.log_retry()
                continue
            if global_analyzer:
                global_analyzer.log_other_error()
            return None
    
    return None

# ============================================================
# UTILITY FUNCTIONS
# ============================================================

def get_site_prefix(site_url):
    normalized = site_url.rstrip('/')
    parts = normalized.split('/')
    if 'sites' in parts:
        idx = parts.index('sites')
        if idx + 1 < len(parts):
            return parts[idx + 1]
    if parts:
        return parts[-1]
    return 'Site'

def normalize_extensions(extensions):
    if not extensions:
        return None
    if isinstance(extensions, str):
        extensions = [extensions]
    normalized = []
    for ext in extensions:
        if not ext:
            continue
        ext = ext.lower().strip()
        if ext.startswith('.'):
            ext = ext[1:]
        if ext:
            normalized.append(ext)
    return normalized or None

def should_process_file(file_name):
    if not ALLOWED_FILE_EXTENSIONS:
        return True
    _, ext = os.path.splitext(file_name or '')
    ext = ext.lower().lstrip('.')
    return ext in ALLOWED_FILE_EXTENSIONS

def safe_int_conversion(value):
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    if isinstance(value, str):
        cleaned = re.sub(r'[^\d.]', '', value)
        try:
            return int(float(cleaned)) if cleaned else 0
        except ValueError:
            return 0
    return 0

def bytes_to_mb(bytes_value):
    bytes_value = safe_int_conversion(bytes_value)
    if bytes_value == 0:
        return 0.00
    return round(bytes_value / (1024 * 1024), 2)

def bytes_to_gb(bytes_value):
    bytes_value = safe_int_conversion(bytes_value)
    if bytes_value == 0:
        return 0.00
    return round(bytes_value / (1024 * 1024 * 1024), 2)

def format_datetime(datetime_str):
    if not datetime_str or datetime_str == "N/A" or datetime_str == "0":
        return "N/A"
    try:
        if 'T' in datetime_str:
            if '.' in datetime_str:
                dt = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%S.%fZ")
            else:
                dt = datetime.strptime(datetime_str, "%Y-%m-%dT%H:%M:%SZ")
            return dt.strftime("%Y-%m-%d %H:%M:%S")
        return datetime_str
    except:
        return datetime_str

def calculate_version_space_savings(versions, keep_last_n):
    if not versions:
        return {
            'total_versions': 0,
            'keep_count': 0,
            'delete_count': 0,
            'keep_size_bytes': 0,
            'delete_size_bytes': 0,
            'space_saved_gb': 0.0,
            'delete_range': 'N/A'
        }
    
    sorted_versions = sorted(versions, key=lambda x: x.get('created', ''))
    total_versions = len(sorted_versions)
    
    if total_versions <= keep_last_n:
        keep_count = total_versions
        delete_count = 0
        keep_size_bytes = sum(v.get('size', 0) for v in sorted_versions)
        delete_size_bytes = 0
        delete_range = 'None (within limit)'
    else:
        keep_count = keep_last_n
        delete_count = total_versions - keep_last_n
        keep_versions = sorted_versions[-keep_last_n:]
        delete_versions = sorted_versions[:-keep_last_n]
        
        keep_size_bytes = sum(v.get('size', 0) for v in keep_versions)
        delete_size_bytes = sum(v.get('size', 0) for v in delete_versions)
        
        first_deleted = delete_versions[0].get('version_label', 'unknown') if delete_versions else 'N/A'
        last_deleted = delete_versions[-1].get('version_label', 'unknown') if delete_versions else 'N/A'
        delete_range = f"{first_deleted} to {last_deleted}" if first_deleted != 'N/A' else 'N/A'
    
    return {
        'total_versions': total_versions,
        'keep_count': keep_count,
        'delete_count': delete_count,
        'keep_size_bytes': keep_size_bytes,
        'delete_size_bytes': delete_size_bytes,
        'space_saved_gb': bytes_to_gb(delete_size_bytes),
        'delete_range': delete_range
    }

# ============================================================
# SHAREPOINT DATA RETRIEVAL
# ============================================================

def get_all_libraries(site_url):
    print("\n📁 Getting document libraries...")
    lists_url = f"{site_url}/_api/web/lists"
    all_libraries = []
    next_url = lists_url
    
    while next_url:
        response = make_sharepoint_request(site_url, next_url)
        
        if not response or 'd' not in response:
            break
        
        if 'results' in response['d']:
            for lst in response['d']['results']:
                if lst['BaseTemplate'] == 101:
                    all_libraries.append({
                        'id': lst['Id'],
                        'title': lst['Title']
                    })
        
        next_url = None
        if '__next' in response.get('d', {}):
            next_url = response['d']['__next']
    
    return all_libraries

def get_all_items_from_library(site_url, library_id, max_items=None):
    """Get items with limit for testing"""
    print(f"    Fetching items from library...")
    
    items_url = f"{site_url}/_api/web/lists(guid'{library_id}')/items?$select=Id,Title,FileLeafRef,FileRef,Created,Modified,FileSystemObjectType&$top=5000"
    
    all_items = []
    next_url = items_url
    page_count = 0
    
    while next_url:
        page_count += 1
        print(f"    Fetching page {page_count}...", end="")
        response = make_sharepoint_request(site_url, next_url)
        
        if not response or 'd' not in response:
            print(" ✗ Failed")
            break
        
        if 'results' in response['d']:
            items_in_page = len(response['d']['results'])
            all_items.extend(response['d']['results'])
            print(f" ✓ Got {items_in_page} items")
            
            if max_items and len(all_items) >= max_items:
                print(f"    🎯 Reached {max_items} items limit")
                break
        
        next_url = None
        if '__next' in response.get('d', {}):
            next_url = response['d']['__next']
    
    if max_items and len(all_items) > max_items:
        all_items = all_items[:max_items]
    
    print(f"    Total items fetched: {len(all_items)}")
    return all_items

def get_file_versions(site_url, list_id, item_id):
    try:
        versions_url = f"{site_url}/_api/Web/Lists(guid'{list_id}')/items({item_id})/versions"
        
        response = make_sharepoint_request(site_url, versions_url)
        
        if not response or 'd' not in response:
            return []
        
        if 'results' not in response['d']:
            return []
        
        versions = []
        for version in response['d']['results']:
            version_data = {
                'version_id': version.get('VersionId', 0),
                'version_label': version.get('VersionLabel', ''),
                'created': version.get('Created', ''),
                'is_current': version.get('IsCurrentVersion', False),
                'size': safe_int_conversion(version.get('File_x005f_x0020_x005f_Size', 0)),
                'checkin_comment': version.get('OData__x005f_CheckinComment', '')
            }
            versions.append(version_data)
        
        return versions
        
    except Exception as e:
        return []

# ============================================================
# BATCH PROCESSING WITH PERFORMANCE TRACKING
# ============================================================

def process_single_file(site_url, list_id, item, library_title, output_file):
    """Process single file with performance tracking"""
    global global_analyzer
    
    try:
        item_id = item.get('Id')
        fsob_type = item.get('FileSystemObjectType', 0)
        
        if fsob_type != 0:
            return None
        
        file_name = item.get('FileLeafRef', '')
        if not file_name:
            file_name = item.get('Title', f"Item_{item_id}")
        
        versions = get_file_versions(site_url, list_id, item_id)
        version_count = len(versions)
        
        current_file_size = 0
        if versions:
            sorted_versions = sorted(versions, key=lambda x: x.get('created', ''))
            latest_version = sorted_versions[-1]
            current_file_size = latest_version.get('size', 0)
        
        file_size_mb = bytes_to_mb(current_file_size)
        
        if file_size_mb <= CONFIG['version_settings']['min_file_size_mb']:
            return None
        
        total_versions_size = 0
        for version in versions:
            total_versions_size += version.get('size', 0)
        
        savings_data = {}
        for keep_versions in CONFIG['version_settings']['keep_versions_options']:
            savings = calculate_version_space_savings(versions, keep_versions)
            savings_data[f'Keep_{keep_versions}'] = {
                'delete_count': savings['delete_count'],
                'space_saved_gb': savings['space_saved_gb'],
                'delete_range': savings['delete_range']
            }
        
        if global_analyzer:
            global_analyzer.log_file_processed()
        
        return {
            'library': library_title,
            'list_id': list_id,
            'item_id': item_id,
            'file_name': file_name,
            'file_path': item.get('FileRef', ''),
            'current_file_size_mb': file_size_mb,
            'version_count': version_count,
            'total_versions_size_mb': bytes_to_mb(total_versions_size),
            'savings_data': savings_data
        }
        
    except Exception as e:
        if global_analyzer:
            global_analyzer.log_other_error()
        return None

def process_batch(batch_items, site_url, library_id, library_title, output_file, batch_num, analyzer):
    """Process a batch of files"""
    global global_analyzer
    
    analyzer.log_batch_start(batch_num, len(batch_items))
    
    results = []
    
    with ThreadPoolExecutor(max_workers=analyzer.max_workers) as executor:
        futures = {
            executor.submit(
                process_single_file,
                site_url,
                library_id,
                item,
                library_title,
                output_file
            ): item
            for item in batch_items
        }
        
        for future in as_completed(futures):
            try:
                result = future.result()
                if result:
                    results.append(result)
            except Exception as e:
                if analyzer:
                    analyzer.log_other_error()
    
    analyzer.log_batch_end()
    return results

def process_site_with_analyzer(site_url, analyzer):
    """Process site with performance analyzer"""
    global global_analyzer
    global_analyzer = analyzer
    
    print(f"\n{'#'*80}")
    print(f"📍 SITE: {site_url}")
    print(f"{'#'*80}")
    
    libraries = get_all_libraries(site_url)
    
    if not libraries:
        print("No document libraries found.")
        return
    
    all_results = []
    total_processed = 0
    batch_num = 0
    
    for library in libraries:
        print(f"\n📁 Library: {library['title']}")
        
        items = get_all_items_from_library(
            site_url, 
            library['id'], 
            max_items=CONFIG['performance_test']['max_items_to_process']
        )
        
        if not items:
            print("  No items found")
            continue
        
        files = [item for item in items if item.get('FileSystemObjectType') == 0]
        
        print(f"  Found {len(files)} files")
        
        if not files:
            continue
        
        batch_size = analyzer.batch_size
        total_batches = (len(files) + batch_size - 1) // batch_size
        
        for i in range(0, len(files), batch_size):
            batch_num += 1
            batch = files[i:i + batch_size]
            
            if total_processed + len(batch) > CONFIG['performance_test']['max_items_to_process']:
                remaining = CONFIG['performance_test']['max_items_to_process'] - total_processed
                batch = batch[:remaining]
            
            results = process_batch(
                batch,
                site_url,
                library['id'],
                library['title'],
                None,
                batch_num,
                analyzer
            )
            
            total_processed += len(results)
            all_results.extend(results)
            
            batch = None
            gc.collect()
            
            if total_processed >= CONFIG['performance_test']['max_items_to_process']:
                print(f"\n🎯 Reached {CONFIG['performance_test']['max_items_to_process']} items limit")
                break
        
        if total_processed >= CONFIG['performance_test']['max_items_to_process']:
            break
    
    return all_results

# ============================================================
# PERFORMANCE TEST FUNCTIONS
# ============================================================

def run_performance_tests():
    """Run performance tests with different configurations"""
    
    test_configs = CONFIG['performance_test']['test_configs']
    
    all_results = []
    
    print("\n" + "="*100)
    print("🧪 PERFORMANCE ANALYSIS - {} ITEMS".format(CONFIG['performance_test']['max_items_to_process']))
    print("="*100)
    print(f"\n📌 Test Configurations:")
    for config in test_configs:
        print(f"  - {config['name']}")
    print("="*100)
    
    for config in test_configs:
        try:
            analyzer = PerformanceAnalyzer(
                config['name'], 
                config['batch_size'], 
                config['max_workers']
            )
            analyzer.start()
            
            process_site_with_analyzer(CONFIG['sites'][0], analyzer)
            
            results = analyzer.finish()
            all_results.append(results)
            
            time.sleep(5)
            gc.collect()
            
        except Exception as e:
            print(f"❌ Error in {config['name']}: {str(e)}")
            continue
    
    print_comparison_report(all_results)
    save_summary_to_csv(all_results)
    
    return all_results

# ============================================================
# COMPARISON REPORT
# ============================================================

def print_comparison_report(all_results):
    """Print comparison of all configurations"""
    
    print("\n" + "="*100)
    print("📊 COMPARISON REPORT - ALL CONFIGURATIONS")
    print("="*100)
    
    print("\n" + "-"*100)
    print(f"{'Config':<15} {'Batch':<8} {'Workers':<8} {'Time(s)':<12} {'Files/sec':<12} {'Peak Mem':<12} {'429 Errors':<12} {'Success':<10}")
    print("-"*100)
    
    for results in all_results:
        success = "✅" if results['rate_limit_errors'] == 0 else "⚠️"
        print(f"{results['config']:<15} "
              f"{results['batch_size']:<8} "
              f"{results['max_workers']:<8} "
              f"{results['total_time']:<12.2f} "
              f"{results['throughput']:<12.2f} "
              f"{results['peak_memory']:<12.1f} "
              f"{results['rate_limit_errors']:<12} "
              f"{success:<10}")
    
    print("-"*100)
    
    if all_results:
        best = min(all_results, key=lambda x: x['total_time'])
        fastest = max(all_results, key=lambda x: x['throughput'])
        lowest_memory = min(all_results, key=lambda x: x['peak_memory'])
        no_errors = [r for r in all_results if r['rate_limit_errors'] == 0]
        
        print(f"\n🏆 BEST PERFORMERS:")
        print(f"  Fastest: {best['config']} - {best['total_time']:.2f}s")
        print(f"  Highest Throughput: {fastest['config']} - {fastest['throughput']:.2f} files/sec")
        print(f"  Lowest Memory: {lowest_memory['config']} - {lowest_memory['peak_memory']:.1f} MB")
        
        if no_errors:
            best_no_errors = min(no_errors, key=lambda x: x['total_time'])
            print(f"  Best with no 429 errors: {best_no_errors['config']} - {best_no_errors['total_time']:.2f}s")
    
    print("="*100)
    print("\n💡 RECOMMENDATIONS:")
    
    if no_errors:
        best_no_errors = min(no_errors, key=lambda x: x['total_time'])
        print(f"  ✅ Best overall: {best_no_errors['config']}")
    else:
        print("  ⚠️ All configs had 429 errors - reduce workers")
    
    print("  ✅ For speed: Use smaller batch (100)")
    print("  ✅ For stability: Use fewer workers (20)")
    print("  ❌ Avoid: Large batch (300) - causes memory issues")
    print("="*100)

def save_summary_to_csv(all_results):
    """Save results to CSV file"""
    output_dir = CONFIG['output']['output_dir']
    os.makedirs(output_dir, exist_ok=True)
    
    csv_file = os.path.join(output_dir, CONFIG['output']['summary_file'])
    
    fieldnames = [
        'Config Name', 'Batch Size', 'Workers', 'Total Items',
        'Total Time (s)', 'Throughput (files/sec)', 
        'Peak Memory (MB)', '429 Errors', 'Retries', 
        'Other Errors', 'Success'
    ]
    
    with open(csv_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        
        for results in all_results:
            row = {
                'Config Name': results['config'],
                'Batch Size': results['batch_size'],
                'Workers': results['max_workers'],
                'Total Items': results['total_items'],
                'Total Time (s)': round(results['total_time'], 2),
                'Throughput (files/sec)': round(results['throughput'], 2),
                'Peak Memory (MB)': round(results['peak_memory'], 1),
                '429 Errors': results['rate_limit_errors'],
                'Retries': results['retry_count'],
                'Other Errors': results['other_errors'],
                'Success': 'Yes' if results['rate_limit_errors'] == 0 else 'No'
            }
            writer.writerow(row)
    
    print(f"📁 Summary saved to: {csv_file}")

# ============================================================
# MAIN FUNCTION
# ============================================================

def main():
    """Main function to run performance tests"""
    global ALLOWED_FILE_EXTENSIONS
    
    print("="*100)
    print("🔬 PERFORMANCE ANALYSIS SCRIPT")
    print(f"📊 Testing with {CONFIG['performance_test']['max_items_to_process']} items")
    print("="*100)
    
    ALLOWED_FILE_EXTENSIONS = normalize_extensions(CONFIG['file_extensions'])
    
    os.makedirs(CONFIG['output']['output_dir'], exist_ok=True)
    
    print(f"\n📌 Configuration:")
    print(f"  Site: {CONFIG['sites'][0]}")
    print(f"  Max Items: {CONFIG['performance_test']['max_items_to_process']}")
    print(f"  Min File Size: {CONFIG['version_settings']['min_file_size_mb']} MB")
    print(f"  File Extensions: {CONFIG['file_extensions']}")
    print("="*100)
    
    print("\n🔐 Authenticating...")
    access_token = get_cached_token()
    
    if not access_token:
        print("❌ Authentication failed")
        return
    
    print("✅ Authentication successful\n")
    
    all_results = run_performance_tests()
    
    print("\n✅ ANALYSIS COMPLETE!")
    print(f"📁 Results saved in: {CONFIG['output']['output_dir']}")

if __name__ == "__main__":
    main()