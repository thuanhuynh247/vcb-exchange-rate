import requests
import json
import datetime
import urllib3
import concurrent.futures
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def discover_vietinbank_actions():
    import re
    homepage_url = 'https://www.vietinbank.vn/ty-gia-khcn'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
    }
    try:
        r = requests.get(homepage_url, headers=headers, verify=False, timeout=10)
        if r.status_code != 200:
            return None, None
        
        scripts = re.findall(r'<script[^>]+src=["\']([^"\']+)["\']', r.text)
        preloads = re.findall(r'<link[^>]*rel=["\']preload["\'][^>]*as=["\']script["\'][^>]*href=["\']([^"\']+)["\']', r.text)
        all_sources = list(set(scripts + preloads))
        
        chunk_urls = []
        for src in all_sources:
            if '_next/static/chunks/' in src:
                if not src.startswith('http'):
                    src = 'https://www.vietinbank.vn' + src
                chunk_urls.append(src)
        
        print(f"[VietinBank] Scanning {len(chunk_urls)} static chunks...")
        action_candidates = set()
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
            futures = [executor.submit(requests.get, url, headers=headers, verify=False, timeout=10) for url in chunk_urls]
            for fut in concurrent.futures.as_completed(futures):
                try:
                    resp = fut.result()
                    if resp.status_code == 200:
                        matches = re.findall(r'\b[a-f0-9]{40}\b', resp.text)
                        for m in matches:
                            action_candidates.add(m)
                except Exception as e:
                    pass
        
        print(f"[VietinBank] Found {len(action_candidates)} unique action candidates. Testing...")
        current_act = None
        history_act = None
        date_str = datetime.datetime.now().strftime('%Y-%m-%d')
        
        payload_curr = [f"{date_str}T15:45:00", "USD"]
        payload_hist = [date_str, date_str, "USD", "transfer_rate"]
        
        def test_action(act):
            h_test = {
                "User-Agent": headers['User-Agent'],
                "Content-Type": "text/plain;charset=UTF-8",
                "next-action": act,
                "accept": "text/x-component",
                "referer": homepage_url
            }
            # Test for history
            try:
                r_test = requests.post(homepage_url, headers=h_test, data=json.dumps(payload_hist), verify=False, timeout=5)
                if r_test.status_code == 200 and 'currency' in r_test.text and 'apply_date' in r_test.text:
                    return ('history', act)
            except:
                pass
            # Test for current
            try:
                r_test = requests.post(homepage_url, headers=h_test, data=json.dumps(payload_curr), verify=False, timeout=5)
                if r_test.status_code == 200:
                    for line in r_test.text.strip().split('\n'):
                        if line.startswith("1:"):
                            data = json.loads(line[2:])
                            if isinstance(data, list):
                                return ('current', act)
            except:
                pass
            return None

        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as test_executor:
            results = test_executor.map(test_action, action_candidates)
            for res in results:
                if res:
                    type_, act = res
                    if type_ == 'current':
                        current_act = act
                    elif type_ == 'history':
                        history_act = act
                        
        return current_act, history_act
    except Exception as e:
        print("Error during VietinBank discovery:", e)
        return None, None

def discover_seabank_action():
    import re
    homepage_url = 'https://www.seabank.com.vn/cong-cu-tien-ich/ty-gia'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
    }
    try:
        r = requests.get(homepage_url, headers=headers, verify=False, timeout=10)
        if r.status_code != 200:
            return None
        
        scripts = re.findall(r'<script[^>]+src=["\']([^"\']+)["\']', r.text)
        preloads = re.findall(r'<link[^>]*rel=["\']preload["\'][^>]*as=["\']script["\'][^>]*href=["\']([^"\']+)["\']', r.text)
        all_sources = list(set(scripts + preloads))
        
        chunk_urls = []
        for src in all_sources:
            if '_next/static/chunks/' in src:
                if not src.startswith('http'):
                    src = 'https://www.seabank.com.vn' + src
                chunk_urls.append(src)
                
        print(f"[SeaBank] Scanning {len(chunk_urls)} static chunks...")
        action_candidates = set()
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
            futures = [executor.submit(requests.get, url, headers=headers, verify=False, timeout=10) for url in chunk_urls]
            for fut in concurrent.futures.as_completed(futures):
                try:
                    resp = fut.result()
                    if resp.status_code == 200:
                        matches = re.findall(r'\b[a-f0-9]{40}\b', resp.text)
                        for m in matches:
                            action_candidates.add(m)
                except Exception as e:
                    pass
                    
        print(f"[SeaBank] Found {len(action_candidates)} unique action candidates. Testing...")
        date_str = datetime.datetime.now().strftime('%d/%m/%Y')
        payload = [date_str]
        
        def test_action(act):
            h_test = {
                "User-Agent": headers['User-Agent'],
                "Content-Type": "text/plain;charset=UTF-8",
                "next-action": act,
                "accept": "text/x-component",
                "referer": homepage_url
            }
            try:
                r_test = requests.post(homepage_url, headers=h_test, data=json.dumps(payload), verify=False, timeout=5)
                if r_test.status_code == 200 and 'details' in r_test.text and 'currency' in r_test.text:
                    return act
            except:
                pass
            return None
            
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as test_executor:
            results = test_executor.map(test_action, action_candidates)
            for res in results:
                if res:
                    return res
        return None
    except Exception as e:
        print("Error during SeaBank discovery:", e)
        return None

if __name__ == '__main__':
    print("Testing dynamic VietinBank action discovery...")
    v_curr, v_hist = discover_vietinbank_actions()
    print(f"-> Discovered VietinBank current rate Action ID: {v_curr}")
    print(f"-> Discovered VietinBank history rate Action ID: {v_hist}")
    
    print("\nTesting dynamic SeaBank action discovery...")
    s_curr = discover_seabank_action()
    print(f"-> Discovered SeaBank current rate Action ID: {s_curr}")
