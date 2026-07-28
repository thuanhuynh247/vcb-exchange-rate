import requests
import re
import urllib3
import json
import datetime

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = 'https://www.vietinbank.vn/ty-gia-khcn'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

print("Fetching VietinBank main page...")
r = requests.get(url, headers=headers, verify=False)
if r.status_code != 200:
    print(f"Failed to fetch page: {r.status_code}")
    exit(1)

# Find all script src tags
scripts = re.findall(r'<script[^>]+src=["\']([^"\']+)["\']', r.text)
print(f"Found {len(scripts)} scripts in HTML.")

# Download scripts and find action IDs
action_ids = set()
for s in scripts:
    if not s.startswith('http'):
        s_url = 'https://www.vietinbank.vn' + s
    else:
        s_url = s
    
    try:
        resp = requests.get(s_url, headers=headers, verify=False, timeout=10)
        if resp.status_code == 200:
            actions = re.findall(r'\b[a-f0-9]{40}\b', resp.text)
            if actions:
                print(f"  {s}: found {len(actions)} actions")
                for act in actions:
                    action_ids.add(act)
    except Exception as e:
        print(f"  Error fetching {s_url}: {e}")

print(f"\nTotal unique Action IDs found: {len(action_ids)}")
print(list(action_ids))

# Test each Action ID with current rates payload
date_str = datetime.datetime.now().strftime('%Y-%m-%d')
payload = [f"{date_str}T15:45:00", "USD"]
print(f"\nTesting Action IDs with payload: {payload}...")

for act in action_ids:
    h = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Content-Type": "text/plain;charset=UTF-8",
        "next-action": act,
        "accept": "text/x-component",
        "referer": "https://www.vietinbank.vn/ty-gia-khcn"
    }
    try:
        resp = requests.post(url, headers=h, data=json.dumps(payload), verify=False, timeout=5)
        if resp.status_code == 200:
            text = resp.text.strip()
            # If the response contains exchange rate keywords or numbers, print it!
            if "rate" in text.lower() or "cash" in text.lower() or "sell" in text.lower() or ("[" in text and len(text) > 100):
                print(f"\n⭐ SUCCESS for Action ID: {act}")
                print(f"Response: {text[:1000]}")
    except Exception as e:
        pass
