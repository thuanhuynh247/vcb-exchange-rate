import requests
import re
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Let's extract the JS chunk URLs from the main page
url = "https://www.vietinbank.vn/ty-gia-khcn"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
r = requests.get(url, headers=headers, verify=False)

js_files = re.findall(r'src="(/_next/static/chunks/[^"]+\.js)"', r.text)
print(f"Found {len(js_files)} JS chunks.")

for js_path in js_files:
    js_url = "https://www.vietinbank.vn" + js_path
    js_r = requests.get(js_url, headers=headers, verify=False)
    if js_r.status_code == 200:
        # Search for transfer_rate or next-action patterns
        if "transfer_rate" in js_r.text:
            print(f"Found 'transfer_rate' in {js_url}")
            # Find all strings around 'transfer_rate'
            matches = re.findall(r'["\'][a-zA-Z0-9_-]*transfer_rate[a-zA-Z0-9_-]*["\']', js_r.text)
            print("  Matches:", matches)
            # Find other rate types near it
            for m in re.finditer(r'transfer_rate', js_r.text):
                start = max(0, m.start() - 150)
                end = min(len(js_r.text), m.end() + 150)
                print(f"  Context: {js_r.text[start:end]}")
