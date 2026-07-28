import requests
import re
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = 'https://www.vietinbank.vn/ty-gia-khcn'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

print("Fetching VietinBank page...")
r = requests.get(url, headers=headers, verify=False)
print(f"Status: {r.status_code}, Length: {len(r.text)}")

# Look for any JS scripts containing action definitions
scripts = re.findall(r'<script[^>]+src=["\']([^"\']+)["\']', r.text)
print("\nFound scripts:")
for s in scripts:
    print(s)

# Search for any action ID in the page itself (40-char hex strings)
hex_actions = re.findall(r'\b[a-f0-9]{40}\b', r.text)
print(f"\nFound {len(hex_actions)} 40-char hex actions in page HTML:")
for ha in set(hex_actions):
    print(ha)
