import requests
import re
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

base_url = 'https://www.vietinbank.vn'
chunks = [
    '/_next/static/chunks/app/%5Blang%5D/%5B...slug%5D/page-04b0b90a124d12b4.js',
    '/_next/static/chunks/main-app-050fd630f446e664.js',
    '/_next/static/chunks/webpack-34218bbe47058fb3.js'
]

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

for chunk in chunks:
    url = base_url + chunk
    print(f"--- Chunk: {chunk} ---")
    r = requests.get(url, headers=headers, verify=False)
    if r.status_code == 200:
        # Let's search for 40-char hex strings
        actions = re.findall(r'\b[a-f0-9]{40}\b', r.text)
        print(f"Found {len(actions)} actions.")
        for act in set(actions):
            print(act)
    else:
        print(f"Failed: {r.status_code}")
