import requests
import re
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

base_url = 'https://www.vietinbank.vn'
chunks = [
    '/_next/static/chunks/app/%5Blang%5D/%5B...slug%5D/page-04b0b90a124d12b4.js',
    '/_next/static/chunks/main-app-050fd630f446e664.js',
    '/_next/static/chunks/webpack-34218bbe47058fb3.js',
    '/_next/static/chunks/4e6af11a-0492817a5ef7718a.js',
    '/_next/static/chunks/164f4fb6-29d019f6917676c9.js',
    '/_next/static/chunks/3bcdfda6-a6746ab3ab542a4d.js',
    '/_next/static/chunks/13b76428-ebdf3012af0e4489.js',
    '/_next/static/chunks/472-5f772e5ecb3d621d.js',
    '/_next/static/chunks/660-ad0172d0bc49ab0b.js',
    '/_next/static/chunks/241-f6b49717a96e2631.js'
]

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

for chunk in chunks:
    url = base_url + chunk
    print(f"Fetching {chunk}...")
    r = requests.get(url, headers=headers, verify=False)
    if r.status_code == 200:
        # Let's search for 40-char hex strings
        actions = re.findall(r'\b[a-f0-9]{40}\b', r.text)
        if actions:
            print(f"  Found {len(actions)} actions:")
            for act in set(actions):
                # Also print the surrounding text context of the match
                idx = r.text.find(act)
                start = max(0, idx - 50)
                end = min(len(r.text), idx + 90)
                context = r.text[start:end].replace('\n', ' ')
                print(f"    - {act} | Context: ... {context} ...")
    else:
        print(f"  Failed with status {r.status_code}")
