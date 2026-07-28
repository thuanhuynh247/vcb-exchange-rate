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

all_actions = []

# Fetch and combine JS chunk text
for chunk in chunks:
    url = base_url + chunk
    r = requests.get(url, headers=headers, verify=False)
    if r.status_code == 200:
        # Find all 40-character hex strings
        actions = re.findall(r'\b[a-f0-9]{40}\b', r.text)
        for act in actions:
            # Get surrounding context (100 chars before and after)
            idx = r.text.find(act)
            context = r.text[max(0, idx-100):min(len(r.text), idx+140)].replace('\n', ' ')
            all_actions.append((act, chunk, context))

print(f"Collected {len(all_actions)} actions.")

# Search context for rate keywords
rate_keywords = ['usd', 'rate', 'transfer', 'sell', 'cash', '15:45']
matching_actions = []
for act, chunk, context in all_actions:
    lower_context = context.lower()
    matches = [k for k in rate_keywords if k in lower_context]
    if matches:
        matching_actions.append((act, chunk, context, matches))

print(f"\nFound {len(matching_actions)} actions matching rate keywords:")
for act, chunk, context, matches in matching_actions:
    print(f"\nAction: {act} (Matches: {matches})")
    print(f"Chunk: {chunk}")
    print(f"Context: {context}")
