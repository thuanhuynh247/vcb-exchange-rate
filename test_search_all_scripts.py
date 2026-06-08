import requests
import re

url = "https://www.vietinbank.vn/ty-gia-khcn"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
r = requests.get(url, headers=headers, verify=False)

scripts = re.findall(r'<script[^>]*>(.*?)</script>', r.text, re.DOTALL)
print(f"Total scripts: {len(scripts)}")

# Let's search all scripts for keywords
keywords = ["api", "exchangerate", "exchange-rate", "ty-gia", "tygia", "rate"]
for i, script in enumerate(scripts):
    found = []
    for kw in keywords:
        if kw in script.lower():
            found.append(kw)
    if found:
        print(f"Script {i} matches: {found}")
        # Search for URLs and path-like strings inside this script
        urls = re.findall(r'https?://[^\s"\'\\<>]+', script)
        paths = re.findall(r'/[a-zA-Z0-9_\-\./]+', script)
        for u in set(urls):
            if any(kw in u.lower() for kw in keywords):
                print(f"  URL: {u}")
        for p in set(paths):
            if any(kw in p.lower() for kw in ["api", "exchangerate", "exchange-rate", "ty-gia", "rate"]):
                if len(p) < 150:
                    print(f"  Path: {p}")
