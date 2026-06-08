import requests
import re

url = "https://www.vietinbank.vn/ty-gia-khcn"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
r = requests.get(url, headers=headers, verify=False)

scripts = re.findall(r'<script[^>]*>(.*?)</script>', r.text, re.DOTALL)
script_49 = scripts[49]

# Search for any URLs or api paths in script_49
urls = re.findall(r'https?://[^\s"\'\\<>]+', script_49)
paths = re.findall(r'/[a-zA-Z0-9_\-\./]+', script_49)

print("URLs in Script 49:")
for u in set(urls):
    print(" ", u)

print("\nPaths in Script 49 (matching /api or /ex or /ty-gia or rate):")
for p in set(paths):
    if any(keyword in p.lower() for keyword in ["api", "ex", "rate", "ty-gia", "khcn"]):
        print(" ", p)
        
# Also let's print a subset of script_49 text around occurrences of "api"
for m in re.finditer(r'api', script_49, re.IGNORECASE):
    start = max(0, m.start() - 100)
    end = min(len(script_49), m.end() + 100)
    print(f"\nContext around '{m.group(0)}':")
    print(script_49[start:end])
