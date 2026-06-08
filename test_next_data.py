import requests
import re
import json

url = "https://www.vietinbank.vn/ty-gia-khcn"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
r = requests.get(url, headers=headers, verify=False)

# Let's search for self.__next_f.push lines or script content
scripts = re.findall(r'<script[^>]*>(.*?)</script>', r.text, re.DOTALL)
print(f"Found {len(scripts)} scripts.")

# Let's look for interesting texts in scripts
keywords = ["api", "exchange", "rate", "exchangerate", "ty-gia", "vietinbank.vn/"]
for i, script in enumerate(scripts):
    found_kw = [kw for kw in keywords if kw in script.lower()]
    if found_kw:
        print(f"Script {i} contains keywords: {found_kw}")
        # Let's print a snippet
        lines = script.splitlines()
        print(f"  Total lines: {len(lines)}")
        for line in lines:
            if any(kw in line.lower() for kw in keywords):
                if len(line.strip()) < 300:
                    print(f"    {line.strip()[:200]}")
