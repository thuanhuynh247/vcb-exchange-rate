import requests
import re
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = 'https://www.vietinbank.vn/ty-gia-khcn'
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36'
}
r = requests.get(url, headers=headers, verify=False)
print("HTML Length:", len(r.text))
print("All script elements:")
for m in re.finditer(r'<script[^>]*src=["\']([^"\']+)["\']', r.text):
    print("Script Src:", m.group(1))

print("All preloaded script elements:")
for m in re.finditer(r'<link[^>]*rel=["\']preload["\'][^>]*as=["\']script["\'][^>]*href=["\']([^"\']+)["\']', r.text):
    print("Preload Src:", m.group(1))
