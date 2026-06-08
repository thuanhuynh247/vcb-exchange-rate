import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

csrf_url = "https://techcombank.com/libs/granite/csrf/token.json"
rate_url = "https://techcombank.com/content/techcombank/web/vn/vi/cong-cu-tien-ich/ty-gia/_jcr_content.exchange-rates.integration.json"
session_headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://techcombank.com/cong-cu-tien-ich/ty-gia"
}

session = requests.Session()
csrf_resp = session.get(csrf_url, headers=session_headers, verify=False, timeout=10)
print(f"CSRF status: {csrf_resp.status_code}")

r = session.get(rate_url, headers=session_headers, verify=False, timeout=15)
print(f"Rate API status: {r.status_code}")
print(f"Rate API text (first 500): {r.text[:500]}")

import json
data = r.json()
print("\nexchangeRate labels:")
for item in data.get("exchangeRate", {}).get("data", []):
    print(f"  label: {item.get('label')!r}")
    if item.get("label") == "USD":
        print("  -> Found USD!")
        print(f"     bidRateTM: {item.get('bidRateTM')}")
        print(f"     bidRateCK: {item.get('bidRateCK')}")
        print(f"     askRate: {item.get('askRate')}")
