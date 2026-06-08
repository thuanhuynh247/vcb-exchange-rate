import requests
import json
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

urls = [
    # Pattern 1: yyyy-mm-dd
    "https://techcombank.com/content/techcombank/web/vn/vi/cong-cu-tien-ich/ty-gia/_jcr_content.exchange-rates.2026-01-02.integration.json",
    # Pattern 2: dd-mm-yyyy
    "https://techcombank.com/content/techcombank/web/vn/vi/cong-cu-tien-ich/ty-gia/_jcr_content.exchange-rates.02-01-2026.integration.json",
    # Pattern 3: yyyymmdd
    "https://techcombank.com/content/techcombank/web/vn/vi/cong-cu-tien-ich/ty-gia/_jcr_content.exchange-rates.20260102.integration.json"
]

for url in urls:
    print(f"\nTrying URL: {url}")
    r = requests.get(url, headers=headers, verify=False)
    print(f"Status code: {r.status_code}")
    if r.status_code == 200:
        try:
            data = r.json()
            # Find USD rate
            usd_rate = None
            for item in data.get("exchangeRate", {}).get("data", []):
                if item.get("label") == "USD":
                    usd_rate = item
                    break
            print("SUCCESS! USD Rate data:")
            print(usd_rate)
        except Exception as e:
            print(f"Failed to parse JSON: {e}")
            print(r.text[:200])
