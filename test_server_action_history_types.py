import requests
import json
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://www.vietinbank.vn/ty-gia-khcn"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Content-Type": "text/plain;charset=UTF-8",
    "next-action": "ff24b60505a8da357a655878afe7dd2d1f9f0e52",
    "accept": "text/x-component"
}

# Test different rate types
rate_types = ["transfer_rate", "sell_rate", "cash_rate_big", "cash_rate_small", "cash_rate", "all"]

for rt in rate_types:
    print(f"\nSending payload for rate_type: {rt}")
    payload = ["2026-06-01", "2026-06-08", "USD", rt]
    data = json.dumps(payload)
    r = requests.post(url, headers=headers, data=data, verify=False)
    print(f"Status: {r.status_code}")
    print(f"Response (first 500 chars):")
    print(r.text[:500])
