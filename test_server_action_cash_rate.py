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

# Test different names for cash rate history
rate_types = [
    "cash_buy", "cash_buy_rate", "buy_cash_rate", "buy_cash",
    "cash_rate_big", "cash_buy_big", "buy_rate", "cash_big", "cash"
]

for rt in rate_types:
    payload = ["2026-06-01", "2026-06-08", "USD", rt]
    data = json.dumps(payload)
    r = requests.post(url, headers=headers, data=data, verify=False)
    # Check if we get a non-empty array inside the response
    if "currency" in r.text:
        print(f"SUCCESS for rate_type: {rt}")
        print(r.text[:300])
    else:
        print(f"Empty/failed for rate_type: {rt}")
