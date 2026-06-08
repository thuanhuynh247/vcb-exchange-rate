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

# Let's test different payloads
payloads = [
    # Get transfer_rate history for USD from 2026-01-01 to 2026-06-08
    ["2026-01-01", "2026-06-08", "USD", "transfer_rate"],
    # Try empty dates
    ["", "", "USD", "transfer_rate"]
]

for p in payloads:
    print(f"\nSending payload: {p}")
    # Next.js Server Action payload is raw JSON array representation
    data = json.dumps(p)
    r = requests.post(url, headers=headers, data=data, verify=False)
    print(f"Status: {r.status_code}")
    print(f"Response (first 1000 chars):")
    print(r.text[:1000])
