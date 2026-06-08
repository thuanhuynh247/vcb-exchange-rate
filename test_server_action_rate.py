import requests
import json
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://www.vietinbank.vn/ty-gia-khcn"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Content-Type": "text/plain;charset=UTF-8",
    "next-action": "1e43a43a5124d6cc3cb463bc54021b34f39a4065",
    "accept": "text/x-component"
}

# Test query for a historical date: January 2, 2026
payloads = [
    ["2026-01-02T15:45:00", "USD"],
    ["2026-01-02T08:00:00", "USD"],
    ["2026-01-02T00:00:00", "USD"]
]

for p in payloads:
    print(f"\nSending payload: {p}")
    data = json.dumps(p)
    r = requests.post(url, headers=headers, data=data, verify=False)
    print(f"Status: {r.status_code}")
    print(f"Response:")
    print(r.text)
