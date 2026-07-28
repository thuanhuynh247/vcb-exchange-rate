import requests
import json
import datetime
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://www.vietinbank.vn/ty-gia-khcn"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Content-Type": "text/plain;charset=UTF-8",
    "accept": "text/x-component",
    "referer": "https://www.vietinbank.vn/ty-gia-khcn"
}

current_action = "1e43a43a5124d6cc3cb463bc54021b34f39a4065"
history_action = "ff24b60505a8da357a655878afe7dd2d1f9f0e52"

# Test current action
date_str = datetime.datetime.now().strftime('%Y-%m-%d')
payload_current = [f"{date_str}T15:45:00", "USD"]
print("Testing current action ID:", current_action)
try:
    h = headers.copy()
    h["next-action"] = current_action
    r = requests.post(url, headers=h, data=json.dumps(payload_current), verify=False, timeout=10)
    print("Status code:", r.status_code)
    print("Response snippet:", r.text[:500])
except Exception as e:
    print("Error:", e)

# Test history action
payload_history = ["2026-06-01", "2026-06-10", "USD", "transfer_rate"]
print("\nTesting history action ID:", history_action)
try:
    h = headers.copy()
    h["next-action"] = history_action
    r = requests.post(url, headers=h, data=json.dumps(payload_history), verify=False, timeout=10)
    print("Status code:", r.status_code)
    print("Response snippet:", r.text[:500])
except Exception as e:
    print("Error:", e)
