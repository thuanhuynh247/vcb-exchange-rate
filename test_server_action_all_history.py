import requests
import json
import urllib3
from datetime import datetime

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://www.vietinbank.vn/ty-gia-khcn"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Content-Type": "text/plain;charset=UTF-8",
    "next-action": "ff24b60505a8da357a655878afe7dd2d1f9f0e52",
    "accept": "text/x-component"
}

def clean_response(text):
    # Next.js Server Action returns lines prefixing with numbers like "1:[\n"
    # Let's extract the list
    lines = text.strip().split('\n')
    for line in lines:
        if line.startswith("1:"):
            return json.loads(line[2:])
    return []

# Query transfer_rate history
print("Querying transfer_rate history...")
payload_transfer = ["2026-01-01", "2026-06-08", "USD", "transfer_rate"]
r_transfer = requests.post(url, headers=headers, data=json.dumps(payload_transfer), verify=False)
data_transfer = clean_response(r_transfer.text)

# Query sell_rate history
print("Querying sell_rate history...")
payload_sell = ["2026-01-01", "2026-06-08", "USD", "sell_rate"]
r_sell = requests.post(url, headers=headers, data=json.dumps(payload_sell), verify=False)
data_sell = clean_response(r_sell.text)

print(f"Transfer rate history records: {len(data_transfer)}")
print(f"Sell rate history records: {len(data_sell)}")

# Let's map by date
transfer_map = {item['apply_date']: item for item in data_transfer}
sell_map = {item['apply_date']: item for item in data_sell}

# Merge them by date
all_dates = sorted(list(set(transfer_map.keys()) | set(sell_map.keys())))
print(f"Total unique dates: {len(all_dates)}")

for date in all_dates[:10]:
    t_val = transfer_map.get(date, {}).get('close', 'N/A')
    s_val = sell_map.get(date, {}).get('close', 'N/A')
    print(f"Date: {date} | Transfer Buy: {t_val} | Sell: {s_val}")
