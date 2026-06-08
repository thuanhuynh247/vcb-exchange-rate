import requests
import json
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://www.seabank.com.vn/cong-cu-tien-ich/ty-gia"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Content-Type": "text/plain;charset=UTF-8",
    "next-action": "d65f2411081b93638167328d79ca76cf2bc7ec18",
    "accept": "text/x-component",
    "referer": "https://www.seabank.com.vn/cong-cu-tien-ich/ty-gia"
}

def clean_response(text):
    for line in text.strip().split('\n'):
        if line.startswith("1:"):
            return json.loads(line[2:])
    return None

# Test current date
dates = ["08/06/2026", "01/06/2026", "02/01/2026"]
for d in dates:
    payload = json.dumps([d])
    r = requests.post(url, headers=headers, data=payload, verify=False)
    data = clean_response(r.text)
    if data and isinstance(data, dict) and 'details' in data:
        usd = next((x for x in data['details'] if x.get('currency') == 'USD'), None)
        if usd:
            print(f"Date: {d} -> Buy: {usd['buy']}, Transfer: {usd['transferBuy']}, Sell: {usd['sell']}")
        else:
            print(f"Date: {d} -> No USD found. Currencies: {[x['currency'] for x in data['details']]}")
    else:
        print(f"Date: {d} -> Empty/error response")
        print(f"  Raw: {r.text[:200]}")
