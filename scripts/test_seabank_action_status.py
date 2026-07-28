import requests
import json
import datetime
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://www.seabank.com.vn/cong-cu-tien-ich/ty-gia"
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
    "Content-Type": "text/plain;charset=UTF-8",
    "accept": "text/x-component",
    "referer": "https://www.seabank.com.vn/cong-cu-tien-ich/ty-gia"
}

action = "d65f2411081b93638167328d79ca76cf2bc7ec18"
date_str = datetime.datetime.now().strftime('%d/%m/%Y')
payload = [date_str]

print("Testing SeaBank action ID:", action, "with date:", date_str)
try:
    h = headers.copy()
    h["next-action"] = action
    r = requests.post(url, headers=h, data=json.dumps(payload), verify=False, timeout=10)
    print("Status code:", r.status_code)
    print("Response snippet:", r.text[:1000])
except Exception as e:
    print("Error:", e)
