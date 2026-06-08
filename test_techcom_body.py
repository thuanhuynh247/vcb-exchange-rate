import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

url = "https://techcombank.com/content/techcombank/web/vn/vi/cong-cu-tien-ich/ty-gia/_jcr_content.exchange-rates.2026-01-02.integration.json"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

r = requests.get(url, headers=headers, verify=False)
print("Status:", r.status_code)
print("Body:", r.text)
