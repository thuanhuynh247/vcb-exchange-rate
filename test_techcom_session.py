import requests
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8",
    "Referer": "https://techcombank.com/cong-cu-tien-ich/ty-gia"
}

session = requests.Session()
# First get the CSRF token
csrf_url = "https://techcombank.com/libs/granite/csrf/token.json"
csrf_resp = session.get(csrf_url, headers=headers, verify=False, timeout=10)
print(f"CSRF status: {csrf_resp.status_code}")
if csrf_resp.status_code == 200:
    try:
        csrf_data = csrf_resp.json()
        print(f"CSRF token: {csrf_data}")
    except:
        pass

# Now try the exchange rate endpoint
rate_url = "https://techcombank.com/content/techcombank/web/vn/vi/cong-cu-tien-ich/ty-gia/_jcr_content.exchange-rates.integration.json"
rate_resp = session.get(rate_url, headers=headers, verify=False, timeout=10)
print(f"\nRate API status: {rate_resp.status_code}")
print(f"Response (2000 chars): {rate_resp.text[:2000]}")
