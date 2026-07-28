import requests
import urllib3
urllib3.disable_warnings()

csrf_url = "https://techcombank.com/libs/granite/csrf/token.json"
rate_url = "https://techcombank.com/content/techcombank/web/vn/vi/cong-cu-tien-ich/ty-gia/_jcr_content.exchange-rates.integration.json"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
    "Referer": "https://techcombank.com/cong-cu-tien-ich/ty-gia"
}

def test_param(param_dict):
    session = requests.Session()
    try:
        session.get(csrf_url, headers=HEADERS, verify=False, timeout=10)
        r = session.get(rate_url, params=param_dict, headers=HEADERS, verify=False, timeout=10)
        if r.status_code == 200:
            data = r.json()
            rates = []
            for item in data.get("exchangeRate", {}).get("data", []):
                if item.get("label", "").startswith("USD (50"):
                    rates.append((item.get('bidRateTM'), item.get('askRate')))
            return rates
        return f"Status {r.status_code}"
    except Exception as e:
        return str(e)

# Test cases
print("Today:", test_param({}))
print("date 2026-03-18:", test_param({"date": "2026-03-18"}))
print("date 18/03/2026:", test_param({"date": "18/03/2026"}))
print("date 18-03-2026:", test_param({"date": "18-03-2026"}))
print("d 20260318:", test_param({"d": "20260318"}))
print("t 1773830400:", test_param({"t": "1773830400"}))
