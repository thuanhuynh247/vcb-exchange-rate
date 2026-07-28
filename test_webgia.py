import requests

url = "https://webgia.com/ty-gia/vietinbank/"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
try:
    r = requests.get(url, headers=headers, timeout=10)
    print("Status:", r.status_code)
    html = r.text
    # Print lines containing "USD" or "lịch sử" or date patterns
    lines = html.splitlines()
    found = 0
    for line in lines:
        if "USD" in line or "ngay" in line or "history" in line or "lich-su" in line:
            print(line[:200])
            found += 1
            if found > 40:
                break
except Exception as e:
    print("Error:", e)
