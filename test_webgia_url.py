import requests

formats = [
    "https://webgia.com/ty-gia/vietinbank/15-02-2026.html",
    "https://webgia.com/ty-gia/vietinbank/ngay-15-02-2026.html",
    "https://webgia.com/ty-gia/vietinbank/lich-su-15-02-2026.html",
    "https://webgia.com/ty-gia/vietinbank/lich-su-ngay-15-02-2026.html",
    "https://webgia.com/ty-gia/vietinbank/history-15-02-2026.html",
    "https://webgia.com/ty-gia/vietinbank/2026-02-15.html",
]

headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

for url in formats:
    try:
        r = requests.head(url, headers=headers, timeout=5)
        print(f"URL: {url} -> Status: {r.status_code}")
    except Exception as e:
        print(f"URL: {url} -> Error: {e}")
