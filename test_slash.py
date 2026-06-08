import requests

urls = [
    "https://webgia.com/ty-gia/vietinbank/02-01-2026/",
    "https://webgia.com/ty-gia/vietinbank/02-01-2026.html",
    "https://webgia.com/ty-gia/vietinbank/02-01-2026",
    "https://webgia.com/ty-gia/vietinbank/?date=02/01/2026",
    "https://webgia.com/ty-gia/vietinbank/?date=02-01-2026",
]
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
for url in urls:
    r = requests.get(url, headers=headers, allow_redirects=False)
    print(f"URL: {url} -> Status: {r.status_code}")
    if 'Location' in r.headers:
        print("  Redirects to:", r.headers['Location'])
