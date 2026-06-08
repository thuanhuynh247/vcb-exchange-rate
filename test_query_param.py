import requests
from bs4 import BeautifulSoup

urls = {
    "current": "https://webgia.com/ty-gia/vietinbank/",
    "slash_date": "https://webgia.com/ty-gia/vietinbank/?date=02/01/2026",
    "dash_date": "https://webgia.com/ty-gia/vietinbank/?date=02-01-2026",
    "iso_date": "https://webgia.com/ty-gia/vietinbank/?date=2026-01-02",
}
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

for key, url in urls.items():
    r = requests.get(url, headers=headers)
    soup = BeautifulSoup(r.text, 'html.parser')
    usd = soup.find(id="wdg-usd-m")
    usd_val = usd.text.strip() if usd else "Not found"
    
    # Also find date indicator in the page text to see if the page claims to be for that date
    # Let's search for text containing "02/01/2026" or "02-01-2026"
    has_date_text = "02/01/2026" in r.text or "02-01-2026" in r.text or "2-1-2026" in r.text
    
    print(f"{key}: USD = {usd_val}, contains historical date string: {has_date_text}")
