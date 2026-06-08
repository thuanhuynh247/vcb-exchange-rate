import requests
from bs4 import BeautifulSoup

urls = {
    "current": "https://www.vietinbank.vn/ty-gia-khcn",
    "query_ddmmyyyy": "https://www.vietinbank.vn/ty-gia-khcn?date=02/01/2026",
    "query_yyyymmdd": "https://www.vietinbank.vn/ty-gia-khcn?date=2026-01-02",
    "query_other": "https://www.vietinbank.vn/ty-gia-khcn?date=05/01/2026"
}
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

for key, url in urls.items():
    r = requests.get(url, headers=headers, verify=False)
    soup = BeautifulSoup(r.text, 'html.parser')
    
    # Let's find USD rates in the table
    # We will search for all table rows and find rows containing "USD"
    usd_rates = []
    for row in soup.find_all('tr'):
        text = row.text
        if "USD" in text:
            cells = [c.text.strip().replace('\n', ' ') for c in row.find_all(['td', 'th'])]
            usd_rates.append(cells)
            
    print(f"\n{key}: {url}")
    print(f"USD rate rows found: {len(usd_rates)}")
    for rate in usd_rates:
        print(" ", rate)
