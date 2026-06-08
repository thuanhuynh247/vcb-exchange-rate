import requests
from bs4 import BeautifulSoup

url = "https://webgia.com/ty-gia/vietcombank/"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
r = requests.get(url, headers=headers)
soup = BeautifulSoup(r.text, 'html.parser')
for script in soup.find_all('script'):
    content = script.string or ""
    if "date" in content or "location" in content or "submit" in content or "redirect" in content:
        print("Script:")
        print(content)
        print("-" * 50)
