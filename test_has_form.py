import requests
from bs4 import BeautifulSoup

url = "https://webgia.com/ty-gia/vietinbank/"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
r = requests.get(url, headers=headers)
soup = BeautifulSoup(r.text, 'html.parser')
forms = soup.find_all('form')
print(f"Found {len(forms)} forms:")
for i, f in enumerate(forms):
    print(f"Form {i}: attributes={f.attrs}")
    print(f.prettify()[:500])
