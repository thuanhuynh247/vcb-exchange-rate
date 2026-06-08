import requests
from bs4 import BeautifulSoup

url = "https://webgia.com/ty-gia/vietinbank/"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
r = requests.get(url, headers=headers)
soup = BeautifulSoup(r.text, 'html.parser')
form = soup.find('form', class_='form-inline')
if form:
    print("Form attributes:", form.attrs)
    for inp in form.find_all(['input', 'button', 'select', 'textarea']):
        print(f"Name: {inp.get('name')}, Type: {inp.get('type')}, Value: {inp.get('value')}, Text: {inp.text.strip()}")
else:
    print("Form not found")
