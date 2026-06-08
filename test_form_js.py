import requests
from bs4 import BeautifulSoup

url = "https://webgia.com/ty-gia/vietcombank/"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
r = requests.get(url, headers=headers)
soup = BeautifulSoup(r.text, 'html.parser')
input_tag = soup.find(id="hcalendar")
if input_tag:
    parent = input_tag.find_parent()
    while parent:
        if parent.name == 'form':
            print("Found Form:", parent.attrs)
            print(parent.prettify()[:1000])
            break
        parent = parent.find_parent()
else:
    print("Not found hcalendar")
