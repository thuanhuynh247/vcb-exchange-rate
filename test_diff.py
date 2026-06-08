import requests
from bs4 import BeautifulSoup

url = "https://webgia.com/ty-gia/vietinbank/"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

# 1. GET request (current rates)
r_get = requests.get(url, headers=headers)
soup_get = BeautifulSoup(r_get.text, 'html.parser')
usd_get = soup_get.find(id="wdg-usd-m")
usd_get_text = usd_get.text.strip() if usd_get else "Not found"

# 2. POST request with historical date (e.g. 02/01/2026)
r_post = requests.post(url, headers=headers, data={"date": "02/01/2026"})
soup_post = BeautifulSoup(r_post.text, 'html.parser')
usd_post = soup_post.find(id="wdg-usd-m")
usd_post_text = usd_post.text.strip() if usd_post else "Not found"

print("GET USD rate:", usd_get_text)
print("POST (02/01/2026) USD rate:", usd_post_text)

# Let's print the date display element in the page if any
# e.g., the value of the datepicker or date header
date_post = soup_post.find(id="hcalendar")
print("POST date picker value:", date_post.get('value') if date_post else "Not found")
