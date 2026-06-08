import requests
import re

url = "https://www.vietinbank.vn/ty-gia-khcn"
headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
r = requests.get(url, headers=headers, verify=False)
print("Status Code:", r.status_code)
print("Total HTML Length:", len(r.text))

# Let's search for script tags, datepicker elements, forms, etc.
scripts = re.findall(r'<script[^>]*src="([^"]+)"', r.text)
print("\nJS Imports:")
for s in scripts[:10]:
    print(" ", s)

# Check if there is an input or datepicker or calendar element
matches = re.findall(r'<input[^>]*>', r.text, re.IGNORECASE)
print("\nInputs:")
for m in matches:
    print(" ", m)
    
# Let's also check for form elements
forms = re.findall(r'<form[^>]*>', r.text, re.IGNORECASE)
print("\nForms:")
for f in forms:
    print(" ", f)
