import requests

banks = ["vietcombank", "vietinbank", "techcombank", "seabank"]
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
}
data = {
    "date": "02/01/2026"
}

for bank in banks:
    url = f"https://webgia.com/ty-gia/{bank}/"
    r = requests.post(url, headers=headers, data=data, allow_redirects=False)
    print(f"Bank: {bank}")
    print("  POST Status Code:", r.status_code)
    if 'Location' in r.headers:
        loc = r.headers['Location']
        print("  Redirect Location:", loc)
        # Now check the redirected URL status code
        r2 = requests.get(loc, headers=headers)
        print("  GET Redirect Status Code:", r2.status_code)
    else:
        print("  No Location header found")
