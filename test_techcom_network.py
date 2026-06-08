import asyncio
import os

os.environ['PLAYWRIGHT_BROWSERS_PATH'] = r'D:\ms-playwright'
os.environ['TEMP'] = r'D:\temp'
os.environ['TMP'] = r'D:\temp'

async def main():
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        
        def handle_request(request):
            url = request.url
            if any(k in url.lower() for k in ["rate", "ex", "api", "ty-gia", "json", "graphql"]):
                print(f"Request: {request.method} {url}")
                if request.post_data:
                    print(f"  Post Data: {request.post_data[:300]}")
                
        def handle_response(response):
            url = response.url
            if any(k in url.lower() for k in ["rate", "ex", "api", "ty-gia", "json", "graphql"]):
                print(f"Response: {response.status} {url}")
                
        page.on("request", handle_request)
        page.on("response", handle_response)
        
        print("Navigating to Techcombank page...")
        await page.goto("https://techcombank.com/cong-cu-tien-ich/ty-gia", wait_until="load", timeout=35000)
        await page.wait_for_timeout(7000)
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
