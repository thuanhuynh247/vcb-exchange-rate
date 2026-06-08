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
        
        # Listen for request and response events
        def handle_request(request):
            url = request.url
            if any(k in url.lower() for k in ["rate", "ex", "api", "ty-gia", "json"]):
                print(f"Request: {request.method} {url}")
                
        def handle_response(response):
            url = response.url
            if any(k in url.lower() for k in ["rate", "ex", "api", "ty-gia", "json"]):
                print(f"Response: {response.status} {url}")
                
        page.on("request", handle_request)
        page.on("response", handle_response)
        
        print("Navigating to VietinBank page...")
        await page.goto("https://www.vietinbank.vn/ty-gia-khcn", wait_until="load", timeout=30000)
        await page.wait_for_timeout(5000)
        print("Done waiting.")
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
