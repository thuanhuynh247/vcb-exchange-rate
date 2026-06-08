import asyncio
import os
import json

os.environ['PLAYWRIGHT_BROWSERS_PATH'] = r'D:\ms-playwright'
os.environ['TEMP'] = r'D:\temp'
os.environ['TMP'] = r'D:\temp'

async def main():
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        
        # Intercept ALL requests from techcombank
        api_requests = []
        
        async def handle_request(request):
            url = request.url
            if "techcombank.com" in url and not any(ext in url for ext in [".js", ".css", ".png", ".svg", ".jpg", ".woff", ".ttf", ".ico"]):
                api_requests.append({
                    "method": request.method,
                    "url": url,
                    "post_data": request.post_data
                })
                
        page.on("request", lambda r: asyncio.ensure_future(handle_request(r)))
        
        print("Navigating to Techcombank page...")
        await page.goto("https://techcombank.com/cong-cu-tien-ich/ty-gia", wait_until="load", timeout=35000)
        await page.wait_for_timeout(8000)
        
        # Try clicking a date picker or date input
        date_inputs = await page.query_selector_all("input[type='date'], input[placeholder], .date-picker, [class*='date']")
        print(f"\nFound {len(date_inputs)} date-related elements")
        
        # Try looking for any button to change date
        all_elements = await page.query_selector_all("[data-date], [class*='prev'], [class*='next'], [class*='calendar']")
        print(f"Found {len(all_elements)} navigation elements")
        
        print(f"\nTotal non-static API requests captured: {len(api_requests)}")
        for req in api_requests:
            print(f"\n{req['method']} {req['url']}")
            if req['post_data']:
                print(f"  Body: {req['post_data'][:200]}")
                
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
