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
        
        async def handle_request(request):
            url = request.url
            if request.method == "POST" and "ty-gia-khcn" in url:
                action_id = request.headers.get("next-action") or request.headers.get("Next-Action")
                print(f"\n[REQUEST] POST {url}")
                print(f"  Next-Action Header: {action_id}")
                print(f"  Headers: {dict(request.headers)}")
                print(f"  Post Data: {request.post_data}")
                
        async def handle_response(response):
            url = response.url
            request = response.request
            if request.method == "POST" and "ty-gia-khcn" in url:
                action_id = request.headers.get("next-action") or request.headers.get("Next-Action")
                print(f"\n[RESPONSE] POST {url} for Action {action_id}")
                print(f"  Status: {response.status}")
                try:
                    text = await response.text()
                    print(f"  Body (first 500 chars): {text[:500]}")
                except Exception as e:
                    print(f"  Could not read body: {e}")

        page.on("request", lambda r: asyncio.ensure_future(handle_request(r)))
        page.on("response", lambda r: asyncio.ensure_future(handle_response(r)))
        
        print("Navigating to VietinBank page...")
        await page.goto("https://www.vietinbank.vn/ty-gia-khcn", wait_until="load", timeout=30000)
        await page.wait_for_timeout(5000)
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
