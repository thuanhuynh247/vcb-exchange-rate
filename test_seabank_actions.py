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
        
        async def handle_request(request):
            url = request.url
            if "seabank" in url and request.method == "POST":
                action_id = request.headers.get("next-action") or request.headers.get("Next-Action")
                print(f"\n[POST] {url}")
                print(f"  Next-Action: {action_id}")
                print(f"  Post Data: {request.post_data}")
                
        async def handle_response(response):
            url = response.url
            request = response.request
            if "seabank" in url and request.method == "POST":
                action_id = request.headers.get("next-action") or request.headers.get("Next-Action")
                if action_id:
                    print(f"\n[RESPONSE] Action {action_id}")
                    try:
                        text = await response.text()
                        print(f"  Body: {text[:1000]}")
                    except Exception as e:
                        print(f"  Error: {e}")
        
        page.on("request", lambda r: asyncio.ensure_future(handle_request(r)))
        page.on("response", lambda r: asyncio.ensure_future(handle_response(r)))
        
        print("Navigating to SeaBank page...")
        await page.goto("https://www.seabank.com.vn/cong-cu-tien-ich/ty-gia", wait_until="load", timeout=35000)
        await page.wait_for_timeout(8000)
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
