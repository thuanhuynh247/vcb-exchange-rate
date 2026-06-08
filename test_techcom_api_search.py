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
        
        async def handle_response(response):
            url = response.url
            if "techcombank.com" in url:
                content_type = response.headers.get("content-type", "")
                if "json" in content_type or "graphql" in url:
                    print(f"\nResponse JSON/GraphQL: {response.status} {url}")
                    try:
                        text = await response.text()
                        print(f"  Body (first 1000 chars): {text[:1000]}")
                    except Exception as e:
                        print(f"  Error reading body: {e}")

        page.on("response", lambda r: asyncio.ensure_future(handle_response(r)))
        
        print("Navigating to Techcombank page...")
        await page.goto("https://techcombank.com/cong-cu-tien-ich/ty-gia", wait_until="load", timeout=35000)
        await page.wait_for_timeout(8000)
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
