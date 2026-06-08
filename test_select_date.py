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
            if any(k in url.lower() for k in ["rate", "ex", "api", "ty-gia", "json", "rsc"]):
                print(f"Request: {request.method} {url}")
                
        page.on("request", handle_request)
        
        print("Navigating to VietinBank page...")
        await page.goto("https://www.vietinbank.vn/ty-gia-khcn", wait_until="load", timeout=30000)
        await page.wait_for_timeout(4000)
        
        # Find the input with placeholder "DD/MM/YYYY"
        print("Finding input...")
        date_input = await page.query_selector('input[placeholder="DD/MM/YYYY"]')
        if date_input:
            print("Found input, clicking and typing...")
            # Click and select all to overwrite
            await date_input.click()
            await page.keyboard.press("Control+A")
            await page.keyboard.press("Backspace")
            await date_input.type("02/01/2026")
            await page.keyboard.press("Enter")
            
            print("Typed date, waiting to see if requests are triggered...")
            await page.wait_for_timeout(5000)
            
            # Let's read the value of the input
            val = await date_input.evaluate("el => el.value")
            print("Current input value:", val)
        else:
            print("Input not found!")
            
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
