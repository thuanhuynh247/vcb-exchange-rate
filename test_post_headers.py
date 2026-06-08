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
            if "ty-gia-khcn" in url:
                print(f"Request: {request.method} {url}")
                print("  Headers:", dict(request.headers))
                if request.post_data:
                    print("  Post Data:", request.post_data[:500])
                
        def handle_response(response):
            url = response.url
            if "ty-gia-khcn" in url:
                print(f"Response: {response.status} {url}")
                # We can't easily print rsc response body here if it's binary or stream, but we can print Content-Type
                print("  Response Headers:", dict(response.headers))

        page.on("request", handle_request)
        page.on("response", handle_response)
        
        print("Navigating to VietinBank page...")
        await page.goto("https://www.vietinbank.vn/ty-gia-khcn", wait_until="load", timeout=30000)
        await page.wait_for_timeout(5000)
        
        # Click date input to open datepicker
        date_input = await page.query_selector('input[placeholder="DD/MM/YYYY"]')
        if date_input:
            await date_input.click()
            await page.wait_for_timeout(1000)
            
            # Let's try to click on the calendar days
            # We want to click on a day. Let's look for calendar day elements
            # Common classes for react-datepicker or flatpickr:
            # .react-datepicker__day or .flatpickr-day or similar
            days = await page.query_selector_all(".react-datepicker__day, .flatpickr-day, .day")
            print(f"Found {len(days)} calendar day elements.")
            if days:
                print("Clicking first day...")
                await days[0].click()
                await page.wait_for_timeout(3000)
            else:
                # Let's inspect the page DOM for classes containing calendar or day
                # We can query all elements that have text inside the calendar popup
                pass
                
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
