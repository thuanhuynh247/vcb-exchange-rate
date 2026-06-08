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
        
        # Intercept and log network requests triggered by click
        def handle_request(request):
            url = request.url
            # Log any requests to vietinbank.vn or those that look like APIs
            if "vietinbank.vn" in url and not any(ext in url for ext in [".js", ".css", ".png", ".svg", ".jpg", "_next/image"]):
                print(f"Request: {request.method} {url}")
                if request.post_data:
                    print(f"  Post Data: {request.post_data[:500]}")
                
        page.on("request", handle_request)
        
        print("Navigating to VietinBank page...")
        await page.goto("https://www.vietinbank.vn/ty-gia-khcn", wait_until="load", timeout=30000)
        await page.wait_for_timeout(4000)
        
        print("Current USD Rates:")
        rows = await page.query_selector_all("table tbody tr")
        for row in rows:
            text = await row.inner_text()
            if "USD" in text:
                print("  Current:", text.replace('\n', ' '))
        
        date_input = await page.query_selector('input[placeholder="DD/MM/YYYY"]')
        if date_input:
            print("Clicking date input...")
            await date_input.click()
            await page.wait_for_timeout(1500)
            
            # Let's capture the page DOM or screenshot to see what datepicker is open
            # We can print all elements with classes containing "calendar" or "datepicker" or "popup"
            popups = await page.query_selector_all("[class*='calendar'], [class*='datepicker'], [class*='popup']")
            print(f"Found {len(popups)} calendar-related elements.")
            for i, pop in enumerate(popups[:5]):
                cls = await pop.get_attribute("class")
                text = await pop.inner_text()
                print(f"  Popup {i}: class='{cls}', text snippet: '{text.strip()[:100]}'")
                
            # Let's try to find and click a day. In react-datepicker, days have classes like react-datepicker__day
            # Or maybe we can find a text button or text day, e.g. "1" or "2"
            # Let's look for elements with text "15" or similar inside a calendar container
            # We can use xpath or text-based query selectors
            day_el = await page.query_selector("xpath=//*[contains(@class, 'datepicker') or contains(@class, 'calendar')]//*[text()='15' or text()=' 15 ']")
            if day_el:
                print("Found day element, clicking...")
                await day_el.click()
                await page.wait_for_timeout(3000)
                
                # Check input value now
                val = await date_input.evaluate("el => el.value")
                print("New input value:", val)
                
                # Print new rates
                print("New USD Rates:")
                rows = await page.query_selector_all("table tbody tr")
                for row in rows:
                    text = await row.inner_text()
                    if "USD" in text:
                        print("  New:", text.replace('\n', ' '))
            else:
                print("Day element not found.")
        else:
            print("Date input not found.")
            
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
