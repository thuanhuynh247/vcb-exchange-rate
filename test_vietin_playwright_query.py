import asyncio
import os

os.environ['PLAYWRIGHT_BROWSERS_PATH'] = r'D:\ms-playwright'
os.environ['TEMP'] = r'D:\temp'
os.environ['TMP'] = r'D:\temp'

async def get_rates_for_url(url):
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        print(f"Loading URL: {url}")
        try:
            await page.goto(url, wait_until="load", timeout=30000)
            await page.wait_for_timeout(4000)
            
            # Print the text in the date input if found
            date_input = await page.query_selector('input[placeholder="DD/MM/YYYY"]')
            input_val = await date_input.evaluate("el => el.value") if date_input else "Not found"
            print("  Date Input Value in DOM:", input_val)
            
            # Extract USD rates from the table
            rows = await page.query_selector_all("table tbody tr")
            for row in rows:
                text = await row.inner_text()
                if "USD" in text:
                    cells = await row.query_selector_all("td")
                    if len(cells) >= 4:
                        print("  USD Row:", [await c.inner_text() for c in cells])
        except Exception as e:
            print(f"  Error loading URL: {e}")
        await browser.close()

async def main():
    urls = [
        "https://www.vietinbank.vn/ty-gia-khcn",
        "https://www.vietinbank.vn/ty-gia-khcn?date=02/01/2026",
        "https://www.vietinbank.vn/ty-gia-khcn?date=2026-01-02"
    ]
    for url in urls:
        await get_rates_for_url(url)

if __name__ == '__main__':
    asyncio.run(main())
