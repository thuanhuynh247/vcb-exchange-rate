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
        await page.goto("https://www.vietinbank.vn/ty-gia-khcn", wait_until="load", timeout=30000)
        await page.wait_for_timeout(4000)
        
        # Find all select/option elements
        selects = await page.query_selector_all("select")
        print(f"Found {len(selects)} selects:")
        for idx, sel in enumerate(selects):
            name = await sel.get_attribute("name")
            id_attr = await sel.get_attribute("id")
            print(f"  Select {idx}: name='{name}', id='{id_attr}'")
            options = await sel.query_selector_all("option")
            for opt in options:
                val = await opt.get_attribute("value")
                txt = await opt.inner_text()
                print(f"    Option: value='{val}', text='{txt.strip()}'")
                
        # Find other custom dropdowns (like div/button containing select or options)
        buttons = await page.query_selector_all("button")
        print(f"\nFound {len(buttons)} buttons:")
        for idx, btn in enumerate(buttons):
            txt = await btn.inner_text()
            cls = await btn.get_attribute("class")
            if txt.strip() and len(txt.strip()) < 50:
                print(f"  Button {idx}: class='{cls}', text='{txt.strip()}'")
                
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
