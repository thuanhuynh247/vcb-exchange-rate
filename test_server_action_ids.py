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
        
        actions = []
        
        async def handle_response(response):
            request = response.request
            if request.method == "POST" and "ty-gia-khcn" in request.url:
                action_id = request.headers.get("next-action") or request.headers.get("Next-Action")
                payload = request.post_data
                try:
                    body = await response.text()
                except Exception as e:
                    body = f"Error reading body: {e}"
                actions.append({
                    "action_id": action_id,
                    "payload": payload,
                    "status": response.status,
                    "body": body
                })

        page.on("response", lambda r: asyncio.ensure_future(handle_response(r)))
        
        print("Navigating to VietinBank page...")
        await page.goto("https://www.vietinbank.vn/ty-gia-khcn", wait_until="load", timeout=30000)
        await page.wait_for_timeout(6000)
        
        print(f"\nCaptured {len(actions)} POST requests:")
        for idx, act in enumerate(actions):
            print(f"\n--- ACTION {idx} ---")
            print(f"Action ID: {act['action_id']}")
            print(f"Payload: {act['payload']}")
            print(f"Response Status: {act['status']}")
            print(f"Response Body (truncated 1000): {act['body'][:1000]}")
            
        await browser.close()

if __name__ == '__main__':
    asyncio.run(main())
