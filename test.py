import asyncio
from playwright.async_api import async_playwright
import json

async def get_accessibility_tree(url):
    async with async_playwright() as p:
        browser = await p.firefox.launch_persistent_context(user_data_dir="C:\\Users\\phams\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\bd5r7ma0.default-release - Copy", headless=True)
        page = await browser.new_page()
        await page.goto(url)
        
        # # Get the accessibility tree
        # accessibility_snapshot = await page.accessibility.snapshot()
        
        # with open('C:\\Users\\phams\\Downloads\\accessibility_tree.json', 'w') as f:
        #     json.dump(accessibility_snapshot, f, indent=2)
        
        # await browser.close()
        
        # Get the page source
        page_source = await page.content()
        
        # Save the page source to an HTML file
        with open('C:\\Users\\phams\\Downloads\\page_source.html', 'w', encoding='utf-8') as f:
            f.write(page_source)
        
        await browser.close()

# URL to open
url = 'https://fcivietnam.com/du-an/du-an-nha-may/2024/10/du-an-nha-may-cong-ty-tnhh-cong-nghe-jy-viet-nam'

# Run the function
asyncio.get_event_loop().run_until_complete(get_accessibility_tree(url))

