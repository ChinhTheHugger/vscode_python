import asyncio
from playwright.async_api import async_playwright
import time

async def get_accessibility_tree(url_list):
    async with async_playwright() as p:
        browser = await p.firefox.launch_persistent_context(user_data_dir="C:\\Users\\phams\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\bd5r7ma0.default-release - Copy", headless=False)
        page = await browser.new_page()
        
        for url in url_list:
            await page.goto(url)
            time.sleep(3)
            
        # page_source = await page.content()
            
        # source_file_path = 'path/to/save/source/file'
        
        # with open(source_file_path, 'w', encoding='utf-8') as f:
        #     f.write(page_source)
            
        await browser.close()

url_list = ['https://www.google.com','https://www.youtube.com','https://www.discord.com']
        
asyncio.get_event_loop().run_until_complete(get_accessibility_tree(url_list))