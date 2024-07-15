import asyncio
from playwright.async_api import async_playwright
import openpyxl
from bs4 import BeautifulSoup
import re
from docx import Document
import os
import time
import json

new_folder_path = 'C:\\Users\\phams\\Downloads\\doanh_nghiep\\page_sources'
os.makedirs(new_folder_path, exist_ok=True)

# types = [
#     ('chu-dau-tu',10),
#     ('nha-thau-tu-van',2),
#     ('tong-thau-du-an',4),
#     ('thau-co-dien-vat-tu-thi-cong',1),
#     ('thau-hang-muc-xay-dung-vat-tu-thi-cong',1),
#     ('dich-vu-logicstic',1),
#     ('cong-nghe-va-thiet-bi-san-xuat',1)
# ]

types = [('gian-hang',2)]

for obj in types:
    spreadsheet = openpyxl.load_workbook(f'C:\\Users\\phams\\Downloads\\doanh_nghiep\\{obj[0]}.xlsx')
    sheet = spreadsheet.active
    
    async def get_accessibility_tree(url_list):
        async with async_playwright() as p:
            browser = await p.firefox.launch_persistent_context(user_data_dir="C:\\Users\\phams\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\bd5r7ma0.default-release - Copy", headless=True)
            page = await browser.new_page()
            
            for idx, url in enumerate(url_list):
                # # Get the accessibility tree
                # accessibility_snapshot = await page.accessibility.snapshot()
                
                # with open(f'C:\\Users\\phams\\Downloads\\doanh_nghiep\\page_sources\\{obj[0]}_{idx+1}.json', 'w') as f:
                #     json.dump(accessibility_snapshot, f, indent=2)
                    
                # Get the page source
                
                source_file_path = f'C:\\Users\\phams\\Downloads\\doanh_nghiep\\page_sources\\{obj[0]}_{idx+1}.html'
                
                if os.path.exists(source_file_path):
                    print(f'{obj[0]}_{idx+1}.html already exists')
                    continue
                else:
                    await page.goto(url,timeout=0)
                    
                    page_source = await page.content()
                    
                    # Save the page source to an HTML file
                    with open(source_file_path, 'w', encoding='utf-8') as f:
                        f.write(page_source)
                    
                    print(f'{obj[0]}: saved page_source_{idx+1}')
                
            await browser.close()
    
    url_list = [cell.value for cell in sheet['B'][1:]]
    
    asyncio.get_event_loop().run_until_complete(get_accessibility_tree(url_list))