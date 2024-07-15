import asyncio
from playwright.async_api import async_playwright
import json
import datetime
from datetime import datetime, date
import openpyxl
import os
import re
from openpyxl.styles import Alignment

today = date.today()
formatted_today = today.strftime("%d/%m/%Y")

async def get_accessibility_tree(url):
    async with async_playwright() as p:
        browser = await p.firefox.launch_persistent_context(user_data_dir="C:\\Users\\phams\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\bd5r7ma0.default-release - Copy", headless=True)
        page = await browser.new_page()
        await page.goto(url)
        
        # Get the accessibility tree
        accessibility_snapshot = await page.accessibility.snapshot()
        
        with open('C:\\Users\\Public\\Downloads\\nghi_dinh.json', 'w') as f:
            json.dump(accessibility_snapshot, f, indent=2)
        
        await browser.close()
        
        # # Get the page source
        # page_source = await page.content()
        
        # # Save the page source to an HTML file
        # with open('C:\\Users\\phams\\Downloads\\page_source.html', 'w', encoding='utf-8') as f:
        #     f.write(page_source)
        
        # await browser.close()

# URL to open
url = 'https://danhmuchanhchinh.gso.gov.vn/NghiDinh.aspx'

# Run the function
asyncio.get_event_loop().run_until_complete(get_accessibility_tree(url))

json_path = 'C:\\Users\\Public\\Downloads\\nghi_dinh.json'

# Attempt to open the file with 'utf-8' encoding
try:
    with open(json_path, 'r', encoding='utf-8') as file:
        data = json.load(file)
except UnicodeDecodeError:
    # If 'utf-8' encoding fails, try 'iso-8859-1'
    with open(json_path, 'r', encoding='iso-8859-1') as file:
        data = json.load(file)

new_decrees = []
count = 0

if len(data['children']) != 0:
    decrees = data['children'][-2]['children']
    total = int(len(decrees) / 4)
    
    for i in range(total):
        if datetime.strptime(decrees[1 + (4 * i)]['name'], "%d/%m/%Y").date() >= datetime.strptime(formatted_today, "%d/%m/%Y").date():
            new = {
                'nghi_dinh_so': decrees[0 + (4 * i)]['name'],
                'ngay_ban_hanh': decrees[1 + (4 * i)]['name'],
                'ngay_hieu_luc': decrees[2 + (4 * i)]['name'],
                'noi_dung': decrees[3 + (4 * i)]['name']
            }
            new_decrees.append(new)
            count += 1

message = f'Automated message: Ngày {formatted_today} có {count} nghị định mới'

new_decrees_info = []

if count > 0:
    for data in new_decrees:
        template = '\nNghị định số: {nghi_dinh_so}, Ngày ban hành: {ngay_ban_hanh}, Ngày hiệu lực: {ngay_hieu_luc}, Nội dung: {noi_dung}'
        result = f'{template.format(**data)}'
        new_decrees_info.append(result)

if len(new_decrees_info) != 0:
    new_decrees_info_merge = ''.join(new_decrees_info)
else:
    new_decrees_info_merge = ''

message = message + new_decrees_info_merge + ''

async def send_message_test():
    async with async_playwright() as p:
        browser = await p.firefox.launch_persistent_context(user_data_dir="C:\\Users\\phams\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\bd5r7ma0.default-release - Copy", headless=False)
        page = await browser.new_page()
        
        # # Navigate to site
        await page.goto('https://ghtk.me/channel/internal/recent/6997756873994439959')
        
        # Wait for the chat box to load
        await page.wait_for_selector('div[contenteditable="true"]')
        
        # Find the chat box and type a message
        chat_box = await page.query_selector('div[contenteditable="true"]')
        for char in message:
            await chat_box.type(char, delay=50)
        
        # Press Enter to send the message
        async def click_send_button():
            await page.wait_for_selector('[class="footer-view__rep"]', timeout=20000)
            send_button = await page.query_selector('[class="footer-view__rep"]')  # Replace with the actual selector for the div acting as a button
            if send_button:
                await send_button.click()
                return True
            else:
                print('Send button not found.')
                return False
        
        while not await click_send_button():
            await asyncio.sleep(2)

        # Close the browser (optional)
        await browser.close()

# Run the asyncio event loop
asyncio.get_event_loop().run_until_complete(send_message_test())