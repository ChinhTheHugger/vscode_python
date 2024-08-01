import asyncio
from playwright.async_api import async_playwright
import json
from datetime import datetime, date, timedelta, timezone
import time
import datetime as dt
import logging
import signal
import sys

# Configure logging
logging.basicConfig(
    filename="C:\\Users\\Public\\Downloads\\decrees_log.txt",  # Replace with your log file path
    level=logging.INFO,  # You can set this to DEBUG, INFO, WARNING, ERROR, CRITICAL
    format='%(asctime)s - %(levelname)s - %(message)s',
)

def signal_handler(sig, frame):
    logging.info("Interrupt received, shutting down...")
    for task in asyncio.all_tasks():
        task.cancel()
    asyncio.get_event_loop().stop()

if __name__ == "__main__":
    
    signal.signal(signal.SIGINT, signal_handler)
    
    last_checked = datetime.now(tz=timezone.utc)
    reset_time = (datetime.now(tz=timezone.utc) + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    
    try:
        checkpoint = []
        
        # --- Loop the process
        while True:
            today = date.today()
            formatted_today = today.strftime("%d/%m/%Y")
            
            # # --- For testing
            # formatted_today = '01/01/2024'
            
            now = datetime.now(tz=timezone.utc)
            
            if now >= reset_time:
                checkpoint = []
                reset_time = (now + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)

            async def get_accessibility_tree():
                async with async_playwright() as p:
                    # --- Start new Firefox instance
                    # --- user_data_dir points to the Firefox user profile with log in info for GHTK internal chat system
                    # --- To avoid conflicting with any running Firefox instance, use a different user profile containing the same log in info
                    browser = await p.firefox.launch_persistent_context(user_data_dir="C:\\Users\\phams\\AppData\\Local\\Mozilla\\Firefox\\Profiles\\bd5r7ma0.default-release - Copy", headless=False)
                    
                    # # --- Start new Chrome instance
                    # browser = await p.chromium.launch(headless=True)
                    
                    page = await browser.new_page()
                    
                    await page.goto('https://danhmuchanhchinh.gso.gov.vn/NghiDinh.aspx')
                    
                    # --- Get the accessibility tree (JSON file)
                    accessibility_snapshot = await page.accessibility.snapshot()
                    
                    # --- Save the extracted JSON file to disk (Windows)
                    json_path = 'C:\\Users\\Public\\Downloads\\nghi_dinh.json'
                    
                    with open(json_path, 'w') as f:
                        json.dump(accessibility_snapshot, f, indent=2)

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
                            if datetime.strptime(decrees[1 + (4 * i)]['name'], "%d/%m/%Y").date() >= datetime.strptime(formatted_today, "%d/%m/%Y").date() and (decrees[0 + (4 * i)]['name'],formatted_today) not in checkpoint:
                                new = {
                                    'nghi_dinh_so': decrees[0 + (4 * i)]['name'],
                                    'ngay_ban_hanh': decrees[1 + (4 * i)]['name'],
                                    'ngay_hieu_luc': decrees[2 + (4 * i)]['name'],
                                    'noi_dung': decrees[3 + (4 * i)]['name']
                                }
                                new_decrees.append(new)
                                checkpoint.append((decrees[0 + (4 * i)]['name'],formatted_today))
                                count += 1

                    message = f'Automated message: Ngày {formatted_today} có {count} nghị định mới'

                    # --- Prepare to log the result for new decree(s) and send the result to the desired GHTK group chat
                    if count > 0:
                        new_decrees_info = []
                        
                        # --- Set up new decree(s) info
                        for data in new_decrees:
                            template = '\nNghị định số: {nghi_dinh_so}, Ngày ban hành: {ngay_ban_hanh}, Ngày hiệu lực: {ngay_hieu_luc}, Nội dung: {noi_dung}'
                            result = f'{template.format(**data)}'
                            new_decrees_info.append(result)

                        new_decrees_info_merge = ''.join(new_decrees_info)

                        message = message + new_decrees_info_merge + ''
                        
                        # --- Navigate to the desired GHTK group chat using URL
                        await page.goto('https://GHTK/channel/link/here')
                        
                        # --- Wait for the chat box to load
                        await page.wait_for_selector('div[contenteditable="true"]')
                        
                        # --- Find the chat box and type a message
                        chat_box = await page.query_selector('div[contenteditable="true"]')
                        for char in message:
                            await chat_box.type(char, delay=50)
                        
                        # --- Press Enter to send the message
                        async def click_send_button():
                            await page.wait_for_selector('[class="footer-view__rep"]', timeout=20000)
                            send_button = await page.query_selector('[class="footer-view__rep"]')
                            if send_button:
                                await send_button.click()
                                return True
                            else:
                                return False
                                
                        while not await click_send_button():
                            await asyncio.sleep(2)

                        # --- Log the result for new decree(s)
                        logging.info(message)
                    else:
                        # --- Log the result for no new decree
                        logging.info(f'{dt.datetime.now()}: No new decree')
                    
                    await browser.close()

            # --- Run the process
            
            logging.info(f'{dt.datetime.now()}: Starting browser to check for new decrees...')
            
            asyncio.get_event_loop().run_until_complete(get_accessibility_tree())
            
            # --- Set the process to sleep for a set amount of time
            interval = 15 # seconds
            
            logging.info(f'{dt.datetime.now()}: Going into sleep mode for {interval / 60} minute(s)...')
            
            time.sleep(interval)
            
    except Exception as e:
        # --- Log any detected error
        logging.critical(f"Fatal error: {e}")
    finally:
        # --- Log the process termination
        logging.info("Monitoring stopped")