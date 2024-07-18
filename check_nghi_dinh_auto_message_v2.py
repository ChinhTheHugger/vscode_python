import asyncio
from playwright.async_api import async_playwright
import json
from datetime import datetime, date, timedelta, timezone
import time
import logging
import signal

# Configure logging
logging.basicConfig(
    filename="LOG_nghi_dinh.txt",  # Replace with your log file path
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
                    # --- Start new Chrome instance
                    browser = await p.chromium.launch(headless=True)
                    
                    page = await browser.new_page()
                    
                    await page.goto('https://danhmuchanhchinh.gso.gov.vn/NghiDinh.aspx')
                    
                    # --- Get the accessibility tree (JSON file)
                    accessibility_snapshot = await page.accessibility.snapshot()
                    
                    # --- Close the browser
                    await browser.close()
                    
                    # --- Save the extracted JSON file to disk (Windows)
                    json_path = 'nghi_dinh.json'
                    
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

                    # --- Check for the "children" node in the JSON file, which contains info about decrees
                    if len(data['children']) != 0:
                        decrees = data['children']
                        total = int((len(decrees) - 20) / 4)
                        
                        for i in range(total):
                            # --- Check for new decree(s) based on current date and see if that decree(s) has been checked or not,
                            # --- to prevent duplicating within the same day
                            if datetime.strptime(decrees[17 + (4 * i)]['name'], "%d/%m/%Y").date() >= datetime.strptime(formatted_today, "%d/%m/%Y").date() and (decrees[16 + (4 * i)]['name'],formatted_today) not in checkpoint:
                                new = {
                                    'nghi_dinh_so': decrees[16 + (4 * i)]['name'],
                                    'ngay_ban_hanh': decrees[17 + (4 * i)]['name'],
                                    'ngay_hieu_luc': decrees[18 + (4 * i)]['name'],
                                    'noi_dung': decrees[19 + (4 * i)]['name']
                                }
                                new_decrees.append(new)
                                checkpoint.append((decrees[0 + (4 * i)]['name'],formatted_today))
                                count += 1

                    message = f'Automated message: Ngày {formatted_today} có {count} nghị định mới'

                    # --- Prepare to log the result for new decree(s)
                    if count > 0:
                        new_decrees_info = []
                        
                        # --- Set up new decree(s) info
                        for data in new_decrees:
                            template = '\nNghị định số: {nghi_dinh_so}, Ngày ban hành: {ngay_ban_hanh}, Ngày hiệu lực: {ngay_hieu_luc}, Nội dung: {noi_dung}'
                            result = f'{template.format(**data)}'
                            new_decrees_info.append(result)

                        new_decrees_info_merge = ''.join(new_decrees_info)

                        message = message + new_decrees_info_merge + ''
                        
                        logging.info(message)
                    else:
                        # --- Log the result for no new decree
                        logging.info(f': No new decree')

            # --- Run the process
            
            logging.info(f': Starting browser to check for new decrees...')
            
            asyncio.get_event_loop().run_until_complete(get_accessibility_tree())
            
            # --- Set the process to sleep for a set amount of time
            interval = 20 # seconds
            
            logging.info(f': Going into sleep mode for {interval / 3600} hours...')
            
            time.sleep(interval)
            
    except Exception as e:
        # --- Log any detected error
        logging.critical(f"Fatal error: {e}")
    finally:
        # --- Log the process termination
        logging.info("Monitoring stopped")