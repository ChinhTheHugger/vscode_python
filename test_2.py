import asyncio
from playwright.async_api import async_playwright
import openpyxl
from bs4 import BeautifulSoup
import re
from docx import Document
import os
import time
import datetime
import json
from pyppeteer import launch

async def send_message_test():
    async with async_playwright() as p:
        browser = await p.firefox.launch_persistent_context(user_data_dir="C:\\Users\\phams\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\bd5r7ma0.default-release - Copy", headless=False)
        page = await browser.new_page()

        # # Navigate to Discord
        # await page.goto('https://discord.com/channels/683762358799302756/692444131573891176')
        
        # # Navigate to site
        await page.goto('https://ghtk.me/channel/internal/recent/6997756873994439959')

        # # Wait for the chat box to load - Discord
        # await page.wait_for_selector('div[role="textbox"]', timeout=30000)
        
        # Wait for the chat box to load
        await page.wait_for_selector('div[contenteditable="true"]')
        
        # Get current date and time
        current_time = datetime.datetime.now()

        # # Find the chat box and type a message - Discord
        # chat_box = await page.query_selector('div[role="textbox"]')
        # message = f"Texting automated message with Python script. If you see this, that means the infamous Miqo'te Manstealer has stolen yet another man, at {current_time} GMT +7"
        # for char in message:
        #     await chat_box.type(char, delay=50)
        
        # Find the chat box and type a message
        chat_box = await page.query_selector('div[contenteditable="true"]')
        message = f"Testing automated message - version 3"
        for char in message:
            await chat_box.type(char, delay=50)

        # # Press Enter to send the message - Discord
        # await chat_box.press('Enter')
        
        # Press Enter to send the message
        async def click_send_button():
            await page.wait_for_selector('[class="footer-view__rep"]', timeout=10000)
            send_button = await page.query_selector('[class="footer-view__rep"]')  # Replace with the actual selector for the div acting as a button
            if send_button:
                await send_button.click()
                return True
            else:
                print('Send button not found.')
                return False
        # Retry mechanism
        max_retries = 5
        while not await click_send_button():
            await asyncio.sleep(2)

        # Close the browser (optional)
        await browser.close()

# Run the asyncio event loop
asyncio.get_event_loop().run_until_complete(send_message_test())