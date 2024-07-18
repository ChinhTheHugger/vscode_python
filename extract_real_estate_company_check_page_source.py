import asyncio
from playwright.async_api import async_playwright
import openpyxl
from bs4 import BeautifulSoup
import re
from docx import Document
import os
import time
import json

new_folder_path = 'C:\\Users\\phams\\Downloads\\doanh_nghiep\\thong tin doanh nghiep'
os.makedirs(new_folder_path, exist_ok=True)

types = [
    ('chu-dau-tu',10),
    ('nha-thau-tu-van',2),
    ('tong-thau-du-an',4),
    ('thau-co-dien-vat-tu-thi-cong',1),
    ('thau-hang-muc-xay-dung-vat-tu-thi-cong',1),
    ('dich-vu-logicstic',1),
    ('cong-nghe-va-thiet-bi-san-xuat',1)
]

# types = [('gian-hang',2)]

with open(f'C:\\Users\\phams\\Downloads\\doanh_nghiep\\page_sources\\chu-dau-tu_54.html', 'r', encoding='utf-8') as file:
    html_content = file.read()
    
# Parse the HTML content
soup = BeautifulSoup(html_content, 'html.parser')

checkpoint_text = None
detail_article = soup.find('div', class_='d-flex d-md-none row')
if detail_article:
    h1_tag = detail_article.find('h1')
    if h1_tag:
        checkpoint_text = h1_tag.get_text(separator='\n').strip()

arr = []
if checkpoint_text:
    checkpoint_text = str(checkpoint_text).replace('/','-')
    
    # Find the div with class 'special-class'
    special_div = soup.find('div', class_='col-12 col-md-8')

    # Extract all text inside that div
    if special_div:
        for text_element in special_div.stripped_strings:
            arr.append(text_element)
                    
    # Find the div with class 'special-class'
    special_div = soup.find('div', class_='col-12 col-md-4 mt-3')

    # Extract all text inside that div
    if special_div:
        for text_element in special_div.stripped_strings:
            arr.append(text_element)

print('\n\n'.join(arr))