import asyncio
from playwright.async_api import async_playwright
import openpyxl
from bs4 import BeautifulSoup
import re
from docx import Document
import os
import time
import pickle

with open("C:\\Users\\phams\\Downloads\\du an\\du lieu goc\\Bac_Lieu\\page sources\\page_source_6.html", 'r', encoding='utf-8') as file:
    html_content = file.read()

soup = BeautifulSoup(html_content, 'html.parser')

detail_article = soup.find('article', class_='detail')
if detail_article:
    h1_tag = detail_article.find('h1')
    if h1_tag:
        print(h1_tag.get_text(separator='\n').strip())
        print('---------------------------------------------------------------------------')

article_str = ''
for tag in soup.find_all('article'):
    # print(tag.get_text(separator='\n').strip())
    # print('---------------------------------------------------------------------------')
    string = str(tag.get_text(separator='\n').strip())
    arr = string.split('\n')
    filtered_list = [item.strip() for item in arr if item != '' and item != '\xa0' and item != ' ' and item.strip() != ':'and item.strip() != '']
    print(filtered_list)
    # for a in filtered_list:
    #     print(a)
    print(len(filtered_list))
    article_str = '&'.join(filtered_list)
    print('---------------------------------------------------------------------------')

tbody_str  =''
for tag in soup.find_all('tbody'):
    # print(tag.get_text(separator='\n').strip())
    # print('---------------------------------------------------------------------------')
    string = str(tag.get_text(separator='\n').strip())
    arr = string.split('\n')
    filtered_list = [item.strip() for item in arr if item != '' and item != '\xa0' and item != ' ' and item.strip() != ':']
    print(filtered_list)
    # for a in filtered_list:
    #     print(a)
    print(len(filtered_list))
    tbody_str = '&'.join(filtered_list)
    print('---------------------------------------------------------------------------')

if tbody_str in article_str:
    print('YES')
else:
    print('NO')