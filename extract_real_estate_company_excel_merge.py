import asyncio
from playwright.async_api import async_playwright
import openpyxl
from bs4 import BeautifulSoup
import re
from docx import Document
import os
import time
import json
import pandas

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

types_new = ['Chủ đầu tư','Nhà thầu tư vấn','Tổng thầu dự án','Thầu cơ điện vật tư thi công','Thầu hạng mục xây dựng vật tư thi công','Dịch vụ Logistic','Công nghệ và thiết bị sản xuất']

# types = [('gian-hang',2)]

df_list = []

for idx, obj in enumerate(types):
    df = pandas.read_excel(f'C:\\Users\\phams\\Downloads\\doanh_nghiep\\tong_hop_doanh_nghiep_{obj[0]}.xlsx')
    df.dropna(how='all', inplace=True)
    df_list.append(df)

# Concatenate all DataFrames into one
combined_df = pandas.concat(df_list, ignore_index=True)

combined_df.drop_duplicates()
combined_df.dropna(subset='Tên công ty')

file_path = f'C:\\Users\\phams\\Downloads\\doanh_nghiep\\tong_hop_doanh_nghiep.xlsx'

# Save the combined DataFrame to a single Excel file
combined_df.to_excel(file_path, index=True)

spreadsheet = openpyxl.load_workbook(file_path)
sheet = spreadsheet.active

column_widths_in_inches = {'A':0.5,'B':2,'C':3.25,'D':3,'E':3,'F':1.5,'G':2.75,'H':2.75,'I':2,'J':2,'K':2,'L':2,'M':2,'N':2,'O':2,'P':1}

for col, width in column_widths_in_inches.items():
    # Convert inches to Excel column width units
    excel_column_width = width * 10.71
    sheet.column_dimensions[col].width = excel_column_width
    
# Save the modified workbook
spreadsheet.save(file_path)