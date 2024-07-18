import asyncio
from playwright.async_api import async_playwright
import openpyxl
from bs4 import BeautifulSoup
import re
from docx import Document
import os
import time
import pickle
import ast
import pandas

provinces = [
    'an_giang','ba_ria_vung_tau','bac_kan','bac_lieu','bac_ninh','ben_tre','binh_dinh','binh_phuoc','binh_thuan','ca_mau','can_tho','cao_bang','da_nang','dak_lak','dak_nong','dien_bien','dong_nai','dong_thap',
    'gia_lai','ha_giang','ha_nam','ha_tinh','hai_duong','hau_giang','hoa_binh','hung_yen','khanh_hoa','kien_giang','kon_tum','lai_chau','lam_dong','lang_son','lao_cai','long_an','nam_dinh','nghe_an','ninh_binh',
    'ninh_thuan','phu_tho','phu_yen','quang_binh','quang_nam','quang_ngai','quang_tri','soc_trang','son_la','tay_ninh','thai_binh','thai_binh_duong','thai_nguyen','thanh_hoa','thua_thien_hue','tien_giang','tp_hcm',
    'tra_vinh','tuyen_quang','vinh_long','vinh_phuc','yen_bai'
]

df_list = []

for idx, obj in enumerate(provinces):
    name = obj
    
    parts = name.split('_')
    capitalized_parts = [part.capitalize() for part in parts]
    
    capitalized_name = '_'.join(capitalized_parts)
    
    df = pandas.read_excel(f'C:\\Users\\phams\\Downloads\\du an\\du lieu goc\\{capitalized_name}\\{obj}_projects_data.xlsx')
    
    df_list.append(df)
    
    print(f'Merged data for {capitalized_name}')

# Concatenate all DataFrames into one
combined_df = pandas.concat(df_list, ignore_index=True)

file_path = f'C:\\Users\\phams\\Downloads\\du an\\tong hop\\tong_hop_du_an.xlsx'

# Save the combined DataFrame to a single Excel file
combined_df.to_excel(file_path, index=True)

spreadsheet = openpyxl.load_workbook(file_path)
sheet = spreadsheet.active

column_widths_in_inches = {'A':0.5,'B':1.5,'C':3,'D':2,'E':1.25,'F':1.25,'G':2.75,'H':2,'I':1.5,'J':2,'K':1.25,'L':1.25,'M':1.25,'N':2,'O':1.75,'P':1}

for col, width in column_widths_in_inches.items():
    # Convert inches to Excel column width units
    excel_column_width = width * 10.71
    sheet.column_dimensions[col].width = excel_column_width

# Save the modified workbook
spreadsheet.save(file_path)