import asyncio
from playwright.async_api import async_playwright
import openpyxl
from bs4 import BeautifulSoup
import re
from docx import Document
import os
import time

json_list = [('phu_tho',1),('ha_giang',1),('cao_bang',1),('bac_kan',1),('tuyen_quang',1),('lao_cai',1),('dien_bien',1)
             ,('lai_chau',1),('son_la',1),('yen_bai',1),('hoa_binh',1),('thai_binh',2),('ha_nam',2),('nam_dinh',1),('ninh_binh',1),('thanh_hoa',2)
             ,('nghe_an',1),('ha_tinh',1),('quang_binh',1),('quang_tri',1),('thua_thien_hue',1),('da_nang',1),('quang_nam',1),('quang_ngai',1),('binh_dinh',1)
             ,('phu_yen',1),('khanh_hoa',1),('ninh_thuan',1),('binh_thuan',1),('kon_tum',1),('gia_lai',1),('dak_lak',1),('dak_nong',1),('lam_dong',1)
             ,('binh_phuoc',2),('tay_ninh',1),('thai_binh_duong',3),('dong_nai',2),('ba_ria_vung_tau',2),('tp_hcm',2),('long_an',2),('tien_giang',1),('ben_tre',1)
             ,('tra_vinh',1),('vinh_long',1),('dong_thap',1),('an_giang',1),('kien_giang',1),('can_tho',1),('hau_giang',1),('soc_trang',1),('bac_lieu',1),('ca_mau',1)]

for pair in json_list:
    name, count = pair[0], pair[1]
    
    parts = name.split('_')
    capitalized_parts = [part.capitalize() for part in parts]
    
    capitalized_name = '_'.join(capitalized_parts)
    capitalized_name_with_space = ' '.join(capitalized_parts)
    
    spreadsheet = openpyxl.load_workbook(f"C:\\Users\\phams\\Downloads\\{capitalized_name}\\{name}_projects_links.xlsx")
    sheet = spreadsheet.active

    merge_doc = Document()

    for i in range(2,sheet.max_row+1):
        position = sheet.cell(row=i,column=1).value
        file_name = sheet.cell(row=i,column=3).value
        
        merge_doc.add_heading(f'{position} - {file_name}',level=1)
        
        source_doc = Document(f'C:\\Users\\phams\\Downloads\\{capitalized_name}\\thong tin du an\\{file_name}.docx')
        
        for paragraph in source_doc.paragraphs:
            merge_doc.add_paragraph(paragraph.text)
        
        merge_doc.add_paragraph('')

    merge_doc.save(f"C:\\Users\\phams\\Downloads\\{capitalized_name}\\{capitalized_name_with_space}.docx")
    
    print(f"Merged info for {capitalized_name_with_space}")