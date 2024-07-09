import asyncio
from playwright.async_api import async_playwright
import openpyxl
from bs4 import BeautifulSoup
import re
from docx import Document

spreadsheet = openpyxl.load_workbook("C:\\Users\\phams\\Downloads\\Bac_Ninh\\bac_ninh_projects_links.xlsx")
sheet = spreadsheet.active

merge_doc = Document()

for i in range(2,sheet.max_row+1):
    position = sheet.cell(row=i,column=1).value
    name = sheet.cell(row=i,column=3).value
    
    merge_doc.add_heading(f'{position} - {name}',level=1)
    
    source_doc = Document(f'C:\\Users\\phams\\Downloads\\Bac_Ninh\\thong tin cac du an\\{name}.docx')
    
    for paragraph in source_doc.paragraphs:
        merge_doc.add_paragraph(paragraph.text)
    
    merge_doc.add_paragraph('')

merge_doc.save("C:\\Users\\phams\\Downloads\\Bac_Ninh\\Bac Ninh.docx")