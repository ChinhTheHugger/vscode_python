import openpyxl
from collections import Counter

path = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.02.28 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V17 V4.XLSX"

wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active

# * Check houses of each person in a shared phone group and remove unique houses of people with large number of houses to reduce string length

for x in range(2,sheet_obj.max_row+1):
    if sheet_obj.cell(row = x, column = 2).value == "BIỆN XUÂN KHEN":
        sheet_obj.cell(row = x, column = 14).value = "x"
for i in range(2,sheet_obj.max_row+1):
    if sheet_obj.cell(row = i, column = 14).value == "x":
        for k in range(i+1,2000):
            if sheet_obj.cell(row = k, column = 4).value == sheet_obj.cell(row = i, column = 4).value:
                sheet_obj.cell(row = k, column = 3).value = None
wb_obj.save(path)