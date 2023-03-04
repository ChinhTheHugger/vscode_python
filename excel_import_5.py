import openpyxl
from collections import Counter

path = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.03.03 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V20 - Copy.XLSX"

wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active

# ***

# * Remove duplicate to get list of unique 

# for x in range(2,sheet_obj.max_row+1):
#     for i in range(x+1,x+5000):
#         if sheet_obj.cell(row=i,column=3).value != sheet_obj.cell(row=x,column=3).value:
#             break
#         else:
#             if sheet_obj.cell(row=i,column=5).value != sheet_obj.cell(row=x,column=5).value:
#                 break
#             else:
#                 sheet_obj.cell(row=i,column=2).value = None
#     x=i

# print(wb_obj.sheetnames)

# ***

from openpyxl.styles import PatternFill

sheet_ori = wb_obj["VRSH Sort 22.12.26 tên+căn"]
sheet_des = wb_obj["VRSH Sort 22.12.26 tên+căn 1S"]

fill_cell = PatternFill(patternType='solid',fgColor='FCBA03')

for x in range(2,sheet_des.max_row+1):
    for i in range (2,sheet_ori.max_row+1):
        if sheet_ori.cell(row=i,column=1).value == sheet_des.cell(row=x,column=1).value:
            sheet_des.cell(row=x,column=1).fill = fill_cell
wb_obj.save(path)