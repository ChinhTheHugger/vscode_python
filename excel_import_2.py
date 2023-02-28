import openpyxl
from collections import Counter

path = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.02.28 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V17 V4.XLSX"

wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active

# ***

# * Count general phone and house

# for x in range(2,sheet_obj.max_row+1):
#     #phone
#     phoneCount = sheet_obj.cell(row = x, column = 16).value
#     sheet_obj.cell(row = x, column = 14).value = phoneCount.count(';') + 1
#     #house
#     houseCount = sheet_obj.cell(row = x, column = 3).value
#     sheet_obj.cell(row = x, column = 13).value = houseCount.count(';') + 1

# ***

# * Count house in each area

# for x in range(2,sheet_obj.max_row+1):
#     #VRSH
#     areaCount = sheet_obj.cell(row = x, column = 3).value
#     sheet_obj.cell(row = x, column = 5).value = areaCount.count('VRSH')
#     #VOP1
#     areaCount = sheet_obj.cell(row = x, column = 3).value
#     sheet_obj.cell(row = x, column = 6).value = areaCount.count('VOP1')
#     #VOP2
#     areaCount = sheet_obj.cell(row = x, column = 3).value
#     sheet_obj.cell(row = x, column = 7).value = areaCount.count('VOP2')
#     #ECP
#     areaCount = sheet_obj.cell(row = x, column = 3).value
#     sheet_obj.cell(row = x, column = 8).value = areaCount.count('ECP')
#     #STL
#     areaCount = sheet_obj.cell(row = x, column = 3).value
#     sheet_obj.cell(row = x, column = 9).value = areaCount.count('STL')
#     #VGP
#     areaCount = sheet_obj.cell(row = x, column = 3).value
#     sheet_obj.cell(row = x, column = 10).value = areaCount.count('VGP')
#     #VGV
#     areaCount = sheet_obj.cell(row = x, column = 3).value
#     sheet_obj.cell(row = x, column = 11).value = areaCount.count('VGV')

# ***

# * Count total area

# for x in range(2,sheet_obj.max_row+1):
#     count = 0
#     for i in range(5,11):
#         if sheet_obj.cell(row = x, column = i).value > 0:
#             count += 1
#     sheet_obj.cell(row = x, column = 12).value = count

# ***

# * Count total house

for x in range(2,sheet_obj.max_row+1):
    areaCount = sheet_obj.cell(row = x, column = 10).value
    sheet_obj.cell(row = x, column = 13).value = areaCount.count(';')+1
wb_obj.save(path)