# 

import openpyxl
from collections import Counter
import time

path = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\Book1.XLSX"

wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active

start = time.time()

# Phone group  > house step 1

# for x in range(2,sheet_obj.max_row+1):
#     sheet_obj.cell(row=x,column=4).value = sheet_obj.cell(row=x,column=2).value
#     for i in range(x+1,sheet_obj.max_row+1):
#         if sheet_obj.cell(row=i,column=1).value != sheet_obj.cell(row=x,column=1).value:
#             break
#         else:
#             if sheet_obj.cell(row=i,column=3).value != sheet_obj.cell(row=x,column=3).value:
#                 break
#             else:
#                 temp = sheet_obj.cell(row=x,column=4).value + ";" + sheet_obj.cell(row=i,column=2).value
#                 sheet_obj.cell(row=x,column=4).value = temp
#     x=i

# for x in range(2,sheet_obj.max_row+1):
#     peopleString = str(sheet_obj.cell(row = x, column = 4).value)
#     peopleArrayToSort = peopleString.split(';')
#     peopleArrayToSort = list(dict.fromkeys(peopleArrayToSort))
#     peopleListToStr = ';'.join(map(str,peopleArrayToSort))
#     sheet_obj.cell(row = x, column = 4).value = peopleListToStr

# House group > phone step 2

# for x in range(2,sheet_obj.max_row+1):
#     sheet_obj.cell(row=x,column=5).value = sheet_obj.cell(row=x,column=3).value
#     for i in range(x+1,sheet_obj.max_row+1):
#         if sheet_obj.cell(row=i,column=1).value != sheet_obj.cell(row=x,column=1).value:
#             break
#         else:
#             if sheet_obj.cell(row=i,column=2).value != sheet_obj.cell(row=x,column=2).value:
#                 break
#             else:
#                 temp = sheet_obj.cell(row=x,column=5).value + ";" + sheet_obj.cell(row=i,column=3).value
#                 sheet_obj.cell(row=x,column=5).value = temp
#     x=i

# for x in range(2,sheet_obj.max_row+1):
#     peopleString = str(sheet_obj.cell(row = x, column = 5).value)
#     peopleArrayToSort = peopleString.split(';')
#     peopleArrayToSort = list(dict.fromkeys(peopleArrayToSort))
#     peopleListToStr = ';'.join(map(str,peopleArrayToSort))
#     sheet_obj.cell(row = x, column = 5).value = peopleListToStr

def houseGroup(x,i,phoneStr):
    houseStr = []
    phoneArr = phoneStr.split(';')
    for phone in phoneArr:
        for a in range(x,i-1):
            if sheet_obj.cell(row=x,column=3).value == phone:
                houseStr.append(sheet_obj.cell(row=a,column=2).value)
    houseStr = list(dict.fromkeys(houseStr))
    houseLst = ';'.join(map(str,houseStr))
    return houseLst

for x in range(2,sheet_obj.max_row+1):
    count=1
    for i in range(x+1,sheet_obj.max_row+2):
        if sheet_obj.cell(row=i,column=1).value == sheet_obj.cell(row=x,column=1).value:
            count += 1
        else:
            break
    print(count)
    x=i

# sheet_obj.cell(row=2,column=6).value = houseGroup(2,13,sheet_obj.cell(row=2,column=5).value)
# print(houseGroup(2,13,sheet_obj.cell(row=2,column=5).value))
# sheet_obj.cell(row=2,column=6).value = houseGroup(13,20,sheet_obj.cell(row=13,column=5).value)
# print(houseGroup(13,20,sheet_obj.cell(row=2,column=5).value))
# sheet_obj.cell(row=2,column=6).value = houseGroup(20,26,sheet_obj.cell(row=20,column=5).value)
# print(houseGroup(20,26,sheet_obj.cell(row=2,column=5).value))

# print(sheet_obj.max_row)

wb_obj.save(path)

end = time.time()

print(time.strftime("%H:%M:%S", time.gmtime(end-start)))