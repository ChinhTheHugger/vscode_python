# 

import openpyxl
from collections import Counter
import time
import numpy

pathTest = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\Book1.XLSX"
path = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.03.10 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V21 - for processing.XLSX"

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

def houseGroup(x,i):
    if i-x==0:
        sheet_obj.cell(row=x,column=7).value=sheet_obj.cell(row=x,column=6).value
    else:
        for a in range(x,i+1):
            houseArrTemp=[]
            houseArrTemp.append(sheet_obj.cell(row=a,column=6).value)
            for b in range(x,i+1):
                if sheet_obj.cell(row=b,column=3).value==sheet_obj.cell(row=a,column=3).value:
                    houseArrTemp.append(sheet_obj.cell(row=b,column=6).value)
            houseArr=list(dict.fromkeys(houseArrTemp))
            houseStr=';'.join(map(str,houseArr))
            sheet_obj.cell(row=a,column=7).value=houseStr
    return

arr = [2]
for x in range(2,130):
    for i in range(x+1,131):
        if sheet_obj.cell(row=i,column=2).value != sheet_obj.cell(row=x,column=2).value:
            break
    if i not in arr:
        arr.append(i)
    x=i

print(arr)

newArr = numpy.array(arr)
for x in range(0,len(arr)-2):
    houseGroup(newArr[x],newArr[x+1]-1)

wb_obj.save(path)

end = time.time()

print(time.strftime("%H:%M:%S", time.gmtime(end-start)))