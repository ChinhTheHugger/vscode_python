import openpyxl
from collections import Counter
import time
import numpy
import igraph
import sys

path = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.04.19 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V22 - processed.XLSX"

wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active

start = time.time()
print("Start time = " + time.strftime("%H:%M:%S", time.gmtime(start)))
time.sleep(5)

# ***

sys.setrecursionlimit(4850)

# * Remove duplicate to get list of unique 

def delDup(startNum,endNum):
    for x in range(startNum,endNum):
        for i in range(startNum+1   ,endNum):
            if sheet_obj.cell(row=i,column=3).value != sheet_obj.cell(row=x,column=3).value:
                continue
            else:
                baseStr = str(sheet_obj.cell(row=x,column=3).value) + "&" + str(sheet_obj.cell(row=x,column=4).value) + "&" + str(sheet_obj.cell(row=x,column=15).value) + "&" + str(sheet_obj.cell(row=x,column=16).value)
                cmpStr = str(sheet_obj.cell(row=i,column=3).value) + "&" + str(sheet_obj.cell(row=i,column=4).value) + "&" + str(sheet_obj.cell(row=i,column=15).value) + "&" + str(sheet_obj.cell(row=i,column=16).value)
                if cmpStr == baseStr:
                    sheet_obj.cell(row=i,column=2).value = None
                else:
                    continue
    return


arr = [2]
for x in range(2,sheet_obj.max_row+1):
    for i in range(x+1,sheet_obj.max_row+2):
        if sheet_obj.cell(row=i,column=3).value != sheet_obj.cell(row=x,column=3).value:
            break
    if i not in arr:
        arr.append(i)
    x=i

# print(arr)
print(len(arr))

newArr = numpy.array(arr)
for x in range(0,len(newArr)-1):
    if newArr[x+1]-newArr[x]==1:
        continue
    else:
        for i in range(newArr[x],newArr[x+1]):
            delDup(newArr[x],newArr[x+1])

# print(wb_obj.sheetnames)

# ***

# from openpyxl.styles import PatternFill

# sheet_ori = wb_obj["VRSH Sort 22.12.26 tên+căn"]
# sheet_des = wb_obj["VRSH Sort 22.12.26 tên+căn 1S"]

# fill_cell = PatternFill(patternType='solid',fgColor='FCBA03')

# for x in range(2,sheet_des.max_row+1):
#     for i in range (2,sheet_ori.max_row+1):
#         if sheet_ori.cell(row=i,column=1).value == sheet_des.cell(row=x,column=1).value:
#             sheet_des.cell(row=x,column=1).fill = fill_cell

wb_obj.save(path)

end = time.time()

print("Run time = " + time.strftime("%H:%M:%S", time.gmtime(end-start)))
print("End time = " + time.strftime("%H:%M:%S", time.gmtime(end)))