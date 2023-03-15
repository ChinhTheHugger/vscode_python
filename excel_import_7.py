import openpyxl
from collections import Counter
import time
import numpy

path = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.03.15 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V22 - for processing V1.XLSX"

wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active

start = time.time()

# Sort

# for x in range(2,sheet_obj.max_row+1):
#     arr=str(sheet_obj.cell(row=x,column=5).value).split(';')
#     arr.sort()
#     newStr=';'.join(map(str,arr))
#     sheet_obj.cell(row=x,column=6).value=newStr

# Count length

# for x in range(2,sheet_obj.max_row+1):
#     sheet_obj.cell(row=x,column=7).value=str(sheet_obj.cell(row=x,column=6).value).count(';')+1

# Group name with house list

# for x in range(2,sheet_obj.max_row+1):
#     if sheet_obj.cell(row=x,column=7).value>200:
#         sheet_obj.cell(row=x,column=8).value=str(sheet_obj.cell(row=x,column=2).value)+";"+str(sheet_obj.cell(row=x,column=7).value)+" căn"
#     else:
#         sheet_obj.cell(row=x,column=8).value=str(sheet_obj.cell(row=x,column=2).value)+";"+str(sheet_obj.cell(row=x,column=6).value)

# Group people sharing phone

def pplGroup(x,i):
    if i-x==0:
        sheet_obj.cell(row=x,column=9).value=sheet_obj.cell(row=x,column=8).value
    else:
        for a in range(x,i+1):
            pplArr=[]
            pplArr.append(sheet_obj.cell(row=a,column=8).value)
            for b in range(x,i+1):
                if sheet_obj.cell(row=b,column=8).value not in pplArr:
                    pplArr.append(str(sheet_obj.cell(row=b,column=8).value))
            sheet_obj.cell(row=a,column=9).value='&'.join(pplArr)
    return

arr = [2]
for x in range(2,sheet_obj.max_row+1):
    for i in range(x+1,sheet_obj.max_row+2):
        if sheet_obj.cell(row=i,column=3).value != sheet_obj.cell(row=x,column=3).value:
            break
    if i not in arr:
        arr.append(i)
    x=i

print(arr)
print(len(arr))

newArr = numpy.array(arr)
for x in range(0,len(arr)-1):
    pplGroup(newArr[x],newArr[x+1]-1)



wb_obj.save(path)

end = time.time()

print(time.strftime("%H:%M:%S", time.gmtime(end-start)))