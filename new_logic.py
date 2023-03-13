import openpyxl
from collections import Counter
import time
import numpy

pathTest = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\Book1.XLSX"
path = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.03.13 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V22 - for processing.XLSX"

wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active

start = time.time()

# * First step *

# House group

# def houseGroup(x,i):
#     if i-x==0:
#         sheet_obj.cell(row=x,column=5).value=sheet_obj.cell(row=x,column=4).value
#     else:
#         for a in range(x,i+1):
#             houseArrTemp=[]
#             houseArrTemp.append(sheet_obj.cell(row=a,column=4).value)
#             for b in range(x,i+1):
#                 if sheet_obj.cell(row=b,column=3).value==sheet_obj.cell(row=a,column=3).value and sheet_obj.cell(row=b,column=4).value not in houseArrTemp:
#                     houseArrTemp.append(sheet_obj.cell(row=b,column=4).value)
#             # houseArr=list(dict.fromkeys(houseArrTemp))
#             houseStr=';'.join(map(str,houseArrTemp))
#             sheet_obj.cell(row=a,column=5).value=houseStr
#     return

# arr = [2]
# for x in range(2,sheet_obj.max_row+1):
#     for i in range(x+1,sheet_obj.max_row+2):
#         if sheet_obj.cell(row=i,column=2).value != sheet_obj.cell(row=x,column=2).value:
#             break
#     if i not in arr:
#         arr.append(i)
#     x=i

# print(arr)
# print(len(arr))

# newArr = numpy.array(arr)
# for x in range(0,len(arr)-2):
#     houseGroup(newArr[x],newArr[x+1]-1)

# Phone group - this one will be used to group people sharing phone later

# def phoneGroup(x,i):
#     if i-x==0:
#         sheet_obj.cell(row=x,column=6).value=sheet_obj.cell(row=x,column=3).value
#     else:
#         for a in range(x,i+1):
#             phoneArrTemp=[]
#             phoneArrTemp.append(sheet_obj.cell(row=a,column=3).value)
#             for b in range(x,i+1):
#                 if sheet_obj.cell(row=b,column=4).value==sheet_obj.cell(row=a,column=4).value and sheet_obj.cell(row=b,column=3).value not in phoneArrTemp:
#                     phoneArrTemp.append(sheet_obj.cell(row=b,column=3).value)
#             # houseArr=list(dict.fromkeys(houseArrTemp))
#             phoneStr=';'.join(map(str,phoneArrTemp))
#             sheet_obj.cell(row=a,column=6).value=phoneStr
#     return

# arr = [2]
# for x in range(2,sheet_obj.max_row+1):
#     for i in range(x+1,sheet_obj.max_row+2):
#         if sheet_obj.cell(row=i,column=2).value != sheet_obj.cell(row=x,column=2).value:
#             break
#     if i not in arr:
#         arr.append(i)
#     x=i

# print(arr)
# print(len(arr))

# newArr = numpy.array(arr)
# for x in range(0,len(arr)-2):
#     phoneGroup(newArr[x],newArr[x+1]-1)

# wb_obj.save(path)

# * Second step *

# House group

def houseStepTwo(x,i,ele,houseArrTemp):
    houseArr=houseArrTemp
    for a in range(x,i+1):
        if ele in str(sheet_obj.cell(row=a,column=3).value):
            arrTemp=str(sheet_obj.cell(row=a,column=4).value).split(';')
            for tEle in arrTemp:
                if tEle not in houseArr:
                    houseArr.append(tEle)
    return houseArr

def houseGroup(x,i):
    if i-x==0:
        sheet_obj.cell(row=x,column=5).value=sheet_obj.cell(row=x,column=4).value
    else:
        for a in range(x,i+1):
            pseudoHouseArr=str(sheet_obj.cell(row=a,column=3).value).split(';')
            houseTempStr=""
            houseTempStr=str(sheet_obj.cell(row=a,column=4).value)
            houseArrTemp=houseTempStr.split(';')
            for ele in pseudoHouseArr:
                newHouseArr=houseStepTwo(x,i,ele,houseArrTemp)
            # houseArr=list(dict.fromkeys(houseArrTemp))
            houseStr=';'.join(map(str,newHouseArr))
            sheet_obj.cell(row=a,column=5).value=houseStr
    return

arr = [2]
for x in range(2,sheet_obj.max_row+1):
    for i in range(x+1,sheet_obj.max_row+2):
        if sheet_obj.cell(row=i,column=2).value != sheet_obj.cell(row=x,column=2).value:
            break
    if i not in arr:
        arr.append(i)
    x=i

print(arr)
print(len(arr))

newArr = numpy.array(arr)
for x in range(0,len(arr)-1):
    houseGroup(newArr[x],newArr[x+1]-1)

# Phone group

# def phoneStepTwo(x,i,ele,phoneArrTemp):
#     phoneArr=phoneArrTemp
#     for a in range(x,i+1):
#         if ele in str(sheet_obj.cell(row=a,column=4).value):
#             arrTemp=str(sheet_obj.cell(row=a,column=3).value).split(';')
#             for tEle in arrTemp:
#                 if tEle not in phoneArr:
#                     phoneArr.append(tEle)
#     return phoneArr

# def phoneGroup(x,i):
#     if i-x==0:
#         sheet_obj.cell(row=x,column=6).value=sheet_obj.cell(row=x,column=3).value
#     else:
#         for a in range(x,i+1):
#             pseudoPhoneArr=str(sheet_obj.cell(row=a,column=4).value).split(';')
#             phoneTempStr=""
#             phoneTempStr=str(sheet_obj.cell(row=a,column=3).value)
#             phoneArrTemp=phoneTempStr.split(';')
#             for ele in pseudoPhoneArr:
#                 newPhoneArr=phoneStepTwo(x,i,ele,phoneArrTemp)
#             # houseArr=list(dict.fromkeys(houseArrTemp))
#             phoneStr=';'.join(map(str,phoneArrTemp))
#             sheet_obj.cell(row=a,column=6).value=phoneStr
#     return

# arr = [2]
# for x in range(2,sheet_obj.max_row+1):
#     for i in range(x+1,sheet_obj.max_row+2):
#         if sheet_obj.cell(row=i,column=2).value != sheet_obj.cell(row=x,column=2).value:
#             break
#     if i not in arr:
#         arr.append(i)
#     x=i

# print(arr)
# print(len(arr))

# newArr = numpy.array(arr)
# for x in range(0,len(arr)-2):
#     phoneGroup(newArr[x],newArr[x+1]-1)



wb_obj.save(path)

end = time.time()

print(time.strftime("%H:%M:%S", time.gmtime(end-start)))