import openpyxl
from collections import Counter
import time

path = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.03.07 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V21 - processed.XLSX"

wb_obj = openpyxl.load_workbook(path)
sheet_obj = wb_obj.active

start = time.time()

for x in range(2,sheet_obj.max_row+1):
    for i in range(x+1, 5000):
        if sheet_obj.cell(row=i,column=3).value != sheet_obj.cell(row=x,column=3).value:
            break
        else:
            if sheet_obj.cell(row=i,column=6).value == sheet_obj.cell(row=x,column=6).value:
                sheet_obj.cell(row=i,column=1).value = None
    x=i

wb_obj.save(path)

end = time.time()

print(time.strftime("%H:%M:%S", time.gmtime(end-start)))