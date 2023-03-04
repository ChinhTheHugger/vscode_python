import openpyxl
import pandas

path = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.03.04 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V21 - for processing.XLSX"

pathOne = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.03.04 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V21 - for processing.CSV"
pathTwo = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.03.04 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V17 V4 - Copy - Copy.CSV"

# workbook = openpyxl.load_workbook(pathOne)
# sheet = workbook.active

maindf = pandas.read_csv(pathOne,encoding='latin-1')
df = pandas.read_csv(pathTwo,encoding='latin-1')

# df = df.drop(df.index[1])
# print(df)
# print(len(df.index))

# print(maindf)
# print(len(maindf.index))

# print(maindf.iloc[0,3])

# for x in range(2,sheet.max_row):
#     phone = sheet.cell(row=x,column=6).value
#     house = sheet.cell(row=x,column=7).value
#     bday = sheet.cell(row=x,column=8).value
#     id = sheet.cell(row=x,column=9).value
#     for i in range(0,len(df.index)-1):
#         if df.iloc[i,3] == phone and df.iloc[i,4] == house and df.iloc[i,5] == bday and df.iloc[i,6] == id:
#             sheet.cell(row=x,column=2).value = df.iloc[i,0]
#             df = df.dropna().reset_index(drop=True)

for x in range(0,len(maindf.index)-1):
    phone = maindf.iloc[x,5]
    house = maindf.iloc[x,6]
    bday = maindf.iloc[x,7]
    id = maindf.iloc[x,8]
    for i in range(0,len(df.index)-1):
        if df.iloc[i,3] == phone and df.iloc[i,4] == house and df.iloc[i,5] == bday and df.iloc[i,6] == id:
            maindf.iloc[x,1] = df.iloc[i,0]
            maindf.iloc[x,4] = df.iloc[i,2]
            df = df.dropna().reset_index(drop=True)

print(maindf)

# workbook.save(pathOne)