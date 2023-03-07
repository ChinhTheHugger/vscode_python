import openpyxl
import pandas
import time

path = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.03.04 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V21 - for processing.XLSX"

pathOne = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.03.06 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V21 - for processing.CSV"
pathTwo = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.03.06 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V17 V4 - Copy - Copy.CSV"

# workbook = openpyxl.load_workbook(pathOne)
# sheet = workbook.active

maindf = pandas.read_csv(pathOne,encoding='latin-1')
df = pandas.read_csv(pathTwo,encoding='latin-1')

# df = df.drop(df.index[1])
# print(df)
# print(len(df.index))

df.drop(['Phone temp','ghép c?n theo t?ng s? ?i?n tho?i d?a theo tên','ghép s? ?i?n tho?i theo t?ng c?n d?a theo tên','ghép danh sách c?n ? c?t H có liên quan ??n t?ng s? ?i?n tho?i ? c?t I theo t?ng c?n d?a theo tên','ghép c?t C và c?t J','c?t chép version 1'],axis=1,inplace=True)
maindf.drop(['Count after comparing','Difference','Name after comparing'],axis=1,inplace=True)
# print(df)
# print(maindf)
# df_compare = maindf.compare(df,keep_equal=False)
# print(df_compare)
# df_compare.to_excel("E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\pandas_dataframe_comparing.xlsx")

# print(len(maindf.index))

maindf_copy = maindf

start = time.time(

)
for x in range(0,len(maindf.index)):
    phone = maindf.at[x,'Phone']
    house = maindf.at[x,'House']
    bday = maindf.at[x,'Date of Birth']
    id = maindf.at[x,'ID number']
    for i in range(0,len(df.index)):
        if df.at[i,'Phone'] == phone and df.at[i,'House'] == house and df.at[i,'Date of Birth'] == bday and df.at[i,'ID number'] == id:
            maindf.at[x,'Count'] = df.at[i,'Count']
            maindf.at[x,'Searxh term 1'] = df.at[i,'Search term 1']
            df = df.dropna().reset_index(drop=True)

# maindf.at[0,'Count after comparing'] = "test"
print(maindf_copy.compare(maindf))

end = time.time()

print(time.strftime("%H:%M:%S", time.gmtime(end-start)))

# workbook.save(pathOne)