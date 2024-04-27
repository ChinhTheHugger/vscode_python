import openpyxl
import pandas
import requests
import json

path_true_coordinates = "C:\\Users\\phams\\Downloads\\Đà_Nẵng_Toạ_độ_đúng.XLSX"
path_ip1_ip2_coordinates = "C:\\Users\\phams\\Downloads\\Dataraw_đánh_giá_chất_lượng_phân_tích_địa_chỉ_tại_Đà_Nẵng.XLSX"
path_result = "C:\\Users\\phams\\Downloads\\api_tags.xlsx"

df_true = pandas.read_excel(path_true_coordinates)
df_ip = pandas.read_excel(path_ip1_ip2_coordinates)
df_result = pandas.read_excel(path_result)

df_true.index
df_ip.index



# initialize sheets

header_true = [
    'id','raw_text','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15'
]
header_ip = [
    'id','dcc5','1','2','3','4','5','6','7','8','9','10','11','12','13','14','15'
]

df_true = pandas.DataFrame(columns=header_true)
df_ip = pandas.DataFrame(columns=header_ip)

with pandas.ExcelWriter(path_result, mode='a') as writer:
    df_true.to_excel(writer,sheet_name='True',index=False)
    df_ip.to_excel(writer,sheet_name='ip',index=False)