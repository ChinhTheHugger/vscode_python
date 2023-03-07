import openpyxl
import pandas
import time

# pathOne = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\test_1 (name match the original V21 on mar 6).CSV"
# pathTwo = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\test_2 (copied from test_2_original).CSV"

# df_one = pandas.read_csv(pathOne,encoding='UTF-8')
# df_two = pandas.read_csv(pathTwo,encoding='UTF-8')

# # df_one.drop(['Phone temp','ghép căn theo từng số điện thoại dựa theo tên','ghép số điện thoại theo từng căn dựa theo tên','ghép danh sách căn ở cột H có liên quan đến từng số điện thoại ở cột I theo từng căn dựa theo tên','ghép cột C và cột J','cột chép version 1'],axis=1,inplace=True)
# # df_two.drop(['Phone temp','ghép căn theo từng số điện thoại dựa theo tên','ghép số điện thoại theo từng căn dựa theo tên','ghép danh sách căn ở cột H có liên quan đến từng số điện thoại ở cột I theo từng căn dựa theo tên','ghép cột C và cột J','cột chép version 1'],axis=1,inplace=True)

# df_compare = df_one.compare(df_two,keep_equal=False)
# df_compare.to_excel("E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\pandas_dataframe_comparing.xlsx")

# print(df_one.reset_index(drop=True).equals(df_two.reset_index(drop=True)))

# ***

wb_comparison = openpyxl.load_workbook("E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\pandas_dataframe_comparing.xlsx")
sh_comparison = wb_comparison.active

wb_destination = openpyxl.load_workbook("E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\test_2_original - after fixed V1.xlsx")
sh_destination = wb_destination.active

# print(sh_comparison.cell(row=3,column=1).value)
# print(sh_comparison.cell(row=4,column=1).value)

# print(sh_destination.cell(row=4,column=1).value)
# print(sh_destination.cell(row=4,column=2).value)
