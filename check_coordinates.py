import openpyxl
import pandas
import requests
import json
import time
from datetime import datetime
import haversine
from haversine import Unit
from geopy.geocoders import Nominatim, ArcGIS

path_true_coordinates = "C:\\Users\\phams\\Downloads\\Đà_Nẵng_Toạ_độ_đúng.XLSX"
path_ip1_ip2_coordinates = "C:\\Users\\phams\\Downloads\\Dataraw_đánh_giá_chất_lượng_phân_tích_địa_chỉ_tại_Đà_Nẵng.XLSX"
path_result = "C:\\Users\\phams\\Downloads\\check_coordinates.xlsx"

df_true = pandas.read_excel(path_true_coordinates)
df_ip = pandas.read_excel(path_ip1_ip2_coordinates)
df_result = pandas.read_excel(path_result)

loc = Nominatim(user_agent="check_coordinates")
geolocator_arcgis = ArcGIS()

print("Running")

# def get_lati_longi(api_key, address):

#     url = 'https://maps.googleapis.com/maps/api/geocode/json'

#     params = {
#         "address": address,
#         "key": api_key
#     }

#     response = requests.get(url, params=params)

#     if response.status_code == 200:

#         data = response.json()

#         if data["status"] == "OK":

#             location = data["results"][0]["geometry"]["location"]

#             lat = location["lat"]

#             lng = location["lng"]

#             return lat, lng

#         else:

#             print(f"Error: {data['error_message']}")

#             return 0, 0

#     else:

#         print("Failed to make the request.")

#         return 0, 0



start_calculation = time.time()
start = datetime.now()



header = [
    'pkg_order','raw_text','true_lat','true_lng','ggl_lat','ggl_lng','lat_gap','lng_gap','distance'
]
results = []

api_key = "AIzaSyApQ25Iv-CFEFydCGrCoujAzlp972kh6AU"



for i in range(0,7700):
    result_frame = [
        '','','','','','','','',''
    ]
    
    result_frame[0] = str(df_true.loc[i]['pkg_order'])
    result_frame[1] = str(df_true.loc[i]['raw_text'])
    result_frame[2] = str(df_true.loc[i]['true_lat'])
    result_frame[3] = str(df_true.loc[i]['true_lng'])

    address = str(df_true.loc[i]['raw_text'])

    # lati, longi = get_lati_longi(api_key, address)
    
    # getLoc = loc.geocode(address)
    
    getLoc = geolocator_arcgis.geocode(address)

    result_frame[4] = getLoc.latitude
    result_frame[5] = getLoc.longtitude

    result_frame[6] = abs(float(df_true.loc[i]['true_lat']) - float(getLoc.latitude))
    result_frame[7] = abs(float(df_true.loc[i]['true_lng']) - float(getLoc.longtitude))

    old_coordinates = (float(df_true.loc[i]['true_lat']),float(getLoc.latitude))
    ggl_coordinates = (float(df_true.loc[i]['true_lng']),float(getLoc.longtitude))

    result_frame[8] = haversine.haversine(old_coordinates,ggl_coordinates,unit=Unit.METERS)

    results.append(result_frame)

df_result = pandas.DataFrame(data=results,columns=header)

with pandas.ExcelWriter(path_result, mode='a', if_sheet_exists="replace") as writer:
    df_result.to_excel(writer, sheet_name='coordinates_check_google', index=False)



end_calculation = time.time()
end = datetime.now()

print("Total run time = ", time.strftime("%H:%M:%S", time.gmtime(end_calculation-start_calculation)))