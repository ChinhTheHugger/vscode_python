# # NOTES:
# # - google geocoding api is not required
# # - google geocoding api require payment setup

# import requests


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


# api_key = "AIzaSyApQ25Iv-CFEFydCGrCoujAzlp972kh6AU"

# address = 'Vodickova 704/36, 11000 Prague 1, Czech Republic'


# lati, longi = get_lati_longi(api_key, address)

# print(f"Latitude: {lati}")

# print(f"Longitude: {longi}")





# import geopy

# API_KEY = "ArOIXETqt3OObS0AFEzlRC11hfLt47jM9lhm3DMOXmm2OcRppuuaSkSw7-gKI7vo"  # Replace with your Bing Maps API key

# geolocator = geopy.geocoders.Bing(API_KEY)

# def address_to_coordinates(address):
#     location = geolocator.geocode(address)
#     if hasattr(location, 'latitude') and hasattr(location, 'longitude'):
#         return (location.latitude, location.longitude)
#     else:
#         return None

# addresses = ["216 Phan Đăng Lưu Tiệm xăm Rontattoo"]  # List of addresses to convert
# coordinates = []  # List of coordinates results

# for address in addresses:
#     coordinate = address_to_coordinates(address)
#     coordinates.append(coordinate)

# print(coordinates)





# import googlemaps
# from datetime import datetime

# gmaps = googlemaps.Client(key='AIzaSyApQ25Iv-CFEFydCGrCoujAzlp972kh6AU')

# # Geocoding an address
# geocode_result = gmaps.geocode('1600 Amphitheatre Parkway, Mountain View, CA')

# # Look up an address with reverse geocoding
# reverse_geocode_result = gmaps.reverse_geocode((40.714224, -73.961452))

# # Request directions via public transit
# now = datetime.now()
# directions_result = gmaps.directions("Sydney Town Hall",
#                                      "Parramatta, NSW",
#                                      mode="transit",
#                                      departure_time=now)

# # Validate an address with address validation
# addressvalidation_result =  gmaps.addressvalidation(['1600 Amphitheatre Pk'], 
#                                                     regionCode='US',
#                                                     locality='Mountain View', 
#                                                     enableUspsCass=True)





from bs4 import BeautifulSoup

import requests

 

#Specify URL for Paris on Google Maps

url = "https://www.google.com/maps/place/Paris"

 

#Sending request and getting response object

response = requests.get(url)

 

#Parsing the HTML content with BeautifulSoup

soup = BeautifulSoup(response.text, "html.parser") 

 

#Extract location Information (Coordinates in this case)

coordinates_div = soup.find('meta', property="og:image:location:content")

 

if coordinates_div:

    print("Paris Coordinates:",coordinates_div["content"])

else:

    print("“Coordinates not found”")