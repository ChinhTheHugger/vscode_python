import requests
import json
import polyline
from shapely import LineString


# Known valid encoded polyline string
known_valid_encoded_polyline = '_p~iF~ps|U_ulLnnqC_mqNvxq'

params = {
}

response = requests.get(
    'http://router.project-osrm.org/route/v1/driving/13.388860,52.517037;13.397634,52.529407',
    params=params,
)

data = response.json()
encoded_poly = ''
print(json.dumps(data,indent=4))
for hint in data["waypoints"]:
    encoded_poly += hint["hint"]
print(encoded_poly)

# Function to decode polyline
def decode_polyline(encoded):
    try:
        # Decode the polyline into a list of (latitude, longitude) tuples
        coordinates = polyline.decode(encoded)
        print("Decoded coordinates:", coordinates)
    except IndexError as e:
        print("IndexError decoding polyline:", e)
    except Exception as e:
        print("General error decoding polyline:", e)
    
# Call the function with the known valid encoded polyline
decode_polyline(known_valid_encoded_polyline)
