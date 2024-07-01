import requests
import json
import polyline
from shapely import LineString

headers = {
    'Content-Type': 'application/json; charset=utf-8',
    'Accept': 'application/json, application/geo+json, application/gpx+xml, img/png; charset=utf-8',
    'Authorization': '5b3ce3597851110001cf624882d0e6ed8c5c4aa28b9c89a1157d53f6',
}

json_data = {
    'coordinates': [
        [
            300,
            49.41461,
        ],
        [
            8.687872,
            49.420318,
        ],
    ],
}

response = requests.post('https://api.openrouteservice.org/v2/directions/driving-car/json', headers=headers, json=json_data)

# Note: json_data will not be serialized by requests
# exactly as it was in the original request.
#data = '{"coordinates":[[8.681495,49.41461],[8.687872,49.420318]]}'
#response = requests.post('https://api.openrouteservice.org/v2/directions/driving-car/json', headers=headers, data=data)

data = response.json()

print(json.dumps(data,indent=4))

encoded = data['routes'][0]['geometry']
decoded = polyline.decode(encoded)
decoded = [(lon,lat) for lat, lon in decoded]

print(decoded)
