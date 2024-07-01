import requests
import json
import polyline
from shapely import LineString

params = {
}

response = requests.get(
    'http://router.project-osrm.org/route/v1/driving/13.388860,52.517037;13.397634,52.529407',
    params=params,
)

data = response.json()
encoded_poly = ""
print(json.dumps(data,indent=4))
for hint in data["waypoints"]:
    encoded_poly += hint["hint"]
print(encoded_poly)
