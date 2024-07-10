import requests
import json
import polyline
import ast
import geopandas as gpd
from shapely.geometry import LineString
import psycopg2
import matplotlib.pyplot as plt
import contextily as ctx

# test GHTK route

import requests

cookies = {
    'Phpstorm-1bdbdc0b': 'e072668f-73ea-4ab2-b595-4517c460b36e',
    '_osm_location': '105.86035|20.99552|19|M',
    'iconSize': '32x32',
    'jenkins-timestamper-offset': '-25200000',
}

headers = {
    'Connection': 'keep-alive',
    'sec-ch-ua': '" Not A;Brand";v="99", "Chromium";v="96", "Google Chrome";v="96"',
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'GH-Client': 'web-ui 3.0',
    'sec-ch-ua-mobile': '?0',
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/96.0.4664.93 Safari/537.36',
    'sec-ch-ua-platform': '"Linux"',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Dest': 'empty',
    'Referer': 'http://localhost:8989/maps/?point=15.23119%2C108.262024&point=14.756291%2C108.492737&locale=en-US&elevation=false&profile=car&use_miles=false&layer=OpenStreetMap',
    'Accept-Language': 'en-US,en;q=0.9',
    # 'Cookie': 'Phpstorm-1bdbdc0b=e072668f-73ea-4ab2-b595-4517c460b36e; _osm_location=105.86035|20.99552|19|M; iconSize=32x32; jenkins-timestamper-offset=-25200000',
    'apikey': 'CgmUZhzdammE3A2guUgUXSyj',
    'Authorization': 'Bearer eyJhbGciOiJFUzI1NiIsImtpZCI6IjAxRjVOMThESE02RlkwSEpKSFhFRlE5NzNLXzE2MjA5ODIzODAiLCJ0eXAiOiJKV1QifQ.eyJhdWQiOiJhdXRoIiwiZXhwIjoxNjU2NjkxMjQ1LCJqdGkiOiIwMUc2WDRFTjlFOUM4RDNZSFRKQVlFQjM3NiIsImlhdCI6MTY1NjY4NzY0NSwiaXNzIjoiaHR0cHM6Ly9hdXRoLmdodGtsYWIuY29tIiwic3ViIjoiMDFHMlhNWTQzN1MxNEIyWVY0WDQzNlRTUkIiLCJzY3AiOlsib3BlbmlkIl0sInNpZCI6IlE1VFJFZFpXaHJ6SFJEa3JyRVE1bkM1dzRaRU93aDV3IiwiY2xpZW50X2lkIjoiMDFGNU4xOERITTZGWTBISkpIWEVGUTk3M0siLCJ0eXBlIjoib2F1dGgifQ.rq7Jj0zEBSq3nlEs5AZ1BAz5UL6BYbrjDz7QMyAnrXdVHPAqU9pfZ-VMXTbTrWfyxCB2h1pohcoBKbY4tOYvnQ',
    'Content-Type': 'application/json',
}

json_data = {
    'gh_requests': [
        {
            'points': [
                [
                    105.78739, 20.98039
                ],
                [
                    105.83399, 21.00814
                ],
            ],
            'vehicle': 'foot', # car, bike, motorcycle, xteam_motorcycle
            'request_id': '982884285',
            'calc_points': True,
            'points_encoded': True,
            'instructions': True,
            'locale': 'vi',
            'algorithm': 'alternative_route',
        },
    ],
}

response = requests.post('https://gmap-api-gw.ghtklab.com/route', cookies=cookies, headers=headers, json=json_data)
# print(json.dumps(response.json(),indent=4))

# Note: json_data will not be serialized by requests
# exactly as it was in the original request.
#data = '{\n    "gh_requests": [\n        {\n            "points": [\n                [\n                    105.863657,\n                    20.98368\n                ],\n                [\n                    105.861919,\n                    20.997323\n                ]\n            ],\n            "vehicle": "motorcycle",\n            "request_id": "982884285",\n            "calc_points": true,\n            "points_encoded": true,\n            "instructions": true,\n            "locale": "vi",\n            "algorithm": "alternative_route"\n        }\n    ]\n}'
#response = requests.post('https://gmap-api-gw.ghtklab.com/route', cookies=cookies, headers=headers, data=data)

data = response.json()

print(json.dumps(data,indent=4))

# test = data['gh_responses']
# test = test[0]
# test = test['paths']
# test = test[0]
# test  =test['points']
# test = polyline.decode(test)

# line = LineString([(lon, lat) for lat, lon in test])
# print(line)

# data = {
#     'geometry': [line],
#     'name': ['Route 1']
# }
# gdf = gpd.GeoDataFrame(data=data)
# print(gdf)

# colors = ['blue']

# fig, ax = plt.subplots()
# for idx, row in gdf.iterrows():
#     # Plot the LineString
#     if not row['geometry'].is_empty:
#         gdf.loc[[idx]].plot(ax=ax, label=row['name'], color=colors[idx])

# # Set a fixed aspect ratio
# ax.set_aspect('equal')

# # Customize and show the plot
# plt.title('Route from Point A to Point B')
# plt.xlabel('Longitude')
# plt.ylabel('Latitude')
# plt.legend(loc='lower right')
# ctx.add_basemap(ax, crs="EPSG:4326", source=ctx.providers.OpenStreetMap.Mapnik, zoom=11)
# plt.grid(True)
# plt.show()
