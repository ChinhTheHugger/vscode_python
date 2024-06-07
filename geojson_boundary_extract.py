import geojson
from shapely.geometry import shape, Polygon, MultiPolygon
import chardet

# Detect file encoding
with open('C:\\Users\\phams\\Downloads\\export(1).geojson', 'rb') as f:
    raw_data = f.read()
    result = chardet.detect(raw_data)
    encoding = result['encoding']

# Load the GeoJSON file with the detected encoding
with open('C:\\Users\\phams\\Downloads\\export(1).geojson', encoding=encoding) as f:
    data = geojson.load(f)

# Extract polygons and convert to WKT
polygons = []
for feature in data['features']:
    geom = shape(feature['geometry'])
    if isinstance(geom, Polygon):
        polygons.append(geom)
    elif isinstance(geom, MultiPolygon):
        polygons.extend(geom.geoms)

# Combine all polygons into a single MultiPolygon
combined_geom = MultiPolygon(polygons)

# Print WKT
print(combined_geom.wkt)

# Save WKT to a text file
with open('C:\\Users\\phams\\Downloads\\boundary_wkt.txt', 'w') as f:
    f.write(combined_geom.wkt)

print("WKT has been written to boundary_wkt.txt")
