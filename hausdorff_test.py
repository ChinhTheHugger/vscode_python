import requests
from shapely.geometry import LineString, Point
from shapely.ops import split
from shapely import wkt
import numpy as np
import json
import polyline
import ast
import matplotlib.pyplot as plt
import contextily as ctx
import geopandas as gpd
import psycopg2
import gspread
from google.oauth2.service_account import Credentials
from docx import Document
from docx.shared import Inches
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import io
import os
import pandas as pd
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from PIL import Image
from pyproj import Proj, transform
from geojson_length import calculate_distance, Unit
from geojson import Feature, LineString as GeoJSONLineString
import openpyxl
from scipy.spatial.distance import directed_hausdorff
from docx.shared import Pt
from scipy.spatial.distance import euclidean
from fastdtw import fastdtw
import networkx as nx
from scipy.spatial.distance import cdist



# Haversine distance function
def haversine(lonlat1, lonlat2):
    R = 6371.0  # Earth radius in kilometers
    
    lon1, lat1 = np.radians(lonlat1[:, 0]), np.radians(lonlat1[:, 1])
    lon2, lat2 = np.radians(lonlat2[:, 0]), np.radians(lonlat2[:, 1])
    
    dlon = lon2 - lon1[:, np.newaxis]
    dlat = lat2 - lat1[:, np.newaxis]
    
    a = np.sin(dlat / 2)**2 + np.cos(lat1[:, np.newaxis]) * np.cos(lat2) * np.sin(dlon / 2)**2
    c = 2 * np.arcsin(np.sqrt(a))
    
    return R * c

# Define a function to interpolate points along a LineString
def interpolate_points(line, num_points):
    distances = np.linspace(0, line.length, num_points)
    return np.array([line.interpolate(distance).coords[0] for distance in distances])

# OSM GraphHopper
linestring1 = LineString([(21.00798, 105.83418), (21.00835, 105.83383), (21.00826, 105.83372), (21.00824, 105.83362), (21.00794, 105.83327), (21.00787, 105.83316), (21.00753, 105.83274), (21.00716, 105.83227), (21.00702, 105.83208), (21.00643, 105.83131), (21.00633, 105.83116), (21.00603, 105.83106), (21.00577, 105.83095), (21.00536, 105.83079), (21.00412, 105.83033), (21.00311, 105.83001), (21.00206, 105.8296), (21.00149, 105.82934), (21.00106, 105.82917), (21.00096, 105.82895), (21.0009, 105.82879), (21.00133, 105.82756), (21.00146, 105.82712), (21.00162, 105.82651), (21.00169, 105.82627), (21.00166, 105.82599), (21.00266, 105.82205), (21.00276, 105.82193), (21.00304, 105.82093), (21.00313, 105.82054), (21.00318, 105.82023), (21.00318, 105.82016), (21.00317, 105.82009), (21.00309, 105.81997), (21.0028, 105.81956), (21.00248, 105.81911), (21.00189, 105.81829), (21.00179, 105.81823), (21.00156, 105.81785), (21.00112, 105.81721), (21.00099, 105.81712), (21.00067, 105.81668), (21.00061, 105.81644), (20.99894, 105.81394), (20.99819, 105.81278), (20.99755, 105.81183), (20.99661, 105.81038), (20.9957, 105.80901), (20.99498, 105.80796), (20.99413, 105.80676), (20.99397, 105.80656), (20.99386, 105.80649), (20.99361, 105.80616), (20.99183, 105.80362), (20.98935, 105.80009), (20.98927, 105.79994), (20.98713, 105.79688), (20.98617, 105.79554), (20.98536, 105.79436), (20.98452, 105.79314), (20.98432, 105.79288), (20.98366, 105.79205), (20.98324, 105.79149), (20.98272, 105.79084), (20.982, 105.78992), (20.98182, 105.78967), (20.98168, 105.7895), (20.98135, 105.78908), (20.98084, 105.78847), (20.98071, 105.78831), (20.98037, 105.78788), (20.98, 105.7874), (20.97939, 105.78666), (20.97933, 105.78659), (20.9792, 105.78643), (20.97913, 105.78633), (20.97906, 105.78624), (20.97785, 105.78479), (20.97634, 105.78282), (20.97578, 105.78207), (20.97504, 105.78105), (20.97485, 105.78082), (20.97472, 105.78068), (20.97464, 105.78057), (20.97407, 105.77977), (20.9735, 105.77891), (20.97334, 105.77869), (20.97152, 105.77638), (20.971, 105.77576), (20.9707, 105.77542), (20.96975, 105.77423), (20.96899, 105.77326), (20.96825, 105.77232), (20.9676, 105.77153), (20.96734, 105.77117)])
# GHTK GraphHopper
linestring2 = LineString([(21.00798, 105.83417), (21.00835, 105.83383), (21.00825, 105.83371), (21.00823, 105.83363), (21.00714, 105.8323), (21.00656, 105.83156), (21.00627, 105.83115), (21.00747, 105.82853), (21.00778, 105.82789), (21.00937, 105.8244), (21.00716, 105.82325), (21.00575, 105.82254), (21.0054, 105.82241), (21.00497, 105.82221), (21.00456, 105.82194), (21.00437, 105.82177), (21.00428, 105.82166), (21.00354, 105.82075), (21.00228, 105.81902), (21.00192, 105.81851), (21.00182, 105.81827), (21.00158, 105.81789), (20.9956, 105.80894), (20.99497, 105.80796), (20.99411, 105.80677), (20.99375, 105.80623), (20.99359, 105.80616), (20.9918, 105.80364), (20.98973, 105.80064), (20.98933, 105.8001), (20.98927, 105.79993), (20.98479, 105.79357), (20.98389, 105.79233), (20.98199, 105.78991), (20.98181, 105.78967), (20.98075, 105.78832), (20.98059, 105.78811), (20.97965, 105.78698), (20.97907, 105.78623), (20.97804, 105.78502), (20.97724, 105.78404), (20.97499, 105.78099), (20.97427, 105.78006), (20.97408, 105.77976), (20.9735, 105.7789), (20.9732, 105.77852), (20.97156, 105.77643), (20.971, 105.77575), (20.97069, 105.77542), (20.97023, 105.77485), (20.96825, 105.77232), (20.96735, 105.77116)])

linestring1 = LineString((lon,lat) for lat,lon in linestring1.coords)
linestring2 = LineString((lon,lat) for lat,lon in linestring2.coords)

a = np.array(linestring1.coords)
b = np.array(linestring2.coords)

# Interpolate additional points along each LineString
num_points = 250  # Number of points to interpolate
interpolated_a = interpolate_points(linestring1, num_points)
interpolated_b = interpolate_points(linestring2, num_points)

# Calculate pairwise haversine distances for interpolated points
distances_ab = haversine(interpolated_a, interpolated_b)

# Find the minimum distance for each point in interpolated_a to any point in interpolated_b
min_distances_a_to_b = distances_ab.min(axis=1)

# Calculate key statistics
max_distance = np.max(min_distances_a_to_b)
min_distance = np.min(min_distances_a_to_b)
average_distance = np.mean(min_distances_a_to_b)

# Thresholds in kilometers
minor_divergence_threshold = 0.0035 # meters
major_divergence_threshold = 0.015 # meters

# Identify significant peaks
minor_divergences = np.where((min_distances_a_to_b > minor_divergence_threshold) & 
                             (min_distances_a_to_b <= major_divergence_threshold))[0]
major_divergences = np.where(min_distances_a_to_b > major_divergence_threshold)[0]

# Generate textual summary
summary = f"""
Summary of LineString Comparison:

Key Statistics:
- Maximum distance: {max_distance:.2f} km
- Minimum distance: {min_distance:.2f} km
- Average distance: {average_distance:.2f} km
"""

# if len(minor_divergences) > 0:
#     summary += "\nMinor Divergences (distance > 3.5 meters and <= 15 meters, which is within the width of a normal roads with 1 to 4 lanes):\n"
#     for idx in minor_divergences:
#         summary += f"  - Point {idx} on LineString A has a minimum distance of {min_distances_a_to_b[idx]:.2f} km to LineString B\n"
# else:
#     summary += "\nNo minor divergences (distance > 3.5 meters and <= 15 meters, which is within the width of a normal roads with 1 to 4 lanes).\n"

# if len(major_divergences) > 0:
#     summary += "\nMajor Divergences (distance > 15 meters, which is bigger than the width of a 4 lanes road):\n"
#     for idx in major_divergences:
#         summary += f"  - Point {idx} on LineString A has a minimum distance of {min_distances_a_to_b[idx]:.2f} km to LineString B\n"
# else:
#     summary += "\nNo major divergences (distance > 15 meters, which is bigger than the width of a 4 lanes road).\n"
    
if len(major_divergences) > 0:
    summary += f"\nPercentage of sections of LineString A with major divergence from LineString B (distance > 15 meters, which is bigger than the width of a 4 lanes road): {len(major_divergences) * 100 / num_points} %\n"

print(summary)

# Create a figure with 2 subplots
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))

# Plot the distance graph on the first subplot
ax1.plot(range(len(interpolated_a)), min_distances_a_to_b, marker='o', linestyle='-', color='b', markersize=1, linewidth=1)
ax1.axhline(y=major_divergence_threshold, color='r', linestyle='--', label='Major Divergence Threshold', linewidth=1)
ax1.set_title('Shortest Distance from Each Point on LineString A to LineString B')
ax1.set_xlabel('Index of Point on LineString A')
ax1.set_ylabel('Minimum Distance to LineString B (km)')
ax1.grid(True)

# Plot the linestrings
ls_data = {
    'geometry': [linestring1,linestring2],
    'route_name': ['linestring 1','linestring 2']
}
    
gdf = gpd.GeoDataFrame(data=ls_data, crs="EPSG:4326")

colors = ['red','blue']

for idx, row in gdf.iterrows():
    # Plot the LineString
    if not row['geometry'].is_empty:
        gdf.loc[[idx]].plot(ax=ax2, label=row['route_name'], color=colors[idx % len(colors)], linewidth=2)

# Customize and show the plot
ax2.set_title('Route from Point A to Point B')
ax2.set_xlabel('Longitude')
ax2.set_ylabel('Latitude')
ax2.legend(loc='lower right')
ctx.add_basemap(ax2, crs="EPSG:4326", source=ctx.providers.OpenStreetMap.Mapnik, zoom=15)
ax2.grid(True)

# manager = plt.get_current_fig_manager()
# manager.window.state('zoomed')

# Adjust layout
plt.tight_layout()
plt.show()