from shapely.geometry import LineString, Point, mapping, shape
import matplotlib.pyplot as plt
import math
from scipy.optimize import fsolve
from shapely.ops import unary_union, transform
from pyproj import Proj, Transformer

# Function to convert degrees to radians
def degrees_to_radians(degrees):
    return degrees * math.pi / 180.0

# Function representing the equation
def equation(variable):
    return value - 2 * earth_radius_km * math.asin(math.sqrt(math.sin(variable / (2 * earth_radius_km)) ** 2 + math.cos(lat1_rad) * math.cos(lat1_rad + variable) * math.sin(variable / (2 * earth_radius_km)) ** 2))

# Define EPSG:4326 and a suitable projected coordinate system, e.g., UTM
proj_wgs84 = Proj(init='epsg:4326')
proj_utm = Proj(proj='utm', zone=18, ellps='WGS84')

# Define transformer for EPSG:4326 to UTM
transformer_to_utm = Transformer.from_proj(proj_wgs84, proj_utm)
transformer_to_wgs84 = Transformer.from_proj(proj_utm, proj_wgs84)

# Function to transform coordinates
def to_utm(geom):
    return transform(transformer_to_utm.transform, geom)

def to_wgs84(geom):
    return transform(transformer_to_wgs84.transform, geom)

# Example coordinates for the first and second linestrings, with coords1 being the main linestring
coords1 = [(21.00798, 105.83418), (21.00835, 105.83383), (21.00826, 105.83372), (21.00824, 105.83362), (21.00794, 105.83327), (21.00787, 105.83316), (21.00753, 105.83274), (21.00716, 105.83227), (21.00702, 105.83208), (21.00643, 105.83131), (21.00633, 105.83116), (21.00603, 105.83106), (21.00577, 105.83095), (21.00536, 105.83079), (21.00412, 105.83033), (21.00311, 105.83001), (21.00206, 105.8296), (21.00149, 105.82934), (21.00106, 105.82917), (21.00096, 105.82895), (21.0009, 105.82879), (21.00133, 105.82756), (21.00146, 105.82712), (21.00162, 105.82651), (21.00169, 105.82627), (21.00166, 105.82599), (21.00266, 105.82205), (21.00276, 105.82193), (21.00304, 105.82093), (21.00313, 105.82054), (21.00318, 105.82023), (21.00318, 105.82016), (21.00317, 105.82009), (21.00309, 105.81997), (21.0028, 105.81956), (21.00248, 105.81911), (21.00189, 105.81829), (21.00179, 105.81823), (21.00156, 105.81785), (21.00112, 105.81721), (21.00099, 105.81712), (21.00067, 105.81668), (21.00061, 105.81644), (20.99894, 105.81394), (20.99819, 105.81278), (20.99755, 105.81183), (20.99661, 105.81038), (20.9957, 105.80901), (20.99498, 105.80796), (20.99413, 105.80676), (20.99397, 105.80656), (20.99386, 105.80649), (20.99361, 105.80616), (20.99183, 105.80362), (20.98935, 105.80009), (20.98927, 105.79994), (20.98713, 105.79688), (20.98617, 105.79554), (20.98536, 105.79436), (20.98452, 105.79314), (20.98432, 105.79288), (20.98366, 105.79205), (20.98324, 105.79149), (20.98272, 105.79084), (20.982, 105.78992), (20.98182, 105.78967), (20.98168, 105.7895), (20.98135, 105.78908), (20.98084, 105.78847), (20.98071, 105.78831), (20.98037, 105.78788), (20.98, 105.7874), (20.97939, 105.78666), (20.97933, 105.78659), (20.9792, 105.78643), (20.97913, 105.78633), (20.97906, 105.78624), (20.97785, 105.78479), (20.97634, 105.78282), (20.97578, 105.78207), (20.97504, 105.78105), (20.97485, 105.78082), (20.97472, 105.78068), (20.97464, 105.78057), (20.97407, 105.77977), (20.9735, 105.77891), (20.97334, 105.77869), (20.97152, 105.77638), (20.971, 105.77576), (20.9707, 105.77542), (20.96975, 105.77423), (20.96899, 105.77326), (20.96825, 105.77232), (20.9676, 105.77153), (20.96734, 105.77117)]
coords2 = [(21.00798, 105.83417), (21.00835, 105.83383), (21.00825, 105.83371), (21.00823, 105.83363), (21.00714, 105.8323), (21.00656, 105.83156), (21.00627, 105.83115), (21.00747, 105.82853), (21.00778, 105.82789), (21.00937, 105.8244), (21.00716, 105.82325), (21.00575, 105.82254), (21.0054, 105.82241), (21.00497, 105.82221), (21.00456, 105.82194), (21.00437, 105.82177), (21.00428, 105.82166), (21.00354, 105.82075), (21.00228, 105.81902), (21.00192, 105.81851), (21.00182, 105.81827), (21.00158, 105.81789), (20.9956, 105.80894), (20.99497, 105.80796), (20.99411, 105.80677), (20.99375, 105.80623), (20.99359, 105.80616), (20.9918, 105.80364), (20.98973, 105.80064), (20.98933, 105.8001), (20.98927, 105.79993), (20.98479, 105.79357), (20.98389, 105.79233), (20.98199, 105.78991), (20.98181, 105.78967), (20.98075, 105.78832), (20.98059, 105.78811), (20.97965, 105.78698), (20.97907, 105.78623), (20.97804, 105.78502), (20.97724, 105.78404), (20.97499, 105.78099), (20.97427, 105.78006), (20.97408, 105.77976), (20.9735, 105.7789), (20.9732, 105.77852), (20.97156, 105.77643), (20.971, 105.77575), (20.97069, 105.77542), (20.97023, 105.77485), (20.96825, 105.77232), (20.96735, 105.77116)]

# Create the linestrings
line1 = LineString([lon,lat] for lat, lon in coords1)
line2 = LineString([lon,lat] for lat, lon in coords2)

# Buffer width in degrees
buffer_width_degrees = 0.0125

# Create a buffer around the first linestring (width of 0.1 units)
buffer = line1.buffer(buffer_width_degrees)

# Determine the total length of the second linestring
line2_length = line2.length

# Transform the linestring to UTM
linestring_utm = to_utm(line1)

# Coordinates for conversion (assuming a location near the equator for simplicity)
lat1 = 0  # Starting latitude (in degrees)
lon1 = 0  # Starting longitude (in degrees)

# Convert degrees to radians
lat1_rad = degrees_to_radians(lat1)
lon1_rad = degrees_to_radians(lon1)

# Calculate haversine distance for buffer width
earth_radius_km = 6371.0  # Earth's radius in kilometers
buffer_width_haversine_km = 2 * earth_radius_km * math.asin(math.sqrt(math.sin(buffer_width_degrees / (2 * earth_radius_km)) ** 2 + math.cos(lat1_rad) * math.cos(lat1_rad + buffer_width_degrees) * math.sin(buffer_width_degrees / (2 * earth_radius_km)) ** 2))

# Convert haversine distance from kilometers to meters
buffer_width_haversine_meters = buffer_width_haversine_km * 1000

print(f"Buffer width in haversine distance (meters): {buffer_width_haversine_meters:.2f} meters")

# Given constants
value = 0.015 # kilometer
earth_radius_km = 6370 # kilometer

# # Solve for variable
# initial_guess = 0.02  # Initial guess for variable
# variable_solution = fsolve(equation, initial_guess)[0]
# print(variable_solution)
# print(f"The solution for variable is approximately: {variable_solution:.2f} units")

# # Iterate through each point
# buffer_v2 = []
# for point in line1.coords:
#     long, lat =  point
#     lat1_rad = degrees_to_radians(lat)
#     initial_guess = 0.001
#     buffer_approx = fsolve(equation, initial_guess)[0]
#     print(buffer_approx)
#     buffer_v2.append(Point(point).buffer(buffer_approx))
    
# continuous_buffer = unary_union(buffer_v2)

# Create the buffer in UTM coordinates
buffer_distance = 15  # Example buffer distance in meters
buffer_utm = linestring_utm.buffer(buffer_distance)

# Transform the buffer back to EPSG:4326 (resulting coordinates are in longitude, latitude)
buffer_wgs84 = to_wgs84(buffer_utm)

# Calculate the intersection of the second linestring with the buffer
intersection = line2.intersection(buffer_wgs84)

# Determine the length of the intersecting part
intersection_length = intersection.length

# Calculate the percentage of the second linestring that is within the buffer
percentage_within_buffer = (intersection_length / line2_length) * 100

print(f"Percentage of the second linestring within the buffer: {percentage_within_buffer:.2f}%")

# Create a plot
fig, ax = plt.subplots()

# Plot the linestring 1
x, y = line1.xy
ax.plot(x, y, color='blue', linewidth=1, label='Linestring 1')

# Plot the linestring 2
x, y = line2.xy
ax.plot(x, y, color='red', linewidth=1, label='Linestring 2')

# Plot the buffer
x, y = buffer_wgs84.exterior.xy
ax.fill(x, y, alpha=0.5, fc='red', ec='none', label='Buffer 1')

ax.legend()
plt.show()