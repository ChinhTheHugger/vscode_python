import numpy as np
import math
from shapely.geometry import LineString

def haversine(coord1, coord2):
    lat1, lon1 = coord1
    lat2, lon2 = coord2

    R = 6371e3  # Earth radius in meters
    phi1 = math.radians(lat1)
    phi2 = math.radians(lat2)
    delta_phi = math.radians(lat2 - lat1)
    delta_lambda = math.radians(lon2 - lon1)

    a = math.sin(delta_phi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(delta_lambda / 2) ** 2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

    distance = R * c  # in meters
    return distance

def interpolate_route(route, num_points):
    total_length = 0
    distances = []
    for i in range(1, len(route)):
        distance = haversine(route[i-1], route[i])
        distances.append(distance)
        total_length += distance
    
    interpolated_distances = np.linspace(0, total_length, num_points)
    
    interpolated_route = [route[0]]
    current_length = 0
    for i in range(1, len(route)):
        segment_length = distances[i-1]
        while len(interpolated_route) < num_points and (current_length + segment_length) >= interpolated_distances[len(interpolated_route)]:
            ratio = (interpolated_distances[len(interpolated_route)] - current_length) / segment_length
            lat = route[i-1][0] + ratio * (route[i][0] - route[i-1][0])
            lon = route[i-1][1] + ratio * (route[i][1] - route[i-1][1])
            interpolated_route.append((lat, lon))
        current_length += segment_length
    
    if len(interpolated_route) < num_points:
        interpolated_route.append(route[-1])
    
    return interpolated_route

def calculate_average_gap(interpolated_route1, interpolated_route2):
    total_distance = 0
    count = len(interpolated_route1)
    
    for i in range(count):
        distance = haversine(interpolated_route1[i], interpolated_route2[i])
        total_distance += distance
    
    average_gap = total_distance / count if count > 0 else 0
    return average_gap

# Provided routes
line1 = LineString([
    (105.7876, 20.98015), (105.7874, 20.98), (105.78756, 20.97983), (105.78804, 20.98021), 
    (105.78857, 20.98061), (105.78967, 20.98151), (105.79009, 20.98182), (105.79244, 20.98371), 
    (105.7933, 20.98435), (105.79382, 20.98471), (105.79531, 20.98577), (105.79983, 20.98896), 
    (105.80009, 20.98915), (105.8002, 20.98925), (105.80285, 20.99113), (105.8037, 20.99172), 
    (105.80654, 20.99374), (105.80672, 20.9938), (105.80768, 20.99448), (105.80824, 20.99487), 
    (105.80925, 20.99555), (105.81118, 20.99684), (105.81236, 20.99761), (105.81328, 20.99823), 
    (105.81409, 20.99879), (105.81451, 20.99906), (105.81551, 20.99974), (105.816, 21.00007), 
    (105.81766, 21.00115), (105.81775, 21.00125), (105.81836, 21.0017), (105.81874, 21.00196), 
    (105.81909, 21.00222), (105.82, 21.00289), (105.82126, 21.00385), (105.82176, 21.00425), 
    (105.82201, 21.00451), (105.8222, 21.00474), (105.82232, 21.00485), (105.82263, 21.00549), 
    (105.82278, 21.00583), (105.82288, 21.00601), (105.8232, 21.00676), (105.82383, 21.00805), 
    (105.82392, 21.00818), (105.82426, 21.00882), (105.82479, 21.00897), (105.82495, 21.00897), 
    (105.82512, 21.00894), (105.82524, 21.00889), (105.82799, 21.00766), (105.82872, 21.00732), 
    (105.82948, 21.00698), (105.82967, 21.00687), (105.83062, 21.00644), (105.83078, 21.00635), 
    (105.83095, 21.00625), (105.8311, 21.00618), (105.83119, 21.00615), (105.83128, 21.00622), 
    (105.83142, 21.00631), (105.83218, 21.00691), (105.83326, 21.00779), (105.83333, 21.00788), 
    (105.83368, 21.00817), (105.83372, 21.00826), (105.83383, 21.00835), (105.83401, 21.00816)
])

line2 = LineString([
    (105.7876, 20.98016), (105.78739, 20.97999), (105.78756, 20.97982), (105.78856, 20.98061), 
    (105.78966, 20.9815), (105.79018, 20.98188), (105.79245, 20.98371), (105.79334, 20.98437), 
    (105.79503, 20.98557), (105.79529, 20.98574), (105.79594, 20.98622), (105.80009, 20.98915), 
    (105.80018, 20.98925), (105.80079, 20.9897), (105.80629, 20.99354), (105.80643, 20.9936), 
    (105.8066, 20.99371), (105.81651, 21.00042), (105.81766, 21.00115), (105.81775, 21.00125), 
    (105.81908, 21.00221), (105.8208, 21.0035), (105.82164, 21.00416), (105.822, 21.0045), 
    (105.82234, 21.00495), (105.82321, 21.00689), (105.82382, 21.00804), (105.82446, 21.00925), 
    (105.82576, 21.00868), (105.83096, 21.00628), (105.83123, 21.00617), (105.83217, 21.0069), 
    (105.83302, 21.00759), (105.83368, 21.00815), (105.83371, 21.00825), (105.83383, 21.00835), 
    (105.83401, 21.00815)
])

# Convert LineStrings to list of (latitude, longitude) tuples
route1 = [(lat, lon) for lon, lat in line1.coords]
route2 = [(lat, lon) for lon, lat in line2.coords]

# Debugging outputs
print("Route 1 coordinates:", route1)
print("Route 2 coordinates:", route2)

# Determine number of points based on the route with more nodes
num_points = max(len(route1), len(route2))

# Interpolate routes
interpolated_route1 = interpolate_route(route1, num_points)
interpolated_route2 = interpolate_route(route2, num_points)

# Debugging outputs
print("Interpolated Route 1 coordinates:", interpolated_route1)
print("Interpolated Route 2 coordinates:", interpolated_route2)

# Calculate average gap
average_gap = calculate_average_gap(interpolated_route1, interpolated_route2)
print(f"The average gap between the routes with interpolation is {average_gap} meters.")
