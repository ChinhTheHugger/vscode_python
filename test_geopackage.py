import geopandas as gpd # đọc file dữ liệu địa lý
import matplotlib.pylab as plt
import contextily as ctx # hiển thị map làm nền
# tải về 3 thư viện: pip install geopandas matplotlib contextily

file_path = "C:\\Users\\phams\\Downloads\\gadm41_VNM_3.json"

gdf = gpd.read_file(file_path, driver='GeoJSON')

print(gdf.head(20))  # in 20 dòng đầu tiên
print("\nTotal rows: " + str(len(gdf)) + "\n")

print("Level 1 list:")
print(gdf['NAME_1'].unique())
print("\nTotal level 1: " + str(len(gdf['NAME_1'].unique())) + "\n")

gdf_temp = gdf[gdf['NAME_1'] == 'NinhBình'] # Lọc các dữ liệu của địa chỉ level 1
specific_gdf = gdf_temp[gdf_temp['VARNAME_3'] == 'NinhVan'] # Lọc dữ liệu địa chỉ level 3 con của địa chỉ level 1 bên trên
geometry =  specific_gdf.geometry.values[0]

all_coordinates = []

if geometry.geom_type == 'MultiPolygon':
    # Iterate over each Polygon in the MultiPolygon using the geoms attribute
    all_coordinates = []
    for polygon in geometry.geoms:
        exterior_coords = list(polygon.exterior.coords)
        all_coordinates.append(exterior_coords)
else:
    # If it's not a MultiPolygon, handle it as a single Polygon
    all_coordinates = [list(geometry.exterior.coords)]
print(all_coordinates) # in multipolygon của địa chỉ level 3 bên trên

# Plot the filtered GeoDataFrame with customization
if not specific_gdf.empty:
    # Plot the specific GeoDataFrame over a basemap
    fig, ax = plt.subplots(figsize=(10, 10))
    specific_gdf.plot(ax=ax, alpha=0.5, edgecolor='k', color='orange')
    ctx.add_basemap(ax, source=ctx.providers.OpenStreetMap.Mapnik, zoom=12)  # Adjust zoom level if necessary
    plt.show()
