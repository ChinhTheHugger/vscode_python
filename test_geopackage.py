import geopandas as gpd # đọc file dữ liệu địa lý
import fastkml as kml
import matplotlib.pylab as plt
import contextily as ctx # hiển thị map làm nền
import zipfile
import os

file_path = "C:\\Users\\phams\\Downloads\\gadm41_VNM_3.json"

def load_kmz(file_path):
    with zipfile.ZipFile(file_path, 'r') as kmz:
        kmz.extractall('extracted')
    kml_file = None
    for root, dirs, files in os.walk('extracted'):
        for file in files:
            if file.endswith('.kml'):
                kml_file = os.path.join(root, file)
                break
    if kml_file:
        return kml_file
    else:
        print("No KML file found in the KMZ archive.")
        return None

def kml_to_gdf(kml_file):
    with open(kml_file, 'rt', encoding='utf-8') as f:
        doc = f.read()
    k = kml.KML()
    k.from_string(doc)
    features = list(k.features())
    placemarks = list(features[0].features())
    geoms = []
    names = []
    for placemark in placemarks:
        geoms.append(placemark.geometry)
        names.append(placemark.name)
    gdf = gpd.GeoDataFrame({'name': names, 'geometry': geoms})
    return gdf

def load_geodata(file_path):
    if file_path.endswith('.shp'):
        return gpd.read_file(file_path)
    elif file_path.endswith('.geojson') or file_path.endswith('.json'):
        return gpd.read_file(file_path)
    elif file_path.endswith('.gpkg'):
        return gpd.read_file(file_path)
    elif file_path.endswith('.kmz'):
        kml_file = load_kmz(file_path)
        if kml_file:
            return kml_to_gdf(kml_file)
        else:
            return None
    else:
        print("Unsupported file format.")
        return None

def main():
    gdf = load_geodata(file_path)

    if gdf is not None:
        while True:
            print(gdf.head())  # in 20 dòng đầu tiên
            print("\nTotal rows: " + str(len(gdf)) + "\n")

            print("Level 1 list:")
            print(gdf['NAME_1'].unique())
            print("\nTotal level 1: " + str(len(gdf['NAME_1'].unique())) + "\n")

            while True:
                name_1 = input("Select a level 1 name: ")
                gdf_lv1 = gdf[gdf['NAME_1'] == name_1] # Lọc các dữ liệu của địa chỉ level 1
                if gdf_lv1.empty == False:
                    print("Level 2 list from selected level 1:")
                    for name in gdf_lv1['NAME_2'].unique():
                        print(name)
                    print("Total level 2: " + str(len(gdf_lv1['NAME_2'].unique())) + "\n")
                    break

            while True:
                name_2 = input("Select a level 2 name: ")
                gdf_lv2 = gdf_lv1[gdf_lv1['NAME_2'] == name_2] # Lọc các dữ liệu của địa chỉ level 2 từ địa chỉ level 1 đã chọn
                if gdf_lv2.empty == False:
                    print("Level 3 list from selected level 2:")
                    for name in gdf_lv2['NAME_3'].unique():
                        print(name)
                    print("Total level 3: " + str(len(gdf_lv2['NAME_3'].unique())) + "\n")
                    break

            while True:
                name_3 = input("Select a level 3 name: ")
                specific_gdf = gdf_lv2[gdf_lv2['NAME_3'] == name_3] # Lọc dữ liệu địa chỉ level 3 con từ địa chỉ level 1 và 2 đã chọn
                geometry =  specific_gdf.geometry.values[0]

                all_coordinates = []
                
                if specific_gdf.empty == False:
                    print(name_3 + " boundary:")
                    if geometry.geom_type == 'MultiPolygon':
                        # Iterate over each Polygon in the MultiPolygon using the geoms attribute
                        all_coordinates = []
                        for polygon in geometry.geoms:
                            exterior_coords = list(polygon.exterior.coords)
                            all_coordinates.append(exterior_coords)
                    else:
                        # If it's not a MultiPolygon, handle it as a single Polygon
                        all_coordinates = [list(geometry.exterior.coords)]
                    for coords in all_coordinates:
                        print(coords) # in multipolygon của địa chỉ level 3 bên trên
                        print("\n")

                    # Plot the filtered GeoDataFrame with customization
                    if not specific_gdf.empty:
                        # Plot the specific GeoDataFrame over a basemap
                        fig, ax = plt.subplots(figsize=(10, 10))
                        specific_gdf.plot(ax=ax, alpha=0.5, edgecolor='k', color='orange')
                        ctx.add_basemap(ax, source=ctx.providers.OpenStreetMap.Mapnik, zoom=12)  # Adjust zoom level if necessary
                        plt.show()
                    
                    break
            
            redo = input("Do you want to filter another name? (y/n): ").strip().lower()
            if redo != 'y':
                print("Exiting the program.")
                break
    
if __name__ == "__main__":
    main()
