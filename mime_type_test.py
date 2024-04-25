import mimetypes

file_path = r'C:\Users\Public\Documents\car_shop\uploads\fronts\2019-toyota-sienna-front-view-driving-carbuzz-613433-1600-844858838.jpg'
mime_type, encoding = mimetypes.guess_type(file_path)
if 'image' in mime_type:
    print("yes")
else:
    print("no")