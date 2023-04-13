import openpyxl
import pandas
from collections import Counter
from difflib import SequenceMatcher
from collections import OrderedDict
import time
import numpy
import igraph

import sys

pathTest = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\Book.XLSX"
path = "E:\\Pham Thanh Quyet - 23.12.2022\\DSKH 22.12.23\\VRS VRH\\23.04.13 Riverside+ Harmony Full - Tổng hợp khách hàng và căn V23 - for processing.XLSX"

wb_obj = openpyxl.load_workbook(path)
sheet_base = wb_obj.active

start = time.time()
print("Start time = " + time.strftime("%H:%M:%S", time.gmtime(start)))
time.sleep(5)

graph = igraph.Graph(n=8,edges=[(0,4),(1,4),(1,6),(2,5),(3,6),(3,7)])
graph.vs['name'] = ['H1','H2','H3','H4','P1','P2','P3','P4']

print(graph)

for s in graph.components().subgraphs():
    print(s.vs['name'])
    print(' | '.join(s.vs['name']))

wb_obj.save(path)

end = time.time()

print("Run time = " + time.strftime("%H:%M:%S", time.gmtime(end-start)))