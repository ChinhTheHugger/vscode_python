import igraph
import math
import random
import numpy
import matplotlib.pyplot as plt

def is_power_of_ten(num):
    if num < 1:
        return False
    elif not math.log10(num).is_integer():
        return False
    elif math.log10(num)%2 != 0:
        return False
    else:
        return True

def input_number_of_prisoners():
    while True:
        try:
            num = int(input("Enter the number of prisoners in this experiment: "))
            if is_power_of_ten(num):
                return num
            else:
                print("Invalid input")
        except ValueError:
            print("Invalid input")

def initialize_random_array(num):
    numbers = list(range(1, num+1))
    random.shuffle(numbers)
    return numbers

def print_box_formation(list,num):
    edge = int(math.sqrt(num))
    for y in range(1,edge):
        for x in range(edge*(y-1),edge*y):
            print(list[x+y],end="  ")
        print("\n")
        
def set_up_graph(list,num):
    graph = igraph.Graph(n=num+2)
    arr = numpy.array(list)
    graph.vs['boxes'] = arr
    for i in range(1,num):
        graph.add_edges([(int(i), int(list[i]))])
    return graph

num_of_prisoners = input_number_of_prisoners()
fail_criteria = num_of_prisoners/2

box_placement = initialize_random_array(num_of_prisoners)

print_box_formation(box_placement,num_of_prisoners)

graph = set_up_graph(box_placement,num_of_prisoners)

fig, ax = plt.subplots()

visual_style = {}
visual_style["vertex_size"] = 10
visual_style["vertex_color"] = "lightblue"
visual_style["vertex_label"] = graph.vs.indices

visual_style["vertex_label_dist"] = 3
visual_style["vertex_label_angle"] = 0 

igraph.plot(graph,target=ax,**visual_style)
plt.show()

for sub_g in graph.components().subgraphs():
    print(sub_g)