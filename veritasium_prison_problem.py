import igraph
import math
import random
import numpy
import matplotlib.pyplot as plt

def is_valid(num):
    if num < 1:
        return False
    elif num%2 != 0:
        return False
    elif not math.sqrt(num).is_integer():
        return False
    else:
        return True

def input_number_of_prisoners():
    while True:
        try:
            num = int(input("Enter the number of prisoners in this experiment: "))
            if is_valid(num):
                return num
            else:
                print("The number of prisoners must be a positive even number with natural square root value")
        except ValueError:
            print("The number of prisoners must be a positive even number with natural square root value")

def initialize_random_array(num):
    while True:
        numbers = list(range(1, num+1))
        random.shuffle(numbers)
        if all(numbers[i] != i for i in range(num)):
            return numbers

def print_box_formation(list,num):
    edge = int(math.sqrt(num))
    for y in range(1,edge+1):
        for x in range(edge*(y-1),edge*y):
            print(str("("+str(x+1)+") "+str(list[x])),end="  ")
        print("\n")
    return
        
def set_up_graph(prisoner_num,box_num,num):
    graph = igraph.Graph(n=num+1)
    arr = numpy.array(box_num)
    graph.vs['boxes'] = arr
    edges = []
    for i in range(0,num):
        edges.append([int(prisoner_num[i]),int(box_num[i])])
    graph.add_edges(edges)
    return graph

def check_result(graph,fail_criteria):
    subgraph_count = []
    for idx, subgraph in enumerate(graph.components().subgraphs()):
        print(f"Loop {idx+1} has {subgraph.vcount()} boxes")
        subgraph_count.append(subgraph.vcount())

    if all(subgraph_count[i] <= fail_criteria for i in range(len(subgraph_count))):
        print("\nThe prisoners have earned their freedom")
    else:
        print("\nLuck was not on the prisoners' side this day")
    
    return

def check_result_boolean(graph,fail_criteria):
    subgraph_count = []
    for idx, subgraph in enumerate(graph.components().subgraphs()):
        subgraph_count.append(subgraph.vcount())

    if all(subgraph_count[i] <= fail_criteria for i in range(len(subgraph_count))):
        return 1
    else:
        return 0

def calculate_success_rate():
    list_num_of_prisoners = [100,400,900,1024,2500] # sample size must be a positive even number with natural square root value
    num_of_run = input("Number of run for each set of prisoners: ")

    results = []

    for num in list_num_of_prisoners:
        for run in range(int(num_of_run)):
            fail_criteria = num/2
            
            box_placement = initialize_random_array(num)
            
            graph = set_up_graph(box_placement,sorted(box_placement),num)
            
            results.append(check_result_boolean(graph,fail_criteria))

    print(f"Total runs: {len(results)}")
    print(f"Success rate: {sum(results)*100/len(results)}%")
    
    return

def normal_run_with_data_output():
    num_of_prisoners = input_number_of_prisoners()
    fail_criteria = num_of_prisoners/2

    box_placement = initialize_random_array(num_of_prisoners)

    print_box_formation(box_placement,num_of_prisoners)

    graph = set_up_graph(box_placement,sorted(box_placement),num_of_prisoners)
    
    check_result(graph,fail_criteria)

    # plotting the graph
    
    fig, ax = plt.subplots()

    ax.set_aspect('equal')

    visual_style = {}
    visual_style["vertex_size"] = 10
    visual_style["vertex_color"] = "lightblue"
    visual_style["vertex_label"] = graph.vs.indices

    visual_style["vertex_label_dist"] = 3
    visual_style["vertex_label_angle"] = 0

    igraph.plot(graph,target=ax,**visual_style)

    manager = plt.get_current_fig_manager()
    manager.window.state('zoomed')
        
    plt.show()
    
    return


# for running one case, the number of prisoners is from user input
normal_run_with_data_output()

# for running multiple case for each sample size: 100, 400, 900, 1024, 2500
# the number of cases for each sample size is from user input
# adjust the number of sample and sample sizes in calculate_success_rate()
calculate_success_rate()