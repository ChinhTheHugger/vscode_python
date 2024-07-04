import matplotlib.pyplot as plt
import matplotlib.animation as animation
import numpy as np
from matplotlib.patches import Polygon

# Define a function to create some example polygons
def create_example_polygons():
    polygons = [
        [(1, 1), (1, 4), (4, 4), (4, 1)],
        [(2, 2), (2, 5), (5, 5), (5, 2)],
        [(3, 3), (3, 6), (6, 6), (6, 3)]
    ]
    return polygons

# Create a figure and axis
fig, ax = plt.subplots()

# Set up the axis limits
ax.set_xlim(0, 7)
ax.set_ylim(0, 7)

# Initialization function
def init():
    ax.clear()
    return []

# Animation function
def animate(i):
    ax.clear()  # Clear the previous polygons
    polygons = create_example_polygons()
    current_polygon = polygons[i % len(polygons)]  # Loop through the polygons
    polygon_patch = Polygon(current_polygon, closed=True, fill=None, edgecolor='r')
    ax.add_patch(polygon_patch)
    return [polygon_patch]

# Create the animation
ani = animation.FuncAnimation(fig, animate, init_func=init, frames=30, interval=1000, blit=True)

# Show the animation
plt.show()
