import random

# Maze dimensions
n = 16  # original maze width
m = 16  # original maze height

# Double the dimensions for path generation
maze_width = n * 2 + 1
maze_height = m * 2 + 1

# Initialize maze with walls (0)
maze = [[0 for _ in range(maze_height)] for _ in range(maze_width)]

# Function to find a random spot in the middle of the maze
def find_random_middle_spot():
    while True:
        x = random.randint(1, maze_width - 2)
        y = random.randint(1, maze_height - 2)
        if (x % 2 == 1) and (y % 2 == 1):  # Ensure it's a valid path position
            return (x, y)

# Starting point at top left corner
start_point = (1, 1)

# Function to check if a position is valid
def is_valid(x, y):
    return 0 <= x < maze_width and 0 <= y < maze_height and maze[x][y] == 0

# DFS to generate a path to an end point
def dfs(x, y, ex, ey):
    if (x, y) == (ex, ey):
        return True
    
    directions = [(2, 0), (-2, 0), (0, 2), (0, -2)]
    random.shuffle(directions)
    
    for dx, dy in directions:
        nx, ny = x + dx, y + dy
        if is_valid(nx, ny):
            maze[nx][ny] = 1
            maze[(x + nx) // 2][(y + ny) // 2] = 1  # Remove wall between cells
            if dfs(nx, ny, ex, ey):
                return True
            else:
                maze[nx][ny] = 0  # Backtrack if no valid path found
    
    return False

# Generate one true end point and several pseudo end points
true_end_point = find_random_middle_spot()
pseudo_end_points = [(1, 1), (maze_width - 2, 1), (1, maze_height - 2), (maze_width - 2, maze_height - 2)]

# Additional pseudo end points at the middle points of outer walls
for x in range(3, maze_width - 2, 2):
    pseudo_end_points.append((x, 1))                # Top row
    pseudo_end_points.append((x, maze_height - 2))  # Bottom row

for y in range(3, maze_height - 2, 2):
    pseudo_end_points.append((1, y))                # Left column
    pseudo_end_points.append((maze_width - 2, y))   # Right column

# Generate paths to true end point and pseudo end points
maze[start_point[0]][start_point[1]] = 1
dfs(start_point[0], start_point[1], true_end_point[0], true_end_point[1])
for ex, ey in pseudo_end_points:
    maze[start_point[0]][start_point[1]] = 1
    dfs(start_point[0], start_point[1], ex, ey)

# Function to print the maze to the console
def print_maze():
    for y in range(maze_height):
        for x in range(maze_width):
            if (x, y) == start_point:
                print('S ', end='')
            elif (x, y) == true_end_point:
                print('E ', end='')
            elif (x, y) in pseudo_end_points:
                print('  ', end='')
            else:
                print('  ' if maze[x][y] == 1 else '# ', end='')
        print()

# Print the maze to the console
print_maze()
