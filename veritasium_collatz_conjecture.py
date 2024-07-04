import matplotlib.pyplot as plt
import polyline



def is_valid(start_num, end_num):
    if start_num == 0 or end_num == 0:
        return False
    elif start_num > end_num:
        return False
    elif not isinstance(start_num, int) or not isinstance(end_num, int):
        return False
    else:
        return True

def get_range():
    while True:
        try:
            start = int(input("Start value of the range: "))
            end = int(input("End value of the range: "))
            if is_valid(start,end):
                return start, end
            else:
                print("Start value and end value must be integer, and end value cannot be smaller than start value")
        except:
            print("Start value and end value must be integer, and end value cannot be smaller than start value")

def collatz_conjecture(n):
    """Generate the Collatz sequence for a given integer n."""
    if n > 0:
        sequence = []
        pos = 1
        while n != 1:
            sequence.append((pos,n))
            if n % 2 == 0:
                n = n // 2
                pos += 1
            else:
                n = 3 * n + 1
                pos += 1
        sequence.append((pos,1))  # Add the final 1 to the sequence
        return sequence
    if n < 0:
        if abs(n) in (1,5,17):
            sequence = [(1,n)]
            return sequence
        else:
            sequence = []
            pos = 1
            while abs(n) not in (1,5,17):
                sequence.append((pos,n))
                if n % 2 == 0:
                    n = n // 2
                    pos += 1
                else:
                    n = 3 * n + 1
                    pos += 1
            sequence.append((pos,n))  # Add the final 1 to the sequence
            return sequence

    

start, end = get_range()

collatz_group = []

for num in range(start,end+1):
    if abs(num) == 1 or abs(num) == 0:
        continue
    else:
        collatz_list = collatz_conjecture(num)
        collatz_group.append(collatz_list)
        
for group in collatz_group:
    x, y = zip(*group)
    plt.plot(x, y, marker='o', linestyle='-', markersize=1, linewidth=0.5)

plt.xlabel("Step")
plt.ylabel("Value")
plt.grid(True)

manager = plt.get_current_fig_manager()
manager.window.state('zoomed')

plt.show()