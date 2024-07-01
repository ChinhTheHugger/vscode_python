import random
from PIL import Image

"""
Implements the flood fill algorithm in Python with depth-first search.

Inputs:
	-image file
	-new image file to write the new image to
	-threshold for the color matching
	-new color = color to update pixels to
	-start node = if none, picks a random starting point
"""


class FloodFill:
	def __init__(self, image_file,new_image_file,threshold,new_color,start=None):
		self.im = self.get_image(image_file)
		self.new_image_file = new_image_file
		self.pixels = self.im.load()
		self.width,self.height = self.im.size
		self.threshold = threshold
		self.new_color = new_color
		self.start = self.get_start_point(start)
		self.fill_pixels = self.fill()
		self.update_pixels()
		self.save_image()

	def get_start_point(self,start):
		"""returns a random starting point if one is not supplied"""
		if not start:
			start = (random.randint(0,self.width-1),random.randint(0,self.height-1))
		return start

	def get_image(self,image_file):
		"""returns a PIL image object"""
		im = Image.open(image_file)
		return im

	def is_similar(self,start_color,new_color):
		"""returns True if two color differences are within threshold"""
		red,green,blue,alpha = start_color
		new_red,new_green,new_blue,new_alpha = new_color
		if abs(red-new_red) <= self.threshold and abs(green-new_green) <= self.threshold and abs(blue-new_blue) <= self.threshold:
			return True
		return False

	def fill(self):
		"""depth-first search, returns pixels to change colors"""
		seen = []
		stack = [self.start]
		start_color = self.im.getpixel(self.start)
		while stack:
			node = stack.pop()
			seen.append(node)
			x = node[0]
			y = node[1]
			color = self.im.getpixel(node)
			if self.is_similar(start_color,color):
				#neighbors - includes diagonals
				neighbors = [(x-1,y),(x+1,y),(x-1,y-1),(x+1,y+1),(x-1,y+1),(x+1,y-1),(x,y-1),(x,y+1)]
				for n in neighbors:
					if 0 <= n[0] <= self.width-1 and 0 <= n[1] <= self.height-1:
						if n not in seen and self.is_similar(start_color,self.im.getpixel(n)):
							stack.append(n)
		return seen

	def update_pixels(self):
		"""updates pixel colors"""
		for (i,j) in self.fill_pixels:
			self.pixels[i,j] = self.new_color
		return

	def save_image(self):
		self.im.save(self.new_image_file)
		return