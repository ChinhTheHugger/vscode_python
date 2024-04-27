import numpy as np
import pandas as pd
import seaborn as sn
import matplotlib.pyplot as plt

path = "C:\\Users\\phams\Downloads\\train.csv"

df = pd.read_csv(path, usecols=['Survived','Pclass','Age','Sex'])

corrMatrix = df.corr()
sn.heatmap(corrMatrix, annot=True)
plt.show()
