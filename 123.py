import pandas as pd
import numpy as np

n = 80000
array = np.zeros((n,n))
pd = pd.DataFrame(array)
pd.to_csv('matrix.csv')
print(pd)