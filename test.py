import numpy as np

arr = np.array([[1, 2, 3, 4], [5, 6, 7, 8], [9, 10, 11, 12]])
print(arr)

arr2 = np.delete(arr, [0, 2], 0)
print(arr2)
