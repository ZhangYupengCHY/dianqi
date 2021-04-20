import pandas as pd



a = pd.DataFrame([[1,2,3],[3,4,5],[3,4,6]])
showRow = None
for row,row_value in a.iterrows():
     if 4 in row_value.values:
        showRow = row
        break
print(showRow)

b = set([1,2,3])
d = set([1,2,4,5])
f = b | d
print(f)