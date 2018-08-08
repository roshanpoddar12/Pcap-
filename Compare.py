import pandas as pd
from pandas import *
import numpy as np
from pandas import ExcelWriter
import glob
import xlsxwriter
import os
j=0
i=0
list1 =[]
row=[[]]
xlsx=[]
frame = pd.DataFrame()
# get data file names
for f in glob.glob("*.xlsx"):
    xlsx.append(f)
    df = pd.read_excel(f,index_col=None, header=2)
    df.insert(0, "new_col_name", f.strip('.xlsx'))
    #print df
    list1.append(df)
    i=i+1
frame = pd.concat(list1,sort=False)
size=frame.shape
print size
size=size[0]/i
#print frame['pcapFileName']
writer = ExcelWriter('PythonExport.xlsx')
for k in frame['pcapFileName']:
    if j<size-1:
        frame.loc[j].to_excel(writer,k, index = False)
        j=j+1
writer.save()


