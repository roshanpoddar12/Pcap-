import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
import os


plt.rc('font',**{'family':'sans-serif','sans-serif':['DejaVu Sans'],'size':7})
X_AXIS=[]
ws=[0 for i in range(20)]
da={}
xtick=[]
sum=-0.4
i=0
wb = openpyxl.load_workbook('out.xlsx')
groups = {}
sheet = wb.sheetnames
for i in range(len(sheet)):
    ws = wb.worksheets[i]
    
    df = pd.read_excel(open('out.xlsx','rb'), sheet_name=sheet[i])
    
    df1=pd.DataFrame(df)
    path=(os.getcwd()+"/"+sheet[i])
    column_list= df['new_col_name']
    os.makedirs(path, exist_ok=True)
    df.set_index("new_col_name",drop=False,inplace=True)
    for j, c in enumerate(df.columns):
        if j>=2:

          
           sum=-0.4
           xtick=[]
           #plt.bar(range(len(df)), df[c], color=plt.cm.Paired(np.arange(len(df))),label=df['new_col_name'])
           ax = df[[c]].T.plot(kind='bar',  fontsize=10,label=df['new_col_name'], colormap='Paired',title=c,width=1)
           ax.set_xlim(-1, 1)
           print(len(df))
           values=1.0/len(df)
           print(values)
           xtick.append(sum)
           for k in range(1,len(df)):
              sum=sum+values
              xtick.append(sum)
              print(xtick)
           ax.set_xticks(xtick)
           ax.set_xticklabels(range(len(df)))
           
           plt.savefig(os.getcwd()+'/'+sheet[i]+'/'+c+'.png', bbox_inches='tight')
           plt.show()
    
    
