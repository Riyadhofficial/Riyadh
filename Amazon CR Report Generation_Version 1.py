#!/usr/bin/env python
# coding: utf-8

# In[220]:


import pandas as pd
import numpy as np
import os
import glob
import datetime
from datetime import datetime
from datetime import date
from dateutil.relativedelta import relativedelta
from openpyxl.writer.excel import ExcelWriter
import datetime as dt     
date = dt.date.today
import warnings
warnings.filterwarnings("ignore")

# In[6]:


master_path="C:\\Users\\Hi\\Desktop\\PYTHON\\Master Data Only.xlsx"
master=pd.read_excel(master_path)
master = master.iloc[: ,1:]
master=master.drop(["Frequency"],axis=1)
path ="C:\\Users\\Hi\\Desktop\\PYTHON\\REPORTS"
path1 ="C:\\Users\\Hi\\Desktop\\PYTHON"
excel_files = glob.glob(os.path.join(path, "*.xlsx"))
file = []
print("All reports & master Data has been imported & categorised")

# In[7]:


for f in excel_files:
    df = pd.read_excel(f).assign(Category=f.split("\\")[-1].split("_")[0])
    file.append(df)


# In[158]:


Merged = pd.concat(file)
Merged.dropna(how='all', axis=1, inplace=True)


# In[159]:


Merged["Target-Campaign-Merge"]=Merged.apply(lambda x:'%s|%s' % (x['Target'],x['Campaign Title']),axis=1)

print("Targe-Campaign has been merged")
# In[160]:


Master_Data=pd.concat([master,Merged],ignore_index=True)
print("Consoldiated report has been merged with Master Data")

# In[161]:


Master_Data[['Target','Campaign']] = Master_Data['Target-Campaign-Merge'].str.split('|',expand=True)
print("Targe-Campaign has been unmerged")

# In[162]:


Master_Data["Frequency"]=pd.cut(Master_Data["Ad Spend"],bins=[0,5,10,20,30,40,50,60,70,80,90,100,200,300,400,500,600,700,800,900,1000],right=False,include_lowest=True)
print("Frequency range has been added for Ad Spend")

# In[163]:


table=Master_Data.pivot_table(Master_Data, index=['Target','Campaign','Category','Match Type'], columns=['To Date'], aggfunc={'CR':'mean'})


# In[164]:


table.columns = table.columns.droplevel(0) #remove amount
table.columns.name = None               #remove categories
table = table.reset_index()                #index to columns


# In[207]:


tabl1=table[table.columns[-5:]]
tabl2=table.iloc[: , :4]
final=tabl2.join(tabl1)
final=final.fillna(0)


# In[208]:


final=final[(final.iloc[:,4]<=15)]
final=final[(final.iloc[:,5]<=15)]
final=final[(final.iloc[:,6]<=15)]
final=final[(final.iloc[:,7]<=15)]
final=final[(final.iloc[:,8]<=15)]


# In[209]:


#final['sum']=final.iloc[:, -5:-1].sum(axis=1)


# In[216]:


final['sum']=final.sum(axis=1)
final = final[(final[['sum']] != 0).all(axis=1)]
final = final[(final[['sum']]>=0).all(axis=1)]
final=final.replace(0.0,'')
pivot=final.replace(-1.0,"")
pivot=pivot.drop(['sum'],axis =1)
pivot['Recommendation Date']=dt.datetime.today().strftime("%m/%d/%Y")
first_column = pivot.pop('Recommendation Date')
pivot.insert(0, 'Recommendation Date', first_column)
pivot=pivot[(pivot.iloc[:,9]!="")]#---------------------------------------###############
print("Final Pivot has been Generated")
mask = '%d-%m-%Y'#-------------------------------------------
dte = datetime.now().strftime(mask)#############################
fname = "Amazon_CR_Report_{}.xlsx".format(dte)


# In[227]:

os.chdir(path1)
with pd.ExcelWriter(fname) as writer:  
    Master_Data.to_excel(writer, sheet_name='Master Data',index=False)
    pivot.to_excel(writer, sheet_name='Pivot',index=False)
print("Output Saved")

# In[ ]:




