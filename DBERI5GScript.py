#!/usr/bin/env python
# coding: utf-8

# In[5]:


import os
import pandas as pd
import warnings
from openpyxl import Workbook
import xlwings as xw
import win32com.client
import openpyxl
import datetime
from datetime import datetime
import random
import time

cwd = os.getcwd()
in_dir = cwd
in_files = os.listdir(in_dir)
#out_dir = cwd+"\\"+"Output"
#out_files = os.listdir(out_dir)


# In[6]:


df_infile = pd.read_excel((in_dir+"\\"+"Ericsson_5G_Daily_Dashboard_Report.xlsx"),sheet_name ='RAW')
#df_rawcsv = pd.read_csv((in_dir+"\\"+"5G_NW_Performance_Daily_Report_Input.csv"))
#rawreport = pd.ExcelWriter((in_dir+"\\"+"5G_NW_Performance_Daily_Report_Input.xlsx"))
#df_rawcsv.to_excel(rawreport, index = False, header = True)
#rawreport.save()
df_rawdata = pd.read_excel((in_dir+"\\"+"5G_NW_Performance_Daily_Report_Input.xlsx"))
df_rawdata['DATE_ID'] = pd.to_datetime(df_rawdata['DATE_ID'])
df_rawdata['DATE_ID'] = df_rawdata['DATE_ID'].dt.strftime("%d/%m/%Y")
df_rawdata.head(5)



# In[7]:


df_rawdata_dr = pd.read_excel((in_dir+"\\"+"5G_SgNB_Drop_Rate_Daily_Report_Input.xlsx"),sheet_name ='Report 1')
df_rawdata_dr['Date'] = pd.to_datetime(df_rawdata_dr['Date'])
df_rawdata_dr['Date'] = df_rawdata_dr['Date'].dt.strftime("%d/%m/%Y")
#df_rawdata_dr.head(5)
#Merge
#df_new =  pd.merge(df_rawdata,df_rawdata_dr,how='left',on='Date',sort=False)
df_rawdata = pd.concat([df_rawdata, df_rawdata_dr], axis=1, join='inner')
df_rawdata=df_rawdata.drop(['Date', 'Hour'], axis=1,)
#df_rawdata.to_excel((in_dir+"\\"+"5G_NW_Performance_Daily_Report_Input.xlsx"),sheet_name = 'SubReport 1',index = False)
df_rawdata.head(4)


# In[8]:


df_rawdata['DATE_ID'] = pd.to_datetime(df_rawdata['DATE_ID'])
df_rawdata.insert(43, "Total Data Vol(Gb)","")
df_rawdata['Total Data Vol(Gb)']= df_rawdata['NR DL MAC Vol (Gbyte)'] + df_rawdata['NR UL MAC Vol (Gbyte)']
df_rawdata.insert(44, "DATE","")
df_rawdata['DATE'] = df_rawdata['DATE_ID'].dt.strftime("%m/%d/%Y")
df_rawdata['DATE'] = pd.to_datetime(df_rawdata['DATE'])
df_rawdata.insert(45, "Ref_HR","")
df_rawdata['Ref_HR'] = df_rawdata['HOUR_ID']
df_rawdata.insert(46, "Day","")
df_rawdata['Day'] = df_rawdata['DATE'].dt.day_name()
df_rawdata.insert(47, "Week","")
df_rawdata['Week'] = df_rawdata['DATE'].dt.isocalendar().week
df_rawdata.insert(48, "Year","")
df_rawdata['Year'] = df_rawdata['DATE'].dt.isocalendar().year
df_rawdata.insert(49, "Ref_WK","")
df_rawdata['Ref_WK'] = 'Week ' + df_rawdata['Week'].astype(str)

df_rawdata.head(8)


# In[9]:


df = df_infile.append(df_rawdata, ignore_index=True)
df.tail()


# In[10]:


df.shape
Infile = in_dir+"\\"+"Ericsson_5G_Daily_Dashboard_Report.xlsx"
wb1 = openpyxl.load_workbook(os.path.join(in_dir+"\\"+"Ericsson_5G_Daily_Dashboard_Report.xlsx"))
delsheet=wb1['RAW']
wb1.remove(delsheet)
wb1.save(filename = Infile)
#Infile = in_dir+"\\"+"5G Daily Dashboard_rev2.xlsx"
writer = pd.ExcelWriter(Infile, engine='openpyxl',mode='a',sheet_name='RAW')
writer.book = wb1
df.to_excel(writer, sheet_name='RAW',header=True, index=False,startrow=0,startcol=0)
writer.save()
writer.close()
wb1.save(filename = Infile)
wb1.close()


# In[11]:


in_dir = cwd
File = win32com.client.Dispatch("Excel.Application")
File.DisplayAlerts = 1
#File.visible = True
xlbook = File.Workbooks.Open(in_dir+"\\"+"Ericsson_5G_Daily_Dashboard_Report.xlsx")
xlbook.RefreshAll()
time.sleep(5)
print("Charts Refreshed")
xlbook.Save()
File.Quit()


# In[ ]:




