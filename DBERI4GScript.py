#!/usr/bin/env python
# coding: utf-8

# In[1]:


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
import glob as glob

cwd = os.getcwd()
in_dir = cwd
in_files = os.listdir(in_dir)
#out_dir = cwd+"\\"+"Output"
#out_files = os.listdir(out_dir)


# In[2]:


df_infile = pd.read_excel((in_dir+"\\"+"Ericsson_4G_Daily_Dashboard_Report.xlsx"),sheet_name ='RAW')
df_rawdata = pd.read_excel((in_dir+"\\"+"Ericson_Daily_NW_Report_Input(Hourly)-4G.xlsx"),sheet_name ='NW')
df_rawdata['Period'] = pd.to_datetime(df_rawdata['Period'])
df_rawdata.insert(16, "Total Data Vol(Gb)","")
df_rawdata['Total Data Vol(Gb)']= df_rawdata['DL PDCCP Data Volume (Gbytes)'] + df_rawdata['UL PDCP Data Volume (Gbytes)']
df_rawdata.insert(17, "DATE","")
df_rawdata['DATE'] = df_rawdata['Period'].dt.strftime("%m/%d/%Y")
df_rawdata.insert(18, "Hour","")
df_rawdata['Hour'] = df_rawdata['Period'].dt.strftime("%H")
df_rawdata.insert(19, "Ref_HR","")
df_rawdata['Ref_HR'] = df_rawdata['Period'].dt.strftime("%H")
df_rawdata.insert(20, "Day","")
df_rawdata['Day'] = df_rawdata['Period'].dt.day_name()
df_rawdata.insert(21, "Week","")
df_rawdata['Week'] = df_rawdata['Period'].dt.isocalendar().week
df_rawdata.insert(22, "Year","")
df_rawdata['Year'] = df_rawdata['Period'].dt.isocalendar().year
df_rawdata.insert(23, "Ref_WK","")
df_rawdata['Ref_WK'] = 'Week ' + df_rawdata['Week'].astype(str)
df_rawdata.head()
#df_rawdata.to_excel((in_dir+"\\"+"Ericson_Daily_NW_Report_Input(Hourly)-4G.xlsx"),sheet_name ='NW',index = False)


# In[3]:


df = df_infile.append(df_rawdata, ignore_index=True)
df.head()


# In[4]:


df['Period'] = pd.to_datetime(df['Period'])
df['DATE'] = df['Period'].dt.strftime("%m/%d/%Y")
df['Hour'] = df['Period'].dt.strftime("%H")
df['Ref_HR'] = df['Period'].dt.strftime("%H")
df.head()


# In[5]:


df.shape
Infile = in_dir+"\\"+"Ericsson_4G_Daily_Dashboard_Report.xlsx"
wb1 = openpyxl.load_workbook(os.path.join(in_dir+"\\"+"Ericsson_4G_Daily_Dashboard_Report.xlsx"))
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


# In[6]:


in_dir = cwd
File = win32com.client.Dispatch("Excel.Application")
File.DisplayAlerts = 1
#File.visible = True
xlbook = File.Workbooks.Open(in_dir +"\\"+"Ericsson_4G_Daily_Dashboard_Report.xlsx")
xlbook.RefreshAll()
time.sleep(5)
print("Charts Refreshed")
xlbook.Save()
File.Quit()


# In[ ]:




