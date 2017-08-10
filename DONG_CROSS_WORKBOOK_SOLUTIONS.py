# -*- coding: utf-8 -*-
"""

DONG_CROSS_WORKBOOK_SOLUTIONS

Created on Wed Jul 26 10:21:47 2017

@author: Alex
- coding: utf-8 -*-

"""

import os
import pandas as pd
import pyodbc
import chardet
import csv
from tkinter import Tk
from tkinter import filedialog

server = 'tcp:icsdatabaseanalytics.database.windows.net'
database = 'AROS_WKBK_CONVERSION'
username = 'dr_admin'
password = 'Aslongasibreatheiattack!'
driver= '{SQL Server}'
cnxn = pyodbc.connect('DRIVER='+driver+';PORT=1433;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

root = Tk()
root.withdraw()
path = filedialog.askopenfilename(title='CROSS WORKBOOK OPEN')


df = pd.read_excel(path, sheetname="AllData", skip=15)


def insertvalue(KEY, NEED, TITLE, PARAM, VALUE):
#    cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (" + str(KEY) + "," + PARAM +"," + str(VALUE) + ")")
   
   #try:    
       cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_INTERVEN (SOLUTION_ID, NEED_ID, PARAM, VALUE, NEED_TITLE) VALUES ('" + str(KEY) +"','" + str(NEED) +  "','" + str(PARAM).upper() + "','" + str(VALUE) + "','" + str(TITLE).upper() + "')")
       cursor.commit()
   #except:
   #     print ("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_INTERVEN (SOLUTION_ID, NEED_ID, PARAM, VALUE) VALUES ('" + str(KEY) +"','" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "');")
  #cursor.commit()
    

del df['Name']
cols = list(df.columns)
del cols[0]

for elements in cols:
    if 'H&S_Post_' in elements:
        #print (elements)
        df2 = pd.melt(df,id_vars=['Solution ID','Need ID','Need Title'], value_vars=[elements])
        #print (df2.head(50))
        df3=df2
        #df4 = df3[np.isfinite(df3['Need ID'])]
        #df4 = df3.drop_duplicates(subset=None, keep='first', inplace = False)
        #print (df4)
        for index, row in df3.iterrows():
            #print (row[0], ' + ' ,row[1],' + ',row[2],' + ',row[3])
           #print (row[0],row[1],row[2])
            insertvalue(row[0], row[1],row[2],row[3], row[4])        

for elements in cols:
    if 'Risk_Post_' in elements:
#        print (elements)
        df2 = pd.melt(df,id_vars=['Solution ID','Need ID','Need Title'], value_vars=[elements])
        #print (df2.head(50))
        df3=df2
        #df4 = df3[np.isfinite(df3['Need ID'])]
        #df4 = df3.drop_duplicates(subset=None, keep='first', inplace = False)
        #print (df4)
        for index, row in df3.iterrows():
#            #print (row[0], ' + ' ,row[1],' + ',row[2],' + ',row[3])
            #print (row[0],row[1],row[2])
            insertvalue(row[0], row[1],row[2],row[3], row[4])        

for elements in cols:
    if 'Plant' in elements:
#        print (elements)
        df2 = pd.melt(df,id_vars=['Solution ID','Need ID','Need Title'], value_vars=[elements])
        df3=df2
        #df4 = df3[np.isfinite(df3['Need ID'])]
        df4 = df3.drop_duplicates(subset=['Solution ID'], keep='first', inplace = False)
#        #print (df4)
        for index, row in df4.iterrows():
            #print (row[0],row[1],row[2])
            #print (row[0],row[1],row[2])
            insertvalue(row[0], row[1],row[2],row[3], row[4])            
for elements in cols:
    if 'Unit' in elements:
#        print (elements)
        df2 = pd.melt(df,id_vars=['Solution ID','Need ID','Need Title'], value_vars=[elements])
        df3=df2
        #df4 = df3[np.isfinite(df3['Need ID'])]
        df4 = df3.drop_duplicates(subset=['Solution ID'], keep='first', inplace = False)
#        #print (df4)
        for index, row in df4.iterrows():
#            #print (row[0],row[1],row[2])
            #print (row[0],row[1],row[2])
            insertvalue(row[0], row[1],row[2],row[3], row[4])         
for elements in cols:
    if 'SolDescription' in elements:
#        print (elements)
        df2 = pd.melt(df,id_vars=['Solution ID','Need ID','Need Title'], value_vars=[elements])
        df3=df2
        #df4 = df3[np.isfinite(df3['Need ID'])]
        df4 = df3.drop_duplicates(subset=['Solution ID'], keep='first', inplace = False)
#        #print (df4)
        for index, row in df4.iterrows():
#            #print (row[0],row[1],row[2])
            #print (row[0],row[1],row[2])
            insertvalue(row[0], row[1],row[2],row[3], row[4])    

for elements in cols:
    if 'PotentialExecYear' in elements:
#        print (elements)
        df2 = pd.melt(df,id_vars=['Solution ID','Need ID','Need Title'], value_vars=[elements])
        df3=df2
#        #df4 = df3[np.isfinite(df3['Need ID'])]
        df4 = df3.drop_duplicates(subset=['Solution ID'], keep='first', inplace = False)
#        #print (df4)
        for index, row in df4.iterrows():
#            #print (row[0],row[1],row[2])
            #print (row[0],row[1],row[2])
            insertvalue(row[0], row[1],row[2],row[3], row[4]) 

for elements in cols:
    if 'Env_Post' in elements:
#        print (elements)
        df2 = pd.melt(df,id_vars=['Solution ID','Need ID','Need Title'], value_vars=[elements])
        df3=df2
        #df4 = df3[np.isfinite(df3['Need ID'])]
        df4 = df3.drop_duplicates(subset=['Solution ID'], keep='first', inplace = False)
#        #print (df4)
        for index, row in df4.iterrows():
#            #print (row[0],row[1],row[2])
            #print (row[0],row[1],row[2])
            insertvalue(row[0], row[1],row[2],row[3], row[4]) 

for elements in cols:                                                                                       
    if 'Rep_Post' in elements:
        print (elements)
        df2 = pd.melt(df,id_vars=['Solution ID','Need ID','Need Title'], value_vars=[elements])
        df3=df2
        #df4 = df3[np.isfinite(df3['Need ID'])]
        df4 = df3.drop_duplicates(subset=['Solution ID'], keep='first', inplace = False)
        #print (df4)
        for index, row in df4.iterrows():
            print (row[0],row[1],row[2])
            insertvalue(row[0], row[1],row[2],row[3], row[4])  

            

#df.iloc[7,3]