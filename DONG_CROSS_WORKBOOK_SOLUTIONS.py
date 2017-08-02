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


server = 'tcp:icsdatabaseanalytics.database.windows.net'
database = 'AROS_WKBK_CONVERSION'
username = 'dr_admin'
password = 'Aslongasibreatheiattack!'
driver= '{SQL Server}'
cnxn = pyodbc.connect('DRIVER='+driver+';PORT=1433;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()



df = pd.read_excel('O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Cross WorkBook\\20170227 Cross Workbook Reporting _ 64 V3.xlsm', sheetname="AllData", skip=15)

sql_insert = "INSERT AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (%s,%s,%s)" 


def insertvalue(KEY, NEED, PARAM, VALUE):
#    cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (" + str(KEY) + "," + PARAM +"," + str(VALUE) + ")")
   
   try:    
       cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_INTERVEN (SOLUTION_ID, NEED_ID, PARAM, VALUE) VALUES ('" + str(KEY) +"','" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "')")
       cursor.commit()
   except:
        print ("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_INTERVEN (SOLUTION_ID, NEED_ID, PARAM, VALUE) VALUES ('" + str(KEY) +"','" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "');")
  #cursor.commit()
    

del df['Name']
cols = list(df.columns)
del cols[0]
#for elements in cols:
#    if 'H&S_Post_' in elements:
#        print (elements)
#        df2 = pd.melt(df,id_vars=['Solution ID','Need ID'], value_vars=[elements])
#        #print (df2.head(50))
#        df3=df2
#        #df4 = df3[np.isfinite(df3['Need ID'])]
#        #df4 = df3.drop_duplicates(subset=None, keep='first', inplace = False)
#        #print (df4)
#        for index, row in df3.iterrows():
#            #print (row[0], ' + ' ,row[1],' + ',row[2],' + ',row[3])
#            insertvalue(row[0], row[1],row[2],row[3])

#for elements in cols:
#    if 'Risk_Post_' in elements:
#        print (elements)
#        df2 = pd.melt(df,id_vars=['Solution ID','Need ID'], value_vars=[elements])
#        #print (df2.head(50))
#        df3=df2
#        #df4 = df3[np.isfinite(df3['Need ID'])]
#        #df4 = df3.drop_duplicates(subset=None, keep='first', inplace = False)
#        #print (df4)
#        for index, row in df3.iterrows():
#            #print (row[0], ' + ' ,row[1],' + ',row[2],' + ',row[3])
#            insertvalue(row[0], row[1],row[2],row[3])

#for elements in cols:
#    if 'Plant' in elements:
#        print (elements)
#        df2 = pd.melt(df,id_vars=['Solution ID','Need ID'], value_vars=[elements])
#        df3=df2
#        #df4 = df3[np.isfinite(df3['Need ID'])]
#        df4 = df3.drop_duplicates(subset=['Solution ID'], keep='first', inplace = False)
#        #print (df4)
#        for index, row in df4.iterrows():
#            #print (row[0],row[1],row[2])
#            insertvalue(row[0], row[1],row[2],row[3])       
#for elements in cols:
#    if 'Unit' in elements:
#        print (elements)
#        df2 = pd.melt(df,id_vars=['Solution ID','Need ID'], value_vars=[elements])
#        df3=df2
#        #df4 = df3[np.isfinite(df3['Need ID'])]
#        df4 = df3.drop_duplicates(subset=['Solution ID'], keep='first', inplace = False)
#        #print (df4)
#        for index, row in df4.iterrows():
#            #print (row[0],row[1],row[2])
#            insertvalue(row[0], row[1],row[2],row[3])   
#for elements in cols:
#    if 'SolDescription' in elements:
#        print (elements)
#        df2 = pd.melt(df,id_vars=['Solution ID','Need ID'], value_vars=[elements])
#        df3=df2
#        #df4 = df3[np.isfinite(df3['Need ID'])]
#        df4 = df3.drop_duplicates(subset=['Solution ID'], keep='first', inplace = False)
#        #print (df4)
#        for index, row in df4.iterrows():
#            #print (row[0],row[1],row[2])
#            insertvalue(row[0], row[1],row[2],row[3])  
#for elements in cols:
#    if 'PotentialExecYear' in elements:
#        print (elements)
#        df2 = pd.melt(df,id_vars=['Solution ID','Need ID'], value_vars=[elements])
#        df3=df2
#        #df4 = df3[np.isfinite(df3['Need ID'])]
#        df4 = df3.drop_duplicates(subset=['Solution ID'], keep='first', inplace = False)
#        #print (df4)
#        for index, row in df4.iterrows():
#            #print (row[0],row[1],row[2])
#            insertvalue(row[0], row[1],row[2],row[3])  
#for elements in cols:
#    if 'Env_Post' in elements:
#        print (elements)
#        df2 = pd.melt(df,id_vars=['Solution ID','Need ID'], value_vars=[elements])
#        df3=df2
#        #df4 = df3[np.isfinite(df3['Need ID'])]
#        df4 = df3.drop_duplicates(subset=['Solution ID'], keep='first', inplace = False)
#        #print (df4)
#        for index, row in df4.iterrows():
#            #print (row[0],row[1],row[2])
#            insertvalue(row[0], row[1],row[2],row[3])  
for elements in cols:
    if 'Rep_Post' in elements:
        print (elements)
        df2 = pd.melt(df,id_vars=['Solution ID','Need ID'], value_vars=[elements])
        df3=df2
        #df4 = df3[np.isfinite(df3['Need ID'])]
        df4 = df3.drop_duplicates(subset=['Solution ID'], keep='first', inplace = False)
        #print (df4)
        for index, row in df4.iterrows():
            #print (row[0],row[1],row[2])
            insertvalue(row[0], row[1],row[2],row[3])              

#df.iloc[7,3]