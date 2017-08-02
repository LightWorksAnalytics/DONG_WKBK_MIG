# -*- coding: utf-8 -*-
"""
Created on Fri Jul 14 13:53:29 2017

@author: Alex
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



df = pd.read_excel('O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Cross WorkBook\\20170227 Cross Workbook Reporting _ 64 V3.xlsm', sheetname="AllData")

sql_insert = "INSERT AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (%s,%s,%s)" 


def insertvalue(KEY, PARAM, VALUE):
#    cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (" + str(KEY) + "," + PARAM +"," + str(VALUE) + ")")
   
    try:    
        cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_SUPP (ID, PARAM, VALUE) VALUES ('" + str(KEY) + "','" + PARAM +"','" + str(VALUE) + "')")
        cursor.commit()
    except:
         print ("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_SUPP (ID, PARAM, VALUE) VALUES ('" + str(KEY) + "','" + PARAM +"','" + str(VALUE) + "');")
  #cursor.commit()
    
cols = list(df.columns)
for elements in cols:
    if 'Pre' in elements:
        print (elements)
        df2 = pd.melt(df,id_vars=['Need ID'], value_vars=[elements])
        df3=df2
        df4 = df3[np.isfinite(df3['Need ID'])]
        #df4 = df4.drop_duplicates(subset=None, keep='first', inplace = False)
        #print (df4)
        #for index, row in df4.iterrows():
            #print (row[2].strip('F'))
           # insertvalue(row[0],row[1],row[2])


#df.iloc[7,3]