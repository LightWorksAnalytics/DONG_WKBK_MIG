# -*- coding: utf-8 -*-
"""
Created on Mon Jul 24 15:24:05 2017

@author: Alex
"""

import pandas as pd
import os
import pyodbc


server = 'tcp:icsdatabaseanalytics.database.windows.net'
database = 'AROS_WKBK_CONVERSION'
username = 'dr_admin'
password = 'Aslongasibreatheiattack!'
driver= '{SQL Server}'
cnxn = pyodbc.connect('DRIVER='+driver+';PORT=1433;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()

path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Workbook Sharepoint Extracts\\170714\\ASV PINQ Light bibliotek\\'


def worksheet_getNEEDBASE(path, name):
    #path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Test Workbooks\\TEST_002.xlsm'
    try:
        df = pd.read_excel(path, sheetname = 'Need Base')
        Need_ID = df.iloc[7,3]
        insertvalue(Need_ID, path, name)

    except: 
        print ('WORKBOOK', ' :: ', path, ' :: FAILURE')
#
def insertvalue(KEY, path, Name):
#    cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (" + str(KEY) + "," + PARAM +"," + str(VALUE) + ")")
   
   try: 
   #     print("UPLOADING")
    cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_DETAILS (ID, PATH, NAME) VALUES ('" + str(KEY) + "','" + str(path) +"','" + str(Name) + "')")
    cursor.commit()
   except:
    print ("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_DETAILS (ID, PATH, NAME) VALUES ('" + str(KEY) + "','" + str(path) +"','" + str(NAME) + "')")
    
def list_files(dir):                                                                                                  
    r = []
    for root, dirs, files in os.walk(dir):
        for name in files:
            r = (os.path.join(root, name))
            if '~$'not in name:
                worksheet_getNEEDBASE(r, name)

list_files(path)