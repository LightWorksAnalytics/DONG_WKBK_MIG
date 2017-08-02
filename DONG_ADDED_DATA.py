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
#path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Test Workbooks\\'

def worksheet_getNEEDBASE(path):
    #path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Test Workbooks\\TEST_002.xlsm'
    try:
        df = pd.read_excel(path, sheetname = 'Need Base')
        Need_ID = df.iloc[7,3]
        DML =  df.iloc[7,23]
        AL =  df.iloc[11,15]
        Customer = df.iloc[27,3]
        MDR = df.iloc[30,16]
        HSR = df.iloc[33,16]
        ENR = df.iloc[35,16]
        OLO = df.iloc[37,16]
        NNRR = df.iloc[39,16]
        insertvalue (Need_ID,'DOC_MAN_LOC', DML)
        insertvalue (Need_ID,'ASSETLOC', AL)
        insertvalue (Need_ID,'CUSTOMER', Customer)
        insertvalue (Need_ID,'MUST_DO_REASON', MDR)
        insertvalue (Need_ID,'HS_REASON', HSR)
        insertvalue (Need_ID,'ENV_REASON', ENR)
        insertvalue (Need_ID,'LEGAL_REASON', OLO)
        insertvalue (Need_ID,'NNR_REASON', NNRR)
    except: 
        print ('WORKBOOK', ' :: ', path, ' :: FAILURE')
        
def worksheet_getAVAILRISk(path):
    #path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Test Workbooks\\TEST_002.xlsm'
   # try:
        print('TRUE')
        df = pd.read_excel(path, sheetname = 'Availability Risk')
        print (df.columns)
  #  except: 
   #     print ('WORKBOOK', ' :: ', path, ' :: FAILURE')        

def insertvalue(KEY, PARAM, VALUE):
#    cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (" + str(KEY) + "," + PARAM +"," + str(VALUE) + ")")
   
    try:    
        cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_SUPP (ID, PARAM, VALUE) VALUES ('" + str(KEY) + "','" + PARAM +"','" + str(VALUE) + "')")
        cursor.commit()
    except:
         print ("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_SUPP (ID, PARAM, VALUE) VALUES ('" + str(KEY) + "','" + PARAM +"','" + str(VALUE) + "');")
    
def list_files(dir):                                                                                                  
    r = []
    for root, dirs, files in os.walk(dir):
        for name in files:
            r = (os.path.join(root, name))
            print (name)
            if '~$'not in name:
            #if name == 'TEST_003.xlsm':
                worksheet_getNEEDBASE(r)
                #worksheet_getAVAILRISk(r)

list_files(path)