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

df = pd.read_excel('O:\\Clients\\DONG\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Workbook Sharepoint Extracts\\170714\ESV\\Absorber - Udskiftning af styring til elevator_32_64.xlsm', sheetname=1)

sql_insert = "INSERT AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES ('" + %d + "','" +  %d +"','" + %d +"')") 

def insertvalue(KEY, PARAM, VALUE):
    





df = pd.read_excel('O:\\Clients\\DONG\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Workbook Sharepoint Extracts\\170714\ESV\\Absorber - Udskiftning af styring til elevator_32_64.xlsm', sheetname=1)





df.iloc[7,3]