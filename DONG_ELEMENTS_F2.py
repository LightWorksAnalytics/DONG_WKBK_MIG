# -*- coding: utf-8 -*-
"""
Created on Wed Aug  2 13:18:48 2017

@author: Alex
"""

import pandas as pd
import os
import pyodbc
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

glb_Need_ID = None
glb_offset = None
glb_path = None

def main():
    current_directory = filedialog.askdirectory()
    database_clean()
    list_files(current_directory)

    
def list_files(dir):                                                                                                  
    #r = []
    #file_count = len(files)
    for root, dirs, files in os.walk(dir):
        for name in files:
            global glb_path
            glb_path = (os.path.join(root, name))
            #print (name)
            if '~$'not in name:
            #if name == 'TEST_004_OPEN.xlsm':
                worksheet_getNEEDBASE()
                #worksheet_getAVAILRISk()
                #worksheet_getSolBase()
                
#def worksheet_getSolBase():
#    #try:
#        df = pd.read_excel(glb_path, sheetname = 'Solution Base')    
#        #print (df['Unnamed: 8'])
#        #print (df.iloc[23,6])
#        print (df.iloc[34,8])        
#        insertvalue(glb_Need_ID ,'IDEA_BM_IC',df.iloc[23,6])
#        insertvalue(glb_Need_ID ,'IDEA_BM_EC',df.iloc[24,6])
#        insertvalue(glb_Need_ID ,'IDEA_BE_IC',df.iloc[23,16])    
#        insertvalue(glb_Need_ID ,'IDEA_DE_EC',df.iloc[24,16])    
#        insertvalue(glb_Need_ID ,'IDEA_BE_MAT',df.iloc[25,16])
#        insertvalue(glb_Need_ID ,'IDEA_DII_ANN',df.iloc[34,8])
#        x=0
#        i=1            
#        while x<20:
#            #print(i, df.iloc[34,18+x])
#            insertvalue(glb_Need_ID ,'IDEA_DII_YR' + str(i) ,df.iloc[34,18+x])
#            insertvalue(glb_Need_ID ,'IDEA_DII_YR' + str(i) ,df.iloc[37,18+x])
#            insertvalue(glb_Need_ID ,'IDEA_PEY_YR' + str(i) ,df.iloc[41,6+x])             
#            x=x+2
#            i=i+1
#        insertvalue(glb_Need_ID ,'IDEA_BE_MAT',df.iloc[34,8])
#        insertvalue(glb_Need_ID ,'IDEA_IOC_ANN',df.iloc[37,8])  
#    #except: 
#     #  print ('WORKBOOK', ' :: ', glb_path, ' :: FAILURE')


def worksheet_getNEEDBASE():
#    #path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Test Workbooks\\TEST_002.xlsm'
    try:
        df = pd.read_excel(glb_path, sheetname = 'Need Base')
        global glb_Need_ID
        glb_Need_ID = df.iloc[7,3]
#        print('CRITICALASSET',df.iloc[11,23])
#        print('NEED TITLE',df.iloc[14,3])
#        print('PM',df.iloc[17,3])
#        print('NEED DESCRIPTION',df.iloc[20,3])
#        print('MDR',df.iloc[30,16])
#        print('MUSTDOBY',df.iloc[31,9])
#        print('NNR',df.iloc[39,42])  
        insertvalue(glb_Need_ID, 'PLANT', df.iloc[11,4])
        insertvalue(glb_Need_ID, 'UNIT', df.iloc[11,4])        
        insertvalue(glb_Need_ID, 'ASSETLOC', df.iloc[11,4])
        insertvalue(glb_Need_ID, 'NEED_TITLE', df.iloc[14,3])
        insertvalue(glb_Need_ID, 'PROJECTMANAGER', df.iloc[17,3])
        insertvalue(glb_Need_ID, 'NEEDDESCRIPTION', df.iloc[20,3])
        insertvalue(glb_Need_ID, 'MUSTDO', df.iloc[30,9])
        insertvalue(glb_Need_ID, 'CRITICALASSET', df.iloc[11,23])
        insertvalue(glb_Need_ID, 'MUST_DO_REASON', df.iloc[30,16])        
        insertvalue(glb_Need_ID, 'MUSTDOBY', df.iloc[31,9])     
        insertvalue(glb_Need_ID, 'HEALTHANDSAFETYCOMPLIANCE', df.iloc[33,9])
        insertvalue(glb_Need_ID, 'HS_REASON', df.iloc[33,16])
        insertvalue(glb_Need_ID, 'ENVIRONMENTALCOMPLIANCE', df.iloc[35,9])
        insertvalue(glb_Need_ID, 'ENV_REASON', df.iloc[35,16])
        insertvalue(glb_Need_ID, 'OTHERLEGALOBLIGATION', df.iloc[37,9])         
        insertvalue(glb_Need_ID, 'LEGAL_REASON', df.iloc[37,16])
        insertvalue(glb_Need_ID, 'NEEDNOTREQUIRED', df.iloc[39,42])
        insertvalue(glb_Need_ID, 'NNR_REASON', df.iloc[39,16])
## 
    except: 
       print ('WORKBOOK', ' :: ', glb_path, ' :: FAILURE')
#
def worksheet_getAVAILRISk():
        df = pd.read_excel(glb_path, sheetname = 'Availability Risk')
        #print (df['Unnamed: 7'])
        for index, row in df.iterrows():
            #print (row[6])
            if row[6] == 'Power' or row[6] == 'El':  
                insertvalue(glb_Need_ID, 'POWER', df.iloc[index +1,6])
                insertvalue(glb_Need_ID, 'HEAT', df.iloc[index +1,8])
#                print('power =', df.iloc[index +1,6])
#                print('heat =', df.iloc[index +1,8])
                break
        for index, row in df.iterrows():
            #print (row[28])
            if row[28] == 'Yes/No' or row[28] == 'Ja/Nej':  
                insertvalue(glb_Need_ID, 'UNIT_STOP', df.iloc[index +1,28])
                insertvalue(glb_Need_ID, 'UNIT_STOP_DUR', df.iloc[index +1,31])
#                print('UNIT_STOP =', df.iloc[index +1,28])
#                print('UNIT_STOP_DUR =', df.iloc[index +1,31])
                break
        for index, row in df.iterrows():
            #print (row[5])
            if row[5] == 'Yes/No' or row[5] == 'Ja/Nej':  
                insertvalue(glb_Need_ID, 'FOSSIL_USE', df.iloc[index +1,6])
                insertvalue(glb_Need_ID, 'FOSSIL_USE_DUR_HRS', df.iloc[index +1,10])
                insertvalue(glb_Need_ID, 'FUEL_TYPE', df.iloc[index +1,14])
                insertvalue(glb_Need_ID, 'FUEL_TYPE_VOL_M3', df.iloc[index +1,18])            
#                print('FOSSIL_USE =', df.iloc[index +1,6])
#                print('FOSSIL_USE_DUR_HRS =', df.iloc[index +1,10])
#                print('FUEL_TYPE =', df.iloc[index +1,14])
#                print('FUEL_TYPE_VOL_M3 =', df.iloc[index +1,18])
                break            
        for index, row in df.iterrows():
            #print (row[5])
            if row[5] == 'Expected Frequency of Fine / Penalty per Year' or row[5] == 'Estimeret antal bøder pr. År':  
                insertvalue(glb_Need_ID, 'FINE', df.iloc[index +2,6])
                insertvalue(glb_Need_ID, 'FINE_DESC', df.iloc[index +2,10])
                insertvalue(glb_Need_ID, 'FUEL_TYPE', df.iloc[index +1,14])
                insertvalue(glb_Need_ID, 'FUEL_TYPE_VOL_M3', df.iloc[index +1,18])            
#               print('FINE =', df.iloc[index +2,6])
#               print('FINE_DESC =', df.iloc[index +2,10])
#               print('FINE_P_FINE =', df.iloc[index +2,34])
                break  
    

def insertvalue( NEED, PARAM, VALUE):
#    cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (" + str(KEY) + "," + PARAM +"," + str(VALUE) + ")")
   
  # try:    
       cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT ( ID, PARAM, VALUE) VALUES ('" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "')")
       cursor.commit()
   #except:
    #    print ("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT ( NEED_ID, PARAM, VALUE) VALUES ('" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "');")

def database_clean():
       cursor.execute("DELETE AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT")
       cursor.commit()


if __name__ == '__main__':
    main()                