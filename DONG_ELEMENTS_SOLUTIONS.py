# -*- coding: utf-8 -*-
"""
Created on Wed Jul 26 12:03:25 2017



@author: Alex
@Client: DONG Energy
@Purpose: Extract dedicated fields from the DONG Energy workbooks.
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

#path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Workbook Sharepoint Extracts\\170714\\ASV PINQ Light bibliotek\\'
#path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Test Workbooks\\'

root = Tk()
root.withdraw()

glb_Need_ID = None
glb_offset = None
glb_path = None

def main():
    current_directory = filedialog.askdirectory()
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
            #if name == 'TEST_003_OPEN.xlsm':
                worksheet_getNEEDBASE()
                #worksheet_getAVAILRISk()
                worksheet_getSolBase()
           
def worksheet_getSolBase():
    #try:
        df = pd.read_excel(glb_path, sheetname = 'Solution Base')    
        #print (df['Unnamed: 6'])
        #print (df.iloc[23,6])
        #print (df.iloc[0,6])        
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_BM_IC',df.iloc[23,6])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_BM_EC',df.iloc[24,6])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_BE_IC',df.iloc[23,16])    
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_DE_EC',df.iloc[24,16])    
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_BE_MAT',df.iloc[25,16])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_DII_ANN',df.iloc[34,8])
        x=0
        i=1            
        while x<20:
            #print(i, df.iloc[34,18+x])
            insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_DII_YR' + str(i) ,df.iloc[34,18+x])
            insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_DII_YR' + str(i) ,df.iloc[37,18+x])
            insertvalue(df.iloc[0,6],glb_Need_ID ,'IDEA_PEY_YR' + str(i) ,df.iloc[41,6+x])             
            x=x+2
            i=i+1
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_BE_MAT',df.iloc[34,8])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_IOC_ANN',df.iloc[37,8])  
    #except: 
     #  print ('WORKBOOK', ' :: ', glb_path, ' :: FAILURE')

def worksheet_getNEEDBASE():
    #path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Test Workbooks\\TEST_002.xlsm'
    #print(glb_path)
    try:
        df = pd.read_excel(glb_path, sheetname = 'Need Base')
        global glb_Need_ID
        glb_Need_ID = df.iloc[7,3]
        #print (Need_ID)
        #print (df.columns)
        #print (df['Unnamed: 3'])
        global offset
        offset = 0
        for index, row in df.iterrows():
           # print (index, row[3])
            if index >= 45:
                try:
                    if len(row[3])==10:
                        Solution_ID = df.iloc[index,3]
                        Type = df.iloc[index,24]
                        Years = df.iloc[index,31]
                        Phase = df.iloc[index,33]
                        Status = df.iloc[index,36]
                        #print(Solution_ID, ' :: ' , Type, ' :: ' , Years, ' :: ' , Phase, ' :: ', Status)
                        #insertvalue(Need_ID, Solution_ID,'SOL_TYPE',Type)
                        #insertvalue(Need_ID, Solution_ID,'SOL_YEARS',Years)
                        #insertvalue(Need_ID, Solution_ID,'SOL_RAW_PHASE',Phase)
                        #insertvalue(Need_ID, Solution_ID,'SOL_RAW_STATUS',Status)
                        offset = offset + 1
                except:
                      None      
#        #print(offset)
#        #worksheet_getAVAILRISk(path, offset, Need_ID)
#        worksheet_getTECHAVAIL(path, offset, Need_ID)
 
    except: 
       print ('WORKBOOK', ' :: ', path, ' :: FAILURE')


        #print (df.head(20))
def worksheet_getTECHAVAIL(path, offset, Need_ID):
    #path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Test Workbooks\\TEST_002.xlsm'
    #try:
        
        df = pd.read_excel(path, sheetname = 'Other Technical')
        for index, row in df.iterrows():
            #print (index, row[2])
            if row[2] == 'Plant Output Balance' or row[2] == 'Ændringer i blokkens kapacitet':  
                intloop = 1
                #print (index)
                while intloop <= offset:
                    #print (str(index + intloop + 5))
                    #print (df.iloc[index + intloop + 5,1], ' :: ' , Need_ID , ' :: ' , df.iloc[index + intloop + 5,6])
                    #print (df.iloc[index + intloop + 5,1], ' :: ' , Need_ID , ' :: ' , df.iloc[index + intloop + 5,11])
                    #print (df.iloc[index + intloop + 5,1], ' :: ' , Need_ID , ' :: ' , df.iloc[index + intloop + 5,14])
                    #print (df.iloc[index + intloop + 5,1], ' :: ' , Need_ID , ' :: ' , df.iloc[index + intloop + 5,17])
                    #print (df.iloc[index + intloop + 5,1], ' :: ' ,Need_ID , ' :: ' , df.iloc[index + intloop + 16,6])
                    #print (df.iloc[index + intloop + 5,1], ' :: ' ,Need_ID , ' :: ' , df.iloc[index + intloop + 16,14])
                    #print (df.iloc[index + intloop + 5,1], ' :: ' ,Need_ID , ' :: ' , df.iloc[index + intloop + 30,6])
                    #print (df.iloc[index + intloop + 5,1], ' :: ' ,Need_ID , ' :: ' , df.iloc[index + intloop + 30,34])                   
                    #print (df.iloc[index + intloop + 5,1], ' :: ' ,Need_ID , ' :: ' , df.iloc[index + intloop + 42,6])
                    #print (df.iloc[index + intloop + 5,1], ' :: ' ,Need_ID , ' :: ' , df.iloc[index + intloop + 42,9])
                    #print (df.iloc[index + intloop + 5,1], ' :: ' ,Need_ID , ' :: ' , df.iloc[index + intloop + 42,31])
                    #print (df.iloc[index + intloop + 5,1], ' :: ' ,Need_ID , ' :: ' , df.iloc[index + intloop + 42,35])
                
                    #print (df.iloc[index + intloop + 5,1], ' :: ' ,Need_ID , ' :: ' , df.iloc[index + intloop + 58,6])
                    #print (df.iloc[index + intloop + 5,1], ' :: ' ,Need_ID , ' :: ' , df.iloc[index + intloop + 70,34])
                    #print (df.iloc[index + intloop + 5,1], ' :: ' ,Need_ID , ' :: ' , df.iloc[index + intloop + 82,6])
              
                
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_NOMMAXPWRMAXPWR',df.iloc[index + intloop + 5,6])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_NOMPWRMAXHEAT',df.iloc[index + intloop + 5,11])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_NOMHEAT',df.iloc[index + intloop + 5,14])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_NOMSTEAM',df.iloc[index + intloop + 5,17])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_POB_INCX',df.iloc[index + intloop + 5,34])                
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_UNITHEATLOSSDELTA',df.iloc[index + intloop + 16,6])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_UNITELECCONSUMDELTA',df.iloc[index + intloop + 16,14])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_COAL',df.iloc[index + intloop + 30,6])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_OIL',df.iloc[index + intloop + 30,9])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_NG',df.iloc[index + intloop + 30,12])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_WP',df.iloc[index + intloop + 30,15])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_WC',df.iloc[index + intloop + 30,18])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_S',df.iloc[index + intloop + 30,21])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_W',df.iloc[index + intloop + 30,24])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_OPEXDELTAPY',df.iloc[index + intloop + 30,34])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_FLEX',df.iloc[index + intloop + 30,6])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_FLEX_DESC',df.iloc[index + intloop + 30,9])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_FLEX_OPEXDELTAPY',df.iloc[index + intloop + 30,31])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_FLEX_INCX',df.iloc[index + intloop + 30,35])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_COAL',df.iloc[index + intloop + 58,6])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_OIL',df.iloc[index + intloop + 58,9])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_NG',df.iloc[index + intloop + 58,12])                
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_WP',df.iloc[index + intloop + 58,18])   
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_WC',df.iloc[index + intloop + 58,21])   
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_S',df.iloc[index + intloop + 58,24])   
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_W',df.iloc[index + intloop + 58,30])                   
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_BP',df.iloc[index + intloop + 70,6])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_BP_DESC',df.iloc[index + intloop + 70,9])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_BP_INCX',df.iloc[index + intloop + 70,34])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_ETC',df.iloc[index + intloop + 82,6])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_ETC_DESC',df.iloc[index + intloop + 82,9])
#                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_ETC_INCX',df.iloc[index + intloop + 82,34])
##                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'SOL_RAW_STATUS',Status)
##                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'SOL_RAW_STATUS',Status)
##                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'SOL_RAW_STATUS',Status)
##                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'SOL_RAW_STATUS',Status)
                
                
                    #print ( df.iloc[index + intloop + 87,6])
                    intloop = intloop + 1
        
def worksheet_getAVAILRISk():
    try:
        df = pd.read_excel(glb_path, sheetname = 'Availability Risk')
#        #print (df['Unnamed: 7'])
#       # for index, row in df.iterrows():
#       #     #print (row[6])
#       #     if row[6] == 'Power' or row[6] == 'El':  
#       #         insertvalue(Need_ID, 'POWER', df.iloc[index +1,6])
#       #         insertvalue(Need_ID, 'HEAT', df.iloc[index +1,8])
#       #         #print('power =', df.iloc[index +1,6])
#       #         #print('heat =', df.iloc[index +1,8])
#       #         break
#       # for index, row in df.iterrows():
#       #     #print (row[28])
#       #     if row[28] == 'Yes/No' or row[28] == 'Ja/Nej':  
#       #         insertvalue(Need_ID, 'UNIT_STOP', df.iloc[index +1,28])
#       #         insertvalue(Need_ID, 'UNIT_STOP_DUR', df.iloc[index +1,31])
#       #         #print('UNIT_STOP =', df.iloc[index +1,28])
#       #         #print('UNIT_STOP_DUR =', df.iloc[index +1,31])
#       #         break
###      2017-08-02 1011  :: Updated to include Solution based records

        for index, row in df.iterrows():
            #print (index, ' :: ' , row[5], ' :: ', row[6])           
            if row[5] == 'Yes/No' or row[5] == 'Ja/Nej':
                intloop = 1
#               # print (index)
                while intloop <= offset: 
#                     print ('-------> ', index, index + intloop +2)
                      insertvalue(df.iloc[index + intloop +2,1], Need_ID, 'FOSSIL_USE', df.iloc[index + intloop +2,6])
                      insertvalue(df.iloc[index + intloop +2,1], Need_ID, 'FOSSIL_USE_DUR_HRS', df.iloc[index + intloop +2,10])
                      insertvalue(df.iloc[index + intloop +2,1], Need_ID, 'FUEL_TYPE', df.iloc[index + intloop +2,14])
                      insertvalue(df.iloc[index + intloop +2,1], Need_ID, 'FUEL_TYPE_VOL_M3', df.iloc[index + intloop +2,18])     
#                     print (df.iloc[index + intloop +2,1])
#                     print('FOSSIL_USE =', df.iloc[index + intloop +2,6])
#                     print('FOSSIL_USE_DUR_HRS =', df.iloc[index + intloop +2,10])
#                     print('FUEL_TYPE =', df.iloc[index + intloop +2,14])
#                     print('FUEL_TYPE_VOL_M3 =', df.iloc[index + intloop +2,18])
                      break            
##        for index, row in df.iterrows():
##            #print (row[5])
##            if row[5] == 'Expected Frequency of Fine / Penalty per Year' or row[5] == 'Estimeret antal bøder pr. År':  
##                insertvalue(Need_ID, 'FINE', df.iloc[index +2,6])
##                insertvalue(Need_ID, 'FINE_DESC', df.iloc[index +2,10])
##            #    insertvalue(Need_ID, 'FUEL_TYPE', df.iloc[index +1,14])
##            #    insertvalue(Need_ID, 'FUEL_TYPE_VOL_M3', df.iloc[index +1,18])            
##            #   print('FINE =', df.iloc[index +2,6])
##             #  print('FINE_DESC =', df.iloc[index +2,10])
##                #print('FOSSIL_USE_DUR_HRS =', df.iloc[index +1,10])
##                #print('FUEL_TYPE =', df.iloc[index +1,14])
##                #print('FUEL_TYPE_VOL_M3 =', df.iloc[index +1,18])
##         #       break  
    except: 
         print ('WORKBOOK', ' :: ', path, ' :: FAILURE')        

def insertvalue(KEY, NEED, PARAM, VALUE):
   try:    
       cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_INTERVEN (SOLUTION_ID, NEED_ID, PARAM, VALUE) VALUES ('" + str(KEY) +"','" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "')")
       cursor.commit()
   except:
        print ("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_INTERVEN (SOLUTION_ID, NEED_ID, PARAM, VALUE) VALUES ('" + str(KEY) +"','" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "');")

    
    



 

if __name__ == '__main__':
    main()                