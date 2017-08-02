# -*- coding: utf-8 -*-
"""
Created on Wed Jul 26 12:03:25 2017

--> NEEDS ONLY

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

def isNaN(num):
    return num != num

def worksheet_getNEEDBASE(path):
    #path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Test Workbooks\\TEST_002.xlsm'
    try:
        df = pd.read_excel(path, sheetname = 'Need Base')
        Need_ID = df.iloc[7,3]
       # print (Need_ID)
        #print (df.columns)
        #print (df['Unnamed: 3'])
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
        #print(offset)
        #worksheet_getAVAILRISk(path, offset, Need_ID)
        worksheet_getTECHAVAIL(path, offset, Need_ID)
 
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
                
                
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_NOMMAXPWRMAXPWR',df.iloc[index + intloop + 5,6])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_NOMPWRMAXHEAT',df.iloc[index + intloop + 5,11])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_NOMHEAT',df.iloc[index + intloop + 5,14])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_NOMSTEAM',df.iloc[index + intloop + 5,17])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_POB_INCX',df.iloc[index + intloop + 5,34])                
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_UNITHEATLOSSDELTA',df.iloc[index + intloop + 16,6])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_UNITELECCONSUMDELTA',df.iloc[index + intloop + 16,14])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_COAL',df.iloc[index + intloop + 30,6])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_OIL',df.iloc[index + intloop + 30,9])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_NG',df.iloc[index + intloop + 30,12])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_WP',df.iloc[index + intloop + 30,15])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_WC',df.iloc[index + intloop + 30,18])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_S',df.iloc[index + intloop + 30,21])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_W',df.iloc[index + intloop + 30,24])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_SFT_OPEXDELTAPY',df.iloc[index + intloop + 30,34])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_FLEX',df.iloc[index + intloop + 30,6])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_FLEX_DESC',df.iloc[index + intloop + 30,9])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_FLEX_OPEXDELTAPY',df.iloc[index + intloop + 30,31])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_FLEX_INCX',df.iloc[index + intloop + 30,35])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_COAL',df.iloc[index + intloop + 58,6])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_OIL',df.iloc[index + intloop + 58,9])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_NG',df.iloc[index + intloop + 58,12])                
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_WP',df.iloc[index + intloop + 58,18])   
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_WC',df.iloc[index + intloop + 58,21])   
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_S',df.iloc[index + intloop + 58,24])   
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_IFC_W',df.iloc[index + intloop + 58,30])                   
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_BP',df.iloc[index + intloop + 70,6])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_BP_DESC',df.iloc[index + intloop + 70,9])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_BP_INCX',df.iloc[index + intloop + 70,34])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_ETC',df.iloc[index + intloop + 82,6])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_ETC_DESC',df.iloc[index + intloop + 82,9])
                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'OOCII_ETC_INCX',df.iloc[index + intloop + 82,34])
##                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'SOL_RAW_STATUS',Status)
##                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'SOL_RAW_STATUS',Status)
##                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'SOL_RAW_STATUS',Status)
##                    insertvalue(Need_ID, df.iloc[index + intloop + 5,1],'SOL_RAW_STATUS',Status)
                
                
                    #print ( df.iloc[index + intloop + 87,6])
                    intloop = intloop + 1
        
def worksheet_getAVAILRISk(path, offset, Need_ID):
        df = pd.read_excel(path, sheetname = 'Availability Risk')
        #print (df['Unnamed: 7'])
       # for index, row in df.iterrows():
       #     #print (row[6])
       #     if row[6] == 'Power' or row[6] == 'El':  
       #         insertvalue(Need_ID, 'POWER', df.iloc[index +1,6])
       #         insertvalue(Need_ID, 'HEAT', df.iloc[index +1,8])
       #         #print('power =', df.iloc[index +1,6])
       #         #print('heat =', df.iloc[index +1,8])
       #         break
       # for index, row in df.iterrows():
       #     #print (row[28])
       #     if row[28] == 'Yes/No' or row[28] == 'Ja/Nej':  
       #         insertvalue(Need_ID, 'UNIT_STOP', df.iloc[index +1,28])
       #         insertvalue(Need_ID, 'UNIT_STOP_DUR', df.iloc[index +1,31])
       #         #print('UNIT_STOP =', df.iloc[index +1,28])
       #         #print('UNIT_STOP_DUR =', df.iloc[index +1,31])
       #         break
       # for index, row in df.iterrows():
       #     #print (row[5])
       #     if row[5] == 'Yes/No' or row[5] == 'Ja/Nej':  
       #         insertvalue(Need_ID, 'FOSSIL_USE', df.iloc[index +1,6])
       #         insertvalue(Need_ID, 'FOSSIL_USE_DUR_HRS', df.iloc[index +1,10])
       #         insertvalue(Need_ID, 'FUEL_TYPE', df.iloc[index +1,14])
       #         insertvalue(Need_ID, 'FUEL_TYPE_VOL_M3', df.iloc[index +1,18])            
                #print('FOSSIL_USE =', df.iloc[index +1,6])
                #print('FOSSIL_USE_DUR_HRS =', df.iloc[index +1,10])
                #print('FUEL_TYPE =', df.iloc[index +1,14])
                #print('FUEL_TYPE_VOL_M3 =', df.iloc[index +1,18])
        #        break            
#        for index, row in df.iterrows():
#            #print (row[5])
#            if row[5] == 'Expected Frequency of Fine / Penalty per Year' or row[5] == 'Estimeret antal bøder pr. År':  
#                insertvalue(Need_ID, 'FINE', df.iloc[index +2,6])
#                insertvalue(Need_ID, 'FINE_DESC', df.iloc[index +2,10])
#            #    insertvalue(Need_ID, 'FUEL_TYPE', df.iloc[index +1,14])
#            #    insertvalue(Need_ID, 'FUEL_TYPE_VOL_M3', df.iloc[index +1,18])            
#            #   print('FINE =', df.iloc[index +2,6])
#             #  print('FINE_DESC =', df.iloc[index +2,10])
#                #print('FOSSIL_USE_DUR_HRS =', df.iloc[index +1,10])
#                #print('FUEL_TYPE =', df.iloc[index +1,14])
#                #print('FUEL_TYPE_VOL_M3 =', df.iloc[index +1,18])
#         #       break  
#  #  except: 
#   #     print ('WORKBOOK', ' :: ', path, ' :: FAILURE')        

def insertvalue(KEY, NEED, PARAM, VALUE):
#    cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (" + str(KEY) + "," + PARAM +"," + str(VALUE) + ")")
   
   try:    
       cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_INTERVEN (SOLUTION_ID, NEED_ID, PARAM, VALUE) VALUES ('" + str(KEY) +"','" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "')")
       cursor.commit()
   except:
        print ("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_INTERVEN (SOLUTION_ID, NEED_ID, PARAM, VALUE) VALUES ('" + str(KEY) +"','" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "');")
  #cursor.commit()
    
    
def list_files(dir):                                                                                                  
    r = []
    for root, dirs, files in os.walk(dir):
        for name in files:
            r = (os.path.join(root, name))
            print (name)
            if '~$'not in name:
            #if name == 'TEST_001.xlsm':
                worksheet_getNEEDBASE(r)


list_files(path)