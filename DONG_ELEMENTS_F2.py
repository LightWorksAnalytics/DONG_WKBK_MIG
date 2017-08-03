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
                worksheet_getSolBase()
                
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
#        #insertvalue(Need_ID, Solution_ID,'SOL_RAW_STATUS',Status)                      
#        #print(offset)
# 
#    except: 
#       print ('WORKBOOK', ' :: ', path, ' :: FAILURE')
#
#def worksheet_getAVAILRISk(path, offset, Need_ID):
#        df = pd.read_excel(path, sheetname = 'Availability Risk')
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
#       # for index, row in df.iterrows():
#       #     #print (row[5])
#       #     if row[5] == 'Yes/No' or row[5] == 'Ja/Nej':  
#       #         insertvalue(Need_ID, 'FOSSIL_USE', df.iloc[index +1,6])
#       #         insertvalue(Need_ID, 'FOSSIL_USE_DUR_HRS', df.iloc[index +1,10])
#       #         insertvalue(Need_ID, 'FUEL_TYPE', df.iloc[index +1,14])
#       #         insertvalue(Need_ID, 'FUEL_TYPE_VOL_M3', df.iloc[index +1,18])            
#                #print('FOSSIL_USE =', df.iloc[index +1,6])
#                #print('FOSSIL_USE_DUR_HRS =', df.iloc[index +1,10])
#                #print('FUEL_TYPE =', df.iloc[index +1,14])
#                #print('FUEL_TYPE_VOL_M3 =', df.iloc[index +1,18])
#        #        break            
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
        print ('WORKBOOK', ' :: ', glb_path, ' :: FAILURE')        

def insertvalue( NEED, PARAM, VALUE):
#    cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (" + str(KEY) + "," + PARAM +"," + str(VALUE) + ")")
   
  # try:    
       cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT ( ID, PARAM, VALUE) VALUES ('" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "')")
       cursor.commit()
   #except:
    #    print ("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT ( NEED_ID, PARAM, VALUE) VALUES ('" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "');")


if __name__ == '__main__':
    main()                