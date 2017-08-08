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
    database_clean()
    current_directory = filedialog.askdirectory()
    list_files(current_directory)

    
def list_files(dir):                                                                                                  
    #r = []
    #file_count = len(files)
    for root, dirs, files in os.walk(dir):
        for name in files:
            global glb_path
            glb_path = (os.path.join(root, name))
            print (name)
            if '~$'not in name:
            #if name == 'TEST_004_OPEN.xlsm':
                worksheet_getNEEDBASE()
                worksheet_getAVAILRISk()
                worksheet_getSolBase()
                worksheet_getSolEval()
                
def worksheet_getSolEval():
          df = pd.read_excel(glb_path, sheetname = 'Solution Evaluation')
          #print (df.iloc[45,16])
          insertvalue(df.iloc[0,6], glb_Need_ID ,'EXEC_PC_IC',df.iloc[45,6])
          insertvalue(df.iloc[0,6], glb_Need_ID ,'EXEC_PC_EC',df.iloc[46,6])
          insertvalue(df.iloc[0,6], glb_Need_ID ,'EXEC_PC_MC',df.iloc[47,6])
          
          insertvalue(df.iloc[0,6], glb_Need_ID ,'EXEC_PC_IC_SAP',df.iloc[45,16])
          insertvalue(df.iloc[0,6], glb_Need_ID ,'EXEC_PC_EC_SAP',df.iloc[46,16])
          insertvalue(df.iloc[0,6], glb_Need_ID ,'EXEC_PC_MC_SAP',df.iloc[47,16])
                
           
def worksheet_getSolBase():
    #try:
        df = pd.read_excel(glb_path, sheetname = 'Solution Base')    
        #print (df['Unnamed: 6'], df['Unnamed: 8'])
        #print (df.iloc[362,10])  
        for index, row in df.iterrows():
            #print (index, row[24])
            #MATURATION FIELDS
            
            if (row[24] == 'Løsningens levetid' and index > 182) or (row[24] == 'Lifetime of the Asset' and index > 182) :
                #print (df.iloc[index + 21,9])
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR1_1',df.iloc[index + 8,10])
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR1_2',df.iloc[index + 8,12])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR1_3',df.iloc[index + 8,14])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR1_4',df.iloc[index + 8,16])  
        
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR2_1',df.iloc[index + 8,18])
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR2_2',df.iloc[index + 8,20])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR2_3',df.iloc[index + 8,22])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR2_4',df.iloc[index + 8,24])   
        
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR3_1',df.iloc[index + 8,26])
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR3_2',df.iloc[index + 8,28])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR3_3',df.iloc[index + 8,30])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR3_4',df.iloc[index + 8,32])    
        
        
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR1_1',df.iloc[index + 9,10])
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR1_2',df.iloc[index + 9,12])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR1_3',df.iloc[index + 9,14])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR1_4',df.iloc[index + 9,16])  
        
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR2_1',df.iloc[index + 9,18])
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR2_2',df.iloc[index + 9,20])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR2_3',df.iloc[index + 9,22])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR2_4',df.iloc[index + 9,24])   
        
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR3_1',df.iloc[index + 9,26])
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR3_2',df.iloc[index + 9,28])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR3_3',df.iloc[index + 9,30])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR3_4',df.iloc[index + 9,32])    
        
        
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR1_1',df.iloc[index + 10,10])
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR1_2',df.iloc[index + 10,12])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR1_3',df.iloc[index + 10,14])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR1_4',df.iloc[index + 10,16])  
        
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR2_1',df.iloc[index + 10,18])
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR2_2',df.iloc[index + 10,20])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR2_3',df.iloc[index + 10,22])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR2_4',df.iloc[index + 10,24])   
        
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR3_1',df.iloc[index + 10,26])
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR3_2',df.iloc[index + 10,28])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR3_3',df.iloc[index + 10,30])  
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR3_4',df.iloc[index + 10,32]) 
        
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_DII_ANN',df.iloc[index + 18,9]) 
                insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_IOC_ANN',df.iloc[index + 21,9])                 
                
                x=0
                i=1            
                while x<28:
                     #print ('<-----------------------',str(6+x))
                     #print(i,index,  df.iloc[index + 25,6+x])
#                    print(i, df.iloc[1,6+x])
                     if x < 20:
                         insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_DII_YR' + str(i) ,df.iloc[index + 18,18+x])
                         insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_IOC_YR' + str(i) ,df.iloc[index + 21,18+x])          
                         insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_OD_YR' + str(i) ,df.iloc[index + 25,6+x])
                     else:
                         insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_PEY_YR' + str(i) ,df.iloc[index + 32,6+x])              
#                      
                     x=x+2
                     i=i+1
                    
  

            
            #Lifetime of the Asset
        #insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_BM_IC',df.iloc[23,6])                     #<-- REMOVED 08-08-17
        #insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_BM_EC',df.iloc[24,6])                     #<-- REMOVED 08-08-17
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_SOL_DESC',df.iloc[15,6])                   #<--INTRODUCED 08-08-17         
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_LOA',df.iloc[23,25])                        #<--INTRODUCED 08-08-17 
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_BSU_DESC',df.iloc[29,6])                    #<--INTRODUCED 08-08-17         
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_BSU_PERC',df.iloc[31,6])                    #<--INTRODUCED 08-08-17 
        
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_SOL_DESC',df.iloc[65,6])                   #<--INTRODUCED 08-08-17         
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_LOA',df.iloc[73,25])                        #<--INTRODUCED 08-08-17 
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BSU_DESC',df.iloc[86,6])                    #<--INTRODUCED 08-08-17         
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BSU_PERC',df.iloc[88,6])                    #<--INTRODUCED 08-08-17 
        
        
        #insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BM_IC',df.iloc[73,6]) 
        #insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BM_EC',df.iloc[74,6])        
        
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_BE_IC',df.iloc[23,16])    
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_DE_EC',df.iloc[24,16]) 
        
        
 
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_IC_YR1_1',df.iloc[81,10])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_IC_YR1_2',df.iloc[81,12])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_IC_YR1_3',df.iloc[81,14])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_IC_YR1_4',df.iloc[81,16])  

        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_IC_YR2_1',df.iloc[81,18])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_IC_YR2_2',df.iloc[81,20])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_IC_YR2_3',df.iloc[81,22])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_IC_YR2_4',df.iloc[81,24])   

        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_IC_YR3_1',df.iloc[81,26])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_IC_YR3_2',df.iloc[81,28])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_IC_YR3_3',df.iloc[81,30])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_IC_YR3_4',df.iloc[81,32])    


        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_EC_YR1_1',df.iloc[82,10])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_EC_YR1_2',df.iloc[82,12])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_EC_YR1_3',df.iloc[82,14])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_EC_YR1_4',df.iloc[82,16])  

        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_EC_YR2_1',df.iloc[82,18])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_EC_YR2_2',df.iloc[82,20])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_EC_YR2_3',df.iloc[82,22])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_EC_YR2_4',df.iloc[82,24])   

        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_EC_YR3_1',df.iloc[82,26])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_EC_YR3_2',df.iloc[82,28])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_EC_YR3_3',df.iloc[82,30])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_EC_YR3_4',df.iloc[82,32])    


        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_MAT_YR1_1',df.iloc[83,10])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_MAT_YR1_2',df.iloc[83,12])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_MAT_YR1_3',df.iloc[83,14])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_MAT_YR1_4',df.iloc[83,16])  

        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_MAT_YR2_1',df.iloc[83,18])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_MAT_YR2_2',df.iloc[83,20])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_MAT_YR2_3',df.iloc[83,22])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_MAT_YR2_4',df.iloc[83,24])   

        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_MAT_YR3_1',df.iloc[83,26])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_MAT_YR3_2',df.iloc[83,28])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_MAT_YR3_3',df.iloc[83,30])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_BE_MAT_YR3_4',df.iloc[83,32]) 

###

        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR1_1',df.iloc[362,10])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR1_2',df.iloc[362,12])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR1_3',df.iloc[362,14])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR1_4',df.iloc[362,16])  

        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR2_1',df.iloc[362,18])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR2_2',df.iloc[362,20])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR2_3',df.iloc[362,22])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR2_4',df.iloc[362,24])   

        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR3_1',df.iloc[362,26])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR3_2',df.iloc[362,28])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR3_3',df.iloc[362,30])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_IC_YR3_4',df.iloc[362,32])    


        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR1_1',df.iloc[363,10])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR1_2',df.iloc[363,12])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR1_3',df.iloc[363,14])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR1_4',df.iloc[363,16])  

        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR2_1',df.iloc[363,18])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR2_2',df.iloc[363,20])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR2_3',df.iloc[363,22])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR2_4',df.iloc[363,24])   

        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR3_1',df.iloc[363,26])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR3_2',df.iloc[363,28])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR3_3',df.iloc[363,30])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_EC_YR3_4',df.iloc[363,32])    


        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR1_1',df.iloc[364,10])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR1_2',df.iloc[364,12])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR1_3',df.iloc[364,14])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR1_4',df.iloc[364,16])  

        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR2_1',df.iloc[364,18])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR2_2',df.iloc[364,20])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR2_3',df.iloc[364,22])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR2_4',df.iloc[364,24])   

        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR3_1',df.iloc[364,26])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR3_2',df.iloc[364,28])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR3_3',df.iloc[364,30])  
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_BE_MAT_YR3_4',df.iloc[364,32]) 
        
###        


        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_DII_ANN',df.iloc[372,9]) 
        insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_IOC_ANN',df.iloc[375,9]) 

        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_DII_ANN',df.iloc[91,9]) 
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_IOC_ANN',df.iloc[94,9]) 


             
        
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_BE_MAT',df.iloc[25,16])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_DII_ANN',df.iloc[34,8])
        x=0
        i=1            
        while x<20:
           # print(i, df.iloc[91,18+x])
            #print(i, df.iloc[1,6+x])
            
            insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_DII_YR' + str(i) ,df.iloc[34,18+x])
            insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_IOC_YR' + str(i) ,df.iloc[37,18+x])
            insertvalue(df.iloc[0,6],glb_Need_ID ,'IDEA_OD_YR' + str(i) ,df.iloc[41,6+x])  
            insertvalue(df.iloc[0,6],glb_Need_ID ,'IDEA_PEY_YR' + str(i) ,df.iloc[50,6+x])  
          
            insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_DII_YR' + str(i) ,df.iloc[91,18+x])
            insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_IOC_YR' + str(i) ,df.iloc[94,18+x])          
            insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_OD_YR' + str(i) ,df.iloc[100,6+x])            
            insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_PEY_YR' + str(i) ,df.iloc[107,6+x])        
            
            insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_DII_YR' + str(i) ,df.iloc[372,18+x])
            insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_IOC_YR' + str(i) ,df.iloc[375,18+x])          
            insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_OD_YR' + str(i) ,df.iloc[379,6+x])            
            insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_PEY_YR' + str(i) ,df.iloc[396,6+x])              
            x=x+2
            i=i+1
        while x<28:    
           insertvalue(df.iloc[0,6],glb_Need_ID ,'IDEA_PEY_YR' + str(i) ,df.iloc[50,6+x])
           insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_PEY_YR' + str(i) ,df.iloc[107,6+x])   
           insertvalue(df.iloc[0,6], glb_Need_ID ,'MATU_PEY_YR' + str(i) ,df.iloc[396,6+x])               
           x=x+2
           i=i+1
           
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_BE_MAT',df.iloc[34,8])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_IOC_ANN',df.iloc[37,8])
        insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_OD_DESC',df.iloc[43,6])                        #<--INTRODUCED 08-08-17#
        insertvalue(df.iloc[0,6], glb_Need_ID ,'ANAL_OD_DESC',df.iloc[102,6])                        #<--INTRODUCED 08-08-17        
        #insertvalue(df.iloc[0,6], glb_Need_ID ,'IDEA_STATUS',df.iloc[28,6])                        #<--INTRODUCED 08-08-17     #CHECK FROM XWKBK
    #except: 
     #  print ('WORKBOOK', ' :: ', glb_path, ' :: FAILURE')

def worksheet_getNEEDBASE():
    #path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\01 Client\\Test Workbooks\\TEST_002.xlsm'
    #print(glb_path)
    try:
        df = pd.read_excel(glb_path, sheetname = 'Need Base')
        global glb_Need_ID
        global glb_Need_Title
        glb_Need_ID = df.iloc[7,3]
        glb_Need_Title = df.iloc[14,3]
        #print (Need_ID)
        #print (glb_Need_Title)
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
                        insertvalue(Need_ID, Solution_ID,'SOL_TYPE',Type)
                        insertvalue(Need_ID, Solution_ID,'SOL_YEARS',Years)
                        insertvalue(Need_ID, Solution_ID,'SOL_RAW_PHASE',Phase)
                        insertvalue(Need_ID, Solution_ID,'SOL_RAW_STATUS',Status)
                        offset = offset + 1
                except:
                      None      
#        #print(offset)
#        #worksheet_getAVAILRISk(path, offset, Need_ID)
#        worksheet_getTECHAVAIL(path, offset, Need_ID)
 
    except: 
       print ('WORKBOOK', ' :: ', glb_path, ' :: FAILURE')


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
        
def worksheet_getAVAILRISk():
    try:
        df = pd.read_excel(glb_path, sheetname = 'Availability Risk')
##        #print (df['Unnamed: 7'])
#         for index, row in df.iterrows():
##       #     #print (row[6])
#      #     if row[6] == 'Power' or row[6] == 'El':  
#                insertvalue(Need_ID, 'POWER', df.iloc[index +1,6])
#                insertvalue(Need_ID, 'HEAT', df.iloc[index +1,8])
##       #         #print('power =', df.iloc[index +1,6])
##       #         #print('heat =', df.iloc[index +1,8])
#                 break
#        for index, row in df.iterrows():
#       #     #print (row[28])
#            if row[28] == 'Yes/No' or row[28] == 'Ja/Nej':  
#                insertvalue(Need_ID, 'UNIT_STOP', df.iloc[index +1,28])
#                insertvalue(Need_ID, 'UNIT_STOP_DUR', df.iloc[index +1,31])
##       #         #print('UNIT_STOP =', df.iloc[index +1,28])
##       #         #print('UNIT_STOP_DUR =', df.iloc[index +1,31])
#             break
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
         print ('WORKBOOK', ' :: ', glb_path, ' :: FAILURE')        

def insertvalue(KEY, NEED, PARAM, VALUE):
   #try:    
       cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_INTERVEN (SOLUTION_ID, NEED_ID, PARAM, VALUE, NEED_TITLE) VALUES ('" + str(KEY) +"','" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "','" + str(glb_Need_Title) + "')")
       cursor.commit()
   #except:
   #     print ("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_INTERVEN (SOLUTION_ID, NEED_ID, PARAM, VALUE) VALUES ('" + str(KEY) +"','" + str(NEED) +  "','" + PARAM + "','" + str(VALUE) + "');")

def database_clean():
       cursor.execute("DELETE AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT_INTERVEN")
       cursor.commit()    
    



 

if __name__ == '__main__':
    main()                