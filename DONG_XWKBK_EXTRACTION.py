# -*- coding: utf-8 -*-
"""
Created on Wed Jul 26 14:34:24 2017

@author: Alex
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Jul 26 10:21:47 2017

@author: Alex
"""



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

def main():
    path = filedialog.askopenfilename(title='CROSS WORKBOOK OPEN')
    #database_clean()    
    global df

    df = pd.read_excel(path, sheetname="AllData")
    del df['Name']
 
    global cols
    cols = list(df.columns)
    del cols[0]  
    Risk_Pre()
    TimeFull()
    Partial()
    FullCosts()
    HS()
    ENV()
    Rep()


def insertvalue(KEY,NTitle,  PARAM, VALUE):
#    cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (" + str(KEY) + "," + PARAM +"," + str(VALUE) + ")")
   
    try:    
       cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT ( ID, NEED_TITLE, PARAM, VALUE) VALUES ('" + str(KEY) +  "','" + str(NTitle) + "','" + PARAM + "','" + str(VALUE) + "')")
       cursor.commit()
    except:
       print ("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES ('" + str(KEY) + "','" + PARAM +"','" + str(VALUE) + "');")
  #cursor.commit()
    
def Risk_Pre():

    for elements in cols:
        if 'Risk_Pre' in elements:
            #print (elements)
            df2 = pd.melt(df,id_vars=['Need ID', 'Need Title'], value_vars=[elements])
            df3=df2
            #df4 = df3[np.isfinite(df3['Need ID'])]
            df4 = df3.drop_duplicates(subset=None, keep='first', inplace = False)
            #print (df4)
            for index, row in df4.iterrows():
                #print (row[0],row[1],row[2])
                insertvalue(row[0],row[1],row[2], row[3])

def HS():

    for elements in cols:
        if 'H&S_Pre' in elements:
            #print (elements)
            df2 = pd.melt(df,id_vars=['Need ID', 'Need Title'], value_vars=[elements])
            df3=df2
            #df4 = df3[np.isfinite(df3['Need ID'])]
            df4 = df3.drop_duplicates(subset=None, keep='first', inplace = False)
            #print (df4)
            for index, row in df4.iterrows():
                #print (row[0],row[1],row[2])
                insertvalue(row[0],row[1],row[2], row[3])

def ENV():

    for elements in cols:
        if 'Env_Pre' in elements:
            #print (elements)
            df2 = pd.melt(df,id_vars=['Need ID', 'Need Title'], value_vars=[elements])
            df3=df2
            #df4 = df3[np.isfinite(df3['Need ID'])]
            df4 = df3.drop_duplicates(subset=None, keep='first', inplace = False)
            #print (df4)
            for index, row in df4.iterrows():
                #print (row[0],row[1],row[2])
                insertvalue(row[0],row[1],row[2], row[3])  

def Rep():

    for elements in cols:
        if 'Rep_Pre' in elements:
            #print (elements)
            df2 = pd.melt(df,id_vars=['Need ID', 'Need Title'], value_vars=[elements])
            df3=df2
            #df4 = df3[np.isfinite(df3['Need ID'])]
            df4 = df3.drop_duplicates(subset=None, keep='first', inplace = False)
            #print (df4)
            for index, row in df4.iterrows():
                #print (row[0],row[1],row[2])
                insertvalue(row[0],row[1],row[2], row[3])

            

def TimeFull():



    for elements in cols:
        if 'TimeFullRestoration' in elements:
            #print (elements)
            df2 = pd.melt(df,id_vars=['Need ID', 'Need Title'], value_vars=[elements])
            df3=df2
            #df4 = df3[np.isfinite(df3['Need ID'])]
            df4 = df3.drop_duplicates(subset=None, keep='first', inplace = False)
            #print (df4)
            for index, row in df4.iterrows():
                #print (row[0],row[1],row[2])
                insertvalue(row[0],row[1],row[2], row[3])

def Partial():


    for elements in cols:
        if 'AR_PartialRestorationCosts' in elements:
            #print (elements)
            df2 = pd.melt(df,id_vars=['Need ID', 'Need Title'], value_vars=[elements])
            df3=df2
            #df4 = df3[np.isfinite(df3['Need ID'])]
            df4 = df3.drop_duplicates(subset=None, keep='first', inplace = False)
            #print (df4)
            for index, row in df4.iterrows():
                #print (row[0],row[1],row[2])
                insertvalue(row[0],row[1],row[2], row[3])
                
def FullCosts():


    for elements in cols:
        if 'AR_FullRestorationCosts' in elements:
            #print (elements)
            df2 = pd.melt(df,id_vars=['Need ID', 'Need Title'], value_vars=[elements])
            df3=df2
            #df4 = df3[np.isfinite(df3['Need ID'])]
            df4 = df3.drop_duplicates(subset=None, keep='first', inplace = False)
            #print (df4)
            for index, row in df4.iterrows():
                #print (row[0],row[1],row[2])
                insertvalue(row[0],row[1],row[2], row[3])
                
if __name__ == '__main__':
    main()                