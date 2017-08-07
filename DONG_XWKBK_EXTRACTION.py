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
    path = filedialog.askopenfilename()
    database_clean()    
    df = pd.read_excel(path, sheetname="AllData")
    sql_insert = "INSERT AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (%s,%s,%s)" 


def insertvalue(KEY, PARAM, VALUE):
#    cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES (" + str(KEY) + "," + PARAM +"," + str(VALUE) + ")")
   
    try:    
        cursor.execute("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES ('" + str(KEY) + "','" + PARAM +"','" + str(VALUE) + "')")
        cursor.commit()
    except:
         print ("INSERT INTO AROS_WKBK_CONVERSION.dbo.WORKBOOK_EXTRACT (ID, PARAM, VALUE) VALUES ('" + str(KEY) + "','" + PARAM +"','" + str(VALUE) + "');")
  #cursor.commit()
    
def Risk_Pre():
    del df['Name']
    cols = list(df.columns)
    del cols[0]
    for elements in cols:
        if 'Risk_Pre' in elements:
            print (elements)
            df2 = pd.melt(df,id_vars=['Need ID'], value_vars=[elements])
            df3=df2
            #df4 = df3[np.isfinite(df3['Need ID'])]
            df4 = df3.drop_duplicates(subset=None, keep='first', inplace = False)
            #print (df4)
            for index, row in df4.iterrows():
                #print (row[0],row[1],row[2])
                insertvalue(row[0],row[1],row[2])


if __name__ == '__main__':
    main()                