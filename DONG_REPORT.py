# -*- coding: utf-8 -*-

import plotly.plotly as py
from plotly.graph_objs import *
import pyodbc
import pandas

server = 'tcp:icsdatabaseanalytics.database.windows.net'
database = 'AROS_WKBK_CONVERSION'
username = 'dr_admin'
password = 'Aslongasibreatheiattack!'
driver= '{SQL Server}'
cnxn = pyodbc.connect('DRIVER='+driver+';PORT=1433;SERVER='+server+';PORT=1443;DATABASE='+database+';UID='+username+';PWD='+ password)
cursor = cnxn.cursor()


sql_plnt_rec = "Select [VALUE], count(distinct([ID])) from [dbo].[WORKBOOK_EXTRACT] where [PARAM] = 'PLANT' group by [VALUE]"

sql_plnt_wkbks = "SELECT [FOLDER_SITE], count(distinct(ID)) rec_count FROM (

SELECT [ID], [PATH],[NEED_TITLE], LEFT(REPLACE([PATH], 'O:\Clients\DONG\DONG 02 - Asset Risk and Optimisation Suite\02 Data\01 Input\01 Client\Workbook Sharepoint Extracts\170714\',''),3) FOLDER_SITE
 FROM (
  	SELECT [ID]
		  ,[PATH]
		  ,[NAME]
		  ,NEED_TITLE
		  ,ROW_NUMBER()
		  OVER (PARTITION BY [ID]
		  ,[PATH]
		  ,[NAME]
		  ,NEED_TITLE
		  ORDER BY
		  [ID]
		  ,[PATH]
		  ,[NAME]
		  ,NEED_TITLE
		  ) RN
	  FROM [dbo].[WORKBOOK_DETAILS]
  ) as res where RN =1 )as qa group by [FOLDER_SITE]


df = pd.read_sql(sql_plnt_rec, cnxn)  

print(df)