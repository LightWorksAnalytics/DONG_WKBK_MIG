# -*- coding: utf-8 -*-
"""
Created on Mon Aug  7 16:26:21 2017

@author: Alex
"""
import pandas as pd

path = 'O:\\Clients\\DONG\\DONG 02 - Asset Risk and Optimisation Suite\\02 Data\\01 Input\\02 Internal\\Test Workbooks\\Test Workbook - Good Only.xlsm'
#df = pd.read_excel(path, sheetname = '')

xl = pd.ExcelFile(path)

df = xl.parse('Plant Lookups')