import sqlite3
import pandas as pd 
from sqlalchemy import create_engine
from openpyxl import Workbook

import openpyxl





file = 'one.xlsx'

wb_obj = openpyxl.load_workbook(file)
sheet_obj = wb_obj.active.title 


engine = create_engine('sqlite://', echo=False)
df = pd.read_excel(file)

df.to_sql(sheet_obj,engine,if_exists='replace',index=False)


df5 = pd.read_sql_query("select COUNT(*) as v,depot from sheet1 where MATERIAL='FTB15801' ",engine)

# final = pd.DataFrame(result)

print(df5)