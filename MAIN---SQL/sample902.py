import pypyodbc as pyodbc # pip install pypyodbc
#import pyodbc as pyodbc
import pandas as pd # pip install pandas


"""
Step 1. Importing dataset from Excel/CSV
"""

filename = input("Input the Filename: ")
dfs = pd.read_excel(filename, usecols=['SR_NO','NTN'], sheet_name=None)

d = {}
for k, v in dfs.items():
    d[k] = pd.concat(df for df in dfs.values()).to_numpy().tolist()
print (d.keys())

#for test sheetnames

"""
Step 3.1 Create SQL Servre Connection String
"""
try:
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER=10.1.1.8;DATABASE=UHASGASF;UID=akdiml;PWD=imlccbmgmakd')
except pyodbc.DatabaseError as e:
    print('Database Error')
    print(str(e.value[1]))
except pyodbc.Error as e:
    print('Connection Error:')
    print(str(e.value[1]))
"""
Step 3.3 Create a cursor connection and insert records
"""

sql_insert = '''
    INSERT INTO FilerNonFiler
    VALUES (?,?, GETDATE())
'''

try:
    cursor = conn.cursor()
    cursor.executemany(sql_insert, d[k])
    cursor.commit();
except Exception as e:
    cursor.rollback()
    print(str(e[1]))
finally:
    print('Task is complete.')
    cursor.close()
    conn.close()
