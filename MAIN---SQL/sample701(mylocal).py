import pypyodbc as pyodbc  # pip install pypyodbc
#import pyodbc as pyodbc
import pandas as pd # pip install pandas

"""
Step 1. Importing dataset from Excel/CSV
"""
filename = input("Input File name: ")
dfs = pd.read_excel(filename, usecols=['SR_NO','NTN'], sheet_name=None)

d = {}
for k, v in dfs.items():
    d[k] = pd.concat(df for df in dfs.values()).to_numpy().tolist()
    records = d[k]
print (d.keys())
#for test sheetnames

"""
Step 3.1 Create SQL Servre Connection String
"""
try:
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=DESKTOP-SFGBI2J;DATABASE=Users;UID=as;PWD=1234')
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
    INSERT INTO eg207
    VALUES (?,?, GETDATE())
'''

try:
    cursor = cnxn.cursor()
    cursor.executemany(sql_insert,records)
    cursor.commit();
except Exception as e:
    cursor.rollback()
    print(str(e[1]))
finally:
    print('Task is complete.')
    cursor.close()
    cnxn.close()
