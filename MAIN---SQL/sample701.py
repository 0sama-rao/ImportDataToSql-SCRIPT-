import pypyodbc as odbc # pip install pypyodbc
import pandas as pd # pip install pandas


"""

Step 1. Importing dataset from Excel/CSV
"""
filename = input("Input the Filename: ")
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
DRIVER = 'SQL Server'
SERVER_NAME = 'imldb'
DATABASE_NAME = 'UHASGASF'

def connection_string(driver, server_name, database_name):
    conn_string = f"""
        DRIVER={{{driver}}};
        SERVER={server_name};
        DATABASE={database_name};
        username='akdiml';
        password= 'imlccbmgmakd';
    """
    return conn_string

"""
Step 3.2 Create database connection instance
"""
try:
    conn = odbc.connect(connection_string(DRIVER, SERVER_NAME, DATABASE_NAME))
except odbc.DatabaseError as e:
    print('Database Error:')
    print(str(e.value[1]))
except odbc.Error as e:
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
    cursor.executemany(sql_insert,records)
    cursor.commit();
except Exception as e:
    cursor.rollback()
    print(str(e[1]))
finally:
    print('Task is complete.')
    cursor.close()
    conn.close()
