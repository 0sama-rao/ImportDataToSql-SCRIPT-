import pypyodbc as odbc # pip install pypyodbc
import pandas as pd # pip install pandas

"""
Step 1. Importing dataset from CSV
"""
df = pd.read_excel('test101.xlsx', [0])

"""
Step 2.2 Specify columns we want to import
"""
columns = ['SR_NO', 'NTN', 'NAME', 'BUSINESS_NAME']

df_data = df[columns]
records = df_data.values.tolist()

"""
Step 3.1 Create SQL Servre Connection String
"""
DRIVER = 'SQL Server'
SERVER_NAME = ('DESKTOP-SFGBI2J')
DATABASE_NAME = ('Users')

def connection_string(driver, server_name, database_name):
    conn_string = f"""
        DRIVER={{{driver}}};
        SERVER={server_name};
        DATABASE={database_name};
        Trust_Connection=yes;
        username='as';
        password= '1234';
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
    VALUES (?, ?, ?, ?)
'''

try:
    cursor = conn.cursor()
    cursor.executemany(sql_insert, records)
    cursor.commit();
except Exception as e:
    cursor.rollback()
    print(str(e[1]))
finally:
    print('Task is complete.')
    cursor.close()
    conn.close()
