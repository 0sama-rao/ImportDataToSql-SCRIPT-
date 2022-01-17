import pypyodbc as odbc # pip install pypyodbc
import pandas as pd # pip install pandas
import os
import tkinter as tk
from tkinter.filedialog import askopenfilename
#import pandas as pd


def import_csv_data():
    global v
    filename  = askopenfilename()
    print(filename )
    v.set(filename )
    df = pd.read_csv(filename )

root = tk.Tk()
tk.Label(root, text='File Path').grid(row=0, column=0)
v = tk.StringVar()
entry = tk.Entry(root, textvariable=v).grid(row=0, column=1)
tk.Button(root, text='Browse Data Set',command=import_csv_data).grid(row=1, column=0)
tk.Button(root, text='Close',command=root.destroy).grid(row=1, column=1)
root.mainloop()

"""

Step 1. Importing dataset from Excel/CSV
"""

filename = input("Input the Filename: ")
dfs = pd.read_excel(filename, sheet_name=None)
print (dfs.keys()) #print the name of sheets
d = {k: v[['SR_NO', 'NTN', 'NAME']].values.tolist()
        for k, v in  pd.read_excel(filename, sheet_name=None).items()}

for k, v in dfs.items():
    columns = ['SR_NO', 'NTN', 'NAME']
    df_data = v[columns]
    records = df_data.values.tolist()
    d[k] = records

#for test sheetnames
print (d.keys())




"""
Step 3.1 Create SQL Servre Connection String
"""
DRIVER = 'SQL Server'
SERVER_NAME = 'DESKTOP-SFGBI2J'
DATABASE_NAME = 'Users'

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
    INSERT INTO eg205
    VALUES (?,?,?)
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
