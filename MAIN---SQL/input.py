import tkinter as tk
from tkinter import ttk
import sys
import os
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo

# create the root window
root = tk.Tk()
root.title('Tkinter Open File Dialog')
root.resizable(False, False)
root.geometry('300x150')


def select_file():
    filetypes = (
        ('text files', '*.xlsx'),
        ('All files', '*.*')
    )

    Filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    showinfo(
        title='Selected File',
        message=Filename
    )


# open button
open_button = ttk.Button(
    root,
    text='Open a File',
    command=select_file
)

open_button.pack(expand=True)


# run the application
root.mainloop()

import pypyodbc as pyodbc  # pip install pypyodbc
#import pyodbc as pyodbc
import pandas as pd # pip install pandas

"""
Step 1. Importing dataset from Excel/CSV
"""
filename = get(Filename)
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
    conn = pyodbc.connect('DRIVER={SQL Server};SERVER=DESKTOP-SFGBI2J;DATABASE=Users;UID=as;PWD=1234')
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
