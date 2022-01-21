import tkinter as tk
from tkinter import filedialog


root=tk.Tk()

ent1=tk.Entry(root,font=40)
ent1.grid(row=2,column=2)

def browsefunc():
    filename =tkFileDialog.askopenfilename(filetypes=(("tiff files","*.tiff"),("All files","*.*")))
    ent1.insert(tk.END, filename) # add this

b1=tk.Button(root,text="DEM",font=40,command=browsefunc)
b1.grid(row=2,column=4)

root.mainloop()
