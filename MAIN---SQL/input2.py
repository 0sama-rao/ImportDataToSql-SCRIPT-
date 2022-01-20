import sys
import os
from tkinter import *

window=Tk()

window.title("Running Python Script")
window.geometry('550x200')

def run():
    os.system('opencv_video.py')

btn = Button(window, text="Click Me", bg="black", fg="white",command=run)
btn.grid(column=0, row=0)

window.mainloop()
