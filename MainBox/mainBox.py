from tkinter import *
from tkinter import filedialog
import tkMessageBox
import subprocess
import os
from os import walk
import shutil

from profile import *

root = Tk()
root.geometry('500x500')

def get_file():
    filename = filedialog.askopenfilename( initialdir=basedir, title="select file", filetypes=(("excel files", "*.xlsx"), ("excel files", "*.xlsx")))
    if(len(filename) > 3):
        shutil.copyfile(filename, basedir + "\\yahrzeits.xlsx")

def AddChecks():
    subprocess.call([batdir + '\\addChecks.bat'], shell=False)

def runYahr():
    subprocess.call([batdir + '\\yahr.bat'], shell=False)

def checkOldFile(fname):
    currentFile =str(os.path.getmtime(fname))
    fout = open(".filestat", "r")
    oldFile = str(fout.read())
    fout.close()

    if currentFile == oldFile:
        tkMessageBox.showinfo("Old File", "Please download a new file!")
    else:
        fout = open(".filestat", "w")
        fout.write(str(os.path.getmtime(fname)))
        fout.close()
    runYahr()

def getFilenames():
    f = []
    mypath = basedir

    for (dirpath, dirnames, filenames) in walk(mypath):
        f.extend(filenames)
    for x in f:
        if (x == 'yahrzeits.xlsx'):
            print("found it: " + x)
            fname = mypath + r'\yahrzeits.xlsx'
            checkOldFile(fname)

def leftclick(event):
    print("left")

def doYahr01():
    getFilenames()
    print("doYahr01 - stuff")

frame = Frame(root, width=600, height=400)
button0 = Button(frame, text="Load New Yahrzeit File", command = get_file)
button1 = Button(frame, text='Build Monthly Yahrzeit List', command=doYahr01)
button2 = Button(frame, text='Add checks', command=AddChecks)
button0.pack(side=RIGHT)
button1.pack(side=RIGHT)
button2.pack(side=RIGHT)

button0.place(x=50, y=50, bordermode=OUTSIDE, height=30, width=300)
button1.place(x=50, y=100, bordermode=OUTSIDE, height=30, width=300)
button2.place(x=50, y=150, bordermode=OUTSIDE, height=30, width=300)

frame.pack()

root.mainloop()
