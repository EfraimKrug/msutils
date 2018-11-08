from tkinter import *
import tkMessageBox
import subprocess
import os
from os import walk

batdir = r'c:\Users\KTM\BAT'
basedir = r'c:\Users\KTM\python\msutils'
yahrdir = r'\yahr'

root = Tk()
root.geometry('500x500')

def runYahr():
    subprocess.call([batdir + '\yahr.bat'], shell=False)

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
button1 = Button(frame, text='Build monthly yahrzeit list', command=doYahr01)
button1.pack(side=RIGHT)

button1.place(x=50, y=100, bordermode=OUTSIDE, height=30, width=300)

frame.pack()

root.mainloop()
