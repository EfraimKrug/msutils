from tkinter import *
from tkinter import filedialog
import tkMessageBox
import subprocess
import os
from os import walk
import sys
#########################################################
# get parent directory...
#########################################################

sys.path.append(os.getcwd())
sys.path.append(os.getcwd()[0:os.getcwd().rfind('\\')])

import shutil
from Profile import *

root = Tk()
root.geometry('400x400')
# works
def getNames():
    subprocess.call([batdir + '\\getNames.bat'], shell=False)
# works
def runYahr():
    subprocess.call([batdir + '\\yahr.bat'], shell=False)

def openBooks():
    subprocess.call([batdir + '\\openBooks.bat'], shell=False)

def sendAliyos():
    subprocess.call([batdir + '\\aliyos.bat'], shell=False)

def sendMishberech():
    subprocess.call([batdir + '\\sendMish.bat'], shell=False)

def checkOldFile(fname):
    currentFile =str(os.path.getmtime(fname))
    fout = open(basedir + "\\mainBox\\.filestat", "r")
    oldFile = str(fout.read())
    fout.close()

    if currentFile == oldFile:
        tkMessageBox.showinfo("Old File", "Please download a new file!")
    else:
        fout = open(basedir + "\\mainBox\\.filestat", "w")
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
            fname = mypath + r'\\shulCloud\\yahrzeits.xlsx'
            checkOldFile(fname)

def doYahr01():
    getFilenames()

frame = Frame(root, width=400, height=400)
buttonD1 = Button(frame, text='Build Monthly Yahrzeit List', bg='yellow', command=doYahr01)
buttonD2 = Button(frame, text='Yahrzeit Names for Bulletin', bg='tan', command=getNames)
buttonD3 = Button(frame, text='Open Check/Deposit View', bg='teal', command=openBooks)
buttonD4 = Button(frame, text='Send Email for aliyos', bg='pink', command=sendAliyos)
buttonD5 = Button(frame, text='Send Email for mishberech', bg='pink', command=sendMishberech)

buttonD1.pack(side=RIGHT)
buttonD2.pack(side=RIGHT)
buttonD3.pack(side=RIGHT)
buttonD4.pack(side=RIGHT)
buttonD5.pack(side=RIGHT)

buttonD1.place(x=75, y=50, bordermode=OUTSIDE, height=30, width=200)
buttonD2.place(x=75, y=90, bordermode=OUTSIDE, height=30, width=200)
buttonD3.place(x=75, y=130, bordermode=OUTSIDE, height=30, width=200)
buttonD4.place(x=75, y=270, bordermode=OUTSIDE, height=30, width=200)
buttonD5.place(x=75, y=310, bordermode=OUTSIDE, height=30, width=200)

frame.pack()
root.mainloop()
