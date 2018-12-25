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

def get_file_yahr():
    filename = filedialog.askopenfilename( initialdir=basedir, title="select file", filetypes=(("excel files", "*.xlsx"), ("excel files", "*.xlsx")))
    if(len(filename) > 3):
        shutil.copyfile(filename, basedir + "\\shulCloud\\yahrzeits.xlsx")

def get_file_people():
    filename = filedialog.askopenfilename( initialdir=basedir, title="select file", filetypes=(("excel files", "*.xlsx"), ("excel files", "*.xlsx")))
    if(len(filename) > 3):
        shutil.copyfile(filename, basedir + "\\shulCloud\\people.xlsx")

def get_file_transactions():
    filename = filedialog.askopenfilename( initialdir=basedir, title="select file", filetypes=(("excel files", "*.xlsx"), ("excel files", "*.xlsx")))
    if(len(filename) > 3):
        shutil.copyfile(filename, basedir + "\\shulCloud\\transactions.xlsx")

def get_paypal_file():
    filename = filedialog.askopenfilename( initialdir=basedir, title="select file", filetypes=(("excel files", "*.xlsx"), ("excel files", "*.xlsx")))
    if(len(filename) > 3):
        shutil.copyfile(filename, basedir + "\\shulCloud\\peoplePayPal.xlsx")

def getEmails():
    subprocess.call([batdir + '\\getEmails.bat'], shell=False)

def getNames():
    subprocess.call([batdir + '\\getNames.bat'], shell=False)

def runYahr():
    subprocess.call([batdir + '\\yahr.bat'], shell=False)

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

def leftclick(event):
    print("left")

def doYahr01():
    getFilenames()
    print("doYahr01 - stuff")

frame = Frame(root, width=600, height=400)
button0 = Button(frame, text="Load New Yahrzeit File", command = get_file_yahr)
button1 = Button(frame, text="Load New People File", command = get_file_people)
button2 = Button(frame, text="Load New Transaction File", command = get_file_transactions)
button3 = Button(frame, text="Load New PayPal Transaction File", command = get_paypal_file)

buttonD1 = Button(frame, text='Build Monthly Yahrzeit List', command=doYahr01)
buttonD2 = Button(frame, text='Yahrzeit Names for Bulletin', command=getNames)
buttonD3 = Button(frame, text='Get Email Addresses', command=getEmails)

button0.pack(side=RIGHT)
button1.pack(side=RIGHT)
button2.pack(side=RIGHT)
button3.pack(side=RIGHT)

buttonD1.pack(side=RIGHT)
buttonD2.pack(side=RIGHT)
buttonD3.pack(side=RIGHT)

button0.place(x=50, y=50, bordermode=OUTSIDE, height=30, width=300)
button1.place(x=50, y=90, bordermode=OUTSIDE, height=30, width=300)
button2.place(x=50, y=120, bordermode=OUTSIDE, height=30, width=300)
button3.place(x=50, y=160, bordermode=OUTSIDE, height=30, width=300)

buttonD1.place(x=75, y=250, bordermode=OUTSIDE, height=30, width=200)
buttonD2.place(x=75, y=290, bordermode=OUTSIDE, height=30, width=200)
buttonD3.place(x=75, y=330, bordermode=OUTSIDE, height=30, width=200)

frame.pack()

root.mainloop()
