from datetime import *
from tkinter import *
from tkinter import filedialog
import tkMessageBox
import subprocess
import os
import time
from os import walk
import sys
import csv
import openpyxl
from openpyxl import Workbook
from functools import partial
#########################################################
# get parent directory...
#########################################################

sys.path.append(os.getcwd())
sys.path.append(os.getcwd()[0:os.getcwd().rfind('\\')])

import shutil
from Profile import *

root = Tk()
root.geometry('400x400')
root.title('Kadima Toras-Moshe System')

#fileSwitches
noYahrzeitFileSw = False
noYahrzeitFIXSw = False
noTransactionFileSw = False
noTransactionFIXSw = False

# works
def refreshFile(fileName):
    wb = openpyxl.Workbook()
    ws = wb.active

    if 'yahrzeits' in fileName:
        newFileName = basedir + r'\\shulCloud\\yahrzeits.xlsx'
    if 'transactions' in fileName:
        newFileName = basedir + r'\\shulCloud\\transactions.xlsx'


    with open(fileName) as f:
        reader = csv.reader(f, delimiter=':')
        for row in reader:
            ws.append(row)

    try:
        os.remove(newFileName)
    except:
        print('no file to remove: ' + newFileName)
    wb.save(newFileName)

def getNames():
    subprocess.call([batdir + '\\getNames.bat'], shell=False)
# works
def runYahr():
    subprocess.call([batdir + '\\yahr.bat'], shell=False)

def openBooks():
    subprocess.call([batdir + '\\openBooks.bat'], shell=False)

def openDeposit():
    subprocess.call([batdir + '\\openDeposit.bat'], shell=False)

def sendAliyos():
    subprocess.call([batdir + '\\aliyos.bat'], shell=False)

def sendMishberech():
    subprocess.call([batdir + '\\sendMish.bat'], shell=False)

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

def getFileAge(fileName):
    # print(fileName)
    try:
        fileDate = time.ctime(os.path.getmtime(fileName))
        currDateDt = datetime.now()
        fileDateDt = datetime.strptime(fileDate, '%a %b %d %H:%M:%S %Y')
        return (currDateDt - fileDateDt).days
    except:
        return 365


def noExcelFile(fname):
    f = []

    for (dirpath, dirnames, filenames) in walk(basedir):
        f.extend(filenames)
    for x in f:
        if (x == fname + '.xlsx'):
            fname = basedir + r'\\shulCloud\\' + fname + '.xlsx'
            fAge = getFileAge(fname)
            if fAge > 5:
                os.remove(fname)
                return True

            return False

    return True

def fixFile(fname):
    f = open(downloadPath + fname + '.csv')
    reader = csv.reader(f)

    wb = Workbook()
    dest_filename = basedir + r'\\shulCloud\\' + fname + '.xlsx'

    ws = wb.worksheets[0]
    ws.title = fname

    wsRow = 2
    for row_index, row in enumerate(reader):
        wsCol = 1
        for column_index, cell in enumerate(row):
            ws.cell(row=wsRow,column=wsCol).value = cell
            wsCol += 1
        wsRow += 1

    wb.save(filename = dest_filename)
    setFileSwitches()
    buildScreen()

####################################################################################
def doYahr01():
    # getFilenames()
    runYahr()


def buildScreen():
    if noYahrzeitFIXSw:
        # print('noYahrzeitFIXSw')
        buttonD1 = Button(frame, text='Build Monthly Yahrzeit List', bg='yellow', fg='red')
        buttonD1R = Button(frame, text='FIX', bg='yellow', fg='red', command=partial(fixFile, 'yahrzeits'))
        buttonD1R.pack(side=RIGHT)
        buttonD1R.place(x=285, y=50, bordermode=OUTSIDE, height=30, width=30)
        buttonD2 = Button(frame, text='Yahrzeit Names for Bulletin', bg='tan', fg='red')
    else:
        buttonD1 = Button(frame, text='Build Monthly Yahrzeit List', bg='yellow',  fg='black', command=doYahr01)
        buttonD1R = Button(frame, text='OK', bg='yellow', fg='black')
        buttonD1R.pack(side=RIGHT)
        buttonD1R.place(x=285, y=50, bordermode=OUTSIDE, height=30, width=30)
        buttonD2 = Button(frame, text='Yahrzeit Names for Bulletin', bg='tan',  fg='black', command=getNames)

    buttonD3 = Button(frame, text='Open Check/Deposit View', bg='teal', command=openBooks)
    buttonD4 = Button(frame, text='Open Deposit View', bg='purple', command=openDeposit)

    if noYahrzeitFileSw:
        message = "Please DOWNLOAD and FIX your 'yahrzeits' file"
        labelB14 = Label(frame, text=message, bg="black", fg="red", font='Helvetica 8')
        labelB14.pack(side=RIGHT)
        labelB14.place(x=1, y=1, bordermode=OUTSIDE, height=40, width=270)
    else:
        if noTransactionFileSw:
            message = "Please DOWNLOAD and FIX your 'transactions' file"
            labelB14 = Label(frame, text=message, bg="black", fg="red", font='Helvetica 8')
            labelB14.pack(side=RIGHT)
            labelB14.place(x=1, y=1, bordermode=OUTSIDE, height=40, width=270)

    if noTransactionFIXSw:
        buttonD14 = Button(frame, text='Send Email for aliyos', bg='pink', fg='red')
        buttonD14R = Button(frame, text='FIX', bg='yellow', fg='red', command=partial(fixFile, 'transactions'))
        buttonD14R.pack(side=RIGHT)
        buttonD14R.place(x=285, y=270, bordermode=OUTSIDE, height=30, width=30)
    else:
        buttonD14 = Button(frame, text='Send Email for aliyos', bg='pink', fg='black', command=sendAliyos)
        buttonD14R = Button(frame, text='OK', bg='yellow', fg='black')
        buttonD14R.pack(side=RIGHT)
        buttonD14R.place(x=285, y=270, bordermode=OUTSIDE, height=30, width=30)

    buttonD15 = Button(frame, text='Send Email for mishberech', bg='orange', command=sendMishberech)

    buttonD1.pack(side=RIGHT)
    buttonD2.pack(side=RIGHT)
    buttonD3.pack(side=RIGHT)
    buttonD4.pack(side=RIGHT)

    buttonD14.pack(side=RIGHT)
    buttonD15.pack(side=RIGHT)

    buttonD1.place(x=75, y=50, bordermode=OUTSIDE, height=30, width=200)
    buttonD2.place(x=75, y=90, bordermode=OUTSIDE, height=30, width=200)
    buttonD3.place(x=75, y=130, bordermode=OUTSIDE, height=30, width=200)
    buttonD4.place(x=75, y=170, bordermode=OUTSIDE, height=30, width=200)

    buttonD14.place(x=75, y=270, bordermode=OUTSIDE, height=30, width=200)
    buttonD15.place(x=75, y=310, bordermode=OUTSIDE, height=30, width=200)

def setFileSwitches():
    global noYahrzeitFileSw
    global noYahrzeitFIXSw
    global noTransactionFileSw
    global noTransactionFIXSw

    noYahrzeitFileSw = False
    noYahrzeitFIXSw = False
    noTransactionFileSw = False
    noTransactionFIXSw = False

    if noExcelFile('yahrzeits'):
        noYahrzeitFIXSw = True
        downloadedCSV = downloadPath + "yahrzeits.csv"
        fileAge = getFileAge(downloadedCSV)
        if fileAge > 25:
            noYahrzeitFileSw = True

    if noExcelFile('transactions'):
        noTransactionFIXSw = True
        downloadedCSV = downloadPath + "transactions.csv"
        fileAge = getFileAge(downloadedCSV)
        if fileAge > 25:
            noTransactionFileSw = True

    # downloadedCSV = downloadPath + "transactions.csv"
    # fileAge = getFileAge(downloadedCSV)
    # if fileAge > 1:
    #     noTransactionFileSw = True

frame = Frame(root, width=400, height=400)

setFileSwitches()
buildScreen()
frame.pack()
root.mainloop()
