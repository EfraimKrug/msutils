#!/usr/bin/env python
# -*- coding: utf-8 -*-
import subprocess
import os
import sys
import winreg
#########################################################
# get parent directory...
#sys.path.append(os.getcwd())
#sys.path.append(os.getcwd()[0:os.getcwd().rfind('\\')])
#######################################################
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import tkinter.scrolledtext as tkst
#######################################################
import csv
import shutil
import requests

from datetime import datetime
from datetime import time
from datetime import date

from openpyxl import load_workbook
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl.utils import get_column_letter

from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment

from functools import partial

import smtplib
from Profile import *
from SearchNames import *
from checkDisplay02 import *
from checkDisplay03 import *
from checkDisplay04 import *
#from AlefBet import *
#from periodProcess import *
####################################################################################################
class checkDisplay01:
    def __init__(self, master):
        self.cashcheckSwitch = ''
        self.ds = dict()        # {check_name: [check_number, memo, check_date, arrival_date, check_amount, check_image],
        self.sdata = dict()     # sheet by sheet...
        self.pdata = dict()
        self.cdata = []     # cash
        self.searchObj = ''
        self.depositName = ''

        self.master = master
        self.master.configure(bg="teal", pady=34, padx=17)
        self.master.geometry('700x700')
        self.master.title('Kadima Toras-Moshe Check Tracking')

        self.frame = tk.Frame(self.master, width=460, height=360)
        self.frame.configure(bg="teal", pady=2, padx=2)
        self.frame.grid(row=1, column=1)

        self.people = []
        self.pages = []
        self.tkvar = ''
        self.EXCELEXE = ''

        self.dropInit = True
        self.runProcess()

    def getExcel(self):
        handle = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")

        num_values = winreg.QueryInfoKey(handle)[1]
        for i in range(num_values):
            for x in winreg.EnumValue(handle, i):
                if(str(x).find("EXCEL") > -1):
                    self.EXCELEXE = x

    def getEachKeyStroke(self, key):
        print("here")

    def setOutlabel(self, newLabel):
        self.OutLabel['text'] = newLabel

    def upArrow(self, key):
        print("here")

    def downArrow(self, key):
        print("here")

    # def pre_new_window(self, key):
    #     self.new_window()
    #
    # def new_empty_window(self):
    #     self.newWindow = tk.Toplevel(self.master)
    #
    # def new_config_window(self):
    #     self.newWindow = tk.Toplevel(self.master)

    def show_image(self, img):
        fileName = checkDir + img + ".pdf"
        path_to_pdf = os.path.abspath(fileName)
        path_to_acrobat = os.path.abspath(AcrobatPath)
        process = subprocess.Popen([path_to_acrobat, '/A', 'page=1', path_to_pdf], shell=False, stdout=subprocess.PIPE)
        process.wait()

    def openDailyLog(self):
        wb = load_workbook(dailyLogDir + '\\dailyLog.xlsx')
        return wb

    def createSheet(self):
        sheetNameNew = True
        dt = datetime.today().strftime('%B-%d')
        da = dt.split('-')
        sheetName = da[0] + str(da[1])
        sheetName = "JUNK1"

        dailyLog = self.openDailyLog()
        for name in dailyLog.sheetnames:
            if name == sheetName:
                sheetNameNew = False

        if not sheetNameNew:
            return

        newSheet = dailyLog.create_sheet(title = sheetName)
        
        newSheet = dailyLog[sheetName]
        self.buildPage(newSheet)
        dailyLog.save(dailyLogDir + '\\dailyLog.xlsx')

    def openNewSheet(self):
        self.createSheet()
        self.getExcel()
        os.system("start  \"" + self.EXCELEXE + "\" \"" + dailyLogDir + "\\dailyLog.xlsx\"")

    def buildPage(self, newSheet):
        al = Alignment(horizontal='center', vertical='center')
        newSheet.cell(row=1,column=1).value = datetime.today().strftime('%d-%B-%Y')
        newSheet.cell(row=1,column=7).value = "Deposit: "
        newSheet.cell(row=1,column=7).alignment = al

        newSheet.cell(row=2,column=2).value = "Name on Check"
        newSheet.cell(row=2,column=2).alignment = al
        newSheet.cell(row=2,column=3).value = "Memo"
        newSheet.cell(row=2,column=3).alignment = al
        newSheet.cell(row=2,column=4).value = "Date on Check"
        newSheet.cell(row=2,column=4).alignment = al
        newSheet.cell(row=2,column=5).value = "Amount"
        newSheet.cell(row=2,column=5).alignment = al
        newSheet.cell(row=2,column=6).value = "Image"
        newSheet.cell(row=2,column=6).alignment = al

        newSheet.cell(row=3,column=1).value = "Cash"
        newSheet.cell(row=3,column=1).alignment = al
        newSheet.cell(row=11,column=4).value = "Total: "
        newSheet.cell(row=11,column=5).value = "=SUM(E4:E10)"

        newSheet.cell(row=12,column=1).value = "Check No."
        newSheet.cell(row=12,column=1).alignment = al
        newSheet.cell(row=31,column=3).value = "Sub Total: "
        newSheet.cell(row=31,column=5).value = "=SUM(E13:E30)"
        newSheet.cell(row=32,column=3).value = "Grand Total: "
        newSheet.cell(row=32,column=5).value = "=SUM(E11,E31)"

        #newSheet.column_dimensions[0].width = 20.71
        newSheet.column_dimensions['A'].width = 20
        newSheet.column_dimensions['B'].width = 33
        newSheet.column_dimensions['C'].width = 51
        newSheet.column_dimensions['D'].width = 23
        newSheet.column_dimensions['E'].width = 12
        newSheet.column_dimensions['F'].width = 19
        newSheet.column_dimensions['G'].width = 11

    def parseName(self, name):
        day = name[-2:]
        month = name[0:-2]
        return (day, month)

    def loadRow(self, month, day, sheet, current_row):
        arr = []
        if(sheet.cell(row=current_row, column=2).value in self.ds):
            arr = self.ds[sheet.cell(row=current_row, column=2).value]

        name = month+day
        peep = sheet.cell(row=current_row, column=2).value

        newRow = [sheet.cell(row=current_row, column=1).value,
                  self.depositName,
                  sheet.cell(row=current_row, column=3).value,
                  str(sheet.cell(row=current_row, column=4).value)[0:10],
                  str(month) + "-" + str(day),
                  sheet.cell(row=current_row, column=5).value,
                  sheet.cell(row=current_row, column=6).value,
                  name]

        if not peep in self.people:
            self.people.append(peep)

        arr.append(newRow)
        self.ds[sheet.cell(row=current_row, column=2).value] = arr
        if not name in self.sdata:
            self.sdata[name] = dict()

        self.sdata[name][sheet.cell(row=current_row, column=2).value] = arr
        if not peep in self.pdata:
            self.pdata[peep] = dict()

        self.pdata[peep][sheet.cell(row=current_row, column=2).value] = arr


    def loadRowCash(self, month, day, sheet, current_row):
        arr = []
        if(sheet.cell(row=current_row, column=2).value in self.ds):
            arr = self.ds[sheet.cell(row=current_row, column=2).value]

        name = month+day
        peep = sheet.cell(row=current_row, column=2).value

        newRow = [sheet.cell(row=current_row, column=1).value,
                  self.depositName,
                  sheet.cell(row=current_row, column=3).value,
                  str(sheet.cell(row=current_row, column=4).value)[0:10],
                  str(month) + "-" + str(day),
                  sheet.cell(row=current_row, column=5).value,
                  sheet.cell(row=current_row, column=6).value,
                  name]

        self.cdata.append(newRow)

    def getSheet(self, name, sheet):
        (day, month) = self.parseName(name)
        if not name in self.pages:
            self.pages.append(name)

        self.depositName = str(sheet.cell(row=2,column=7).value)

        for r in range(3, sheet.max_row):
            if(str(sheet.cell(row=r,column=1).value).lower() == 'cash'):
                self.cashcheckSwitch = 'cash'
            if(str(sheet.cell(row=r,column=1).value).lower().find('check') > -1):
                self.cashcheckSwitch = 'check'

            if(sheet.cell(row=r, column=2).value and self.cashcheckSwitch.find('check') > -1):
                self.loadRow(month, day, sheet, r)

            if(sheet.cell(row=r, column=2).value and self.cashcheckSwitch.find('cash') > -1):
                self.loadRowCash(month, day, sheet, r)


# on change dropdown value
    def change_dropdown(self, *args):
        if self.dropInit:
            return

        print( self.tkvar.get() )

    def change_dropdown2(self, *args):
        if self.dropInit:
            self.dropInit = False
            return

        self.newWindow = tk.Toplevel(self.master)
        self.app = checkDisplay02(self.newWindow, self.tkvar2.get())

    def showCash(self):
        total = 0
        for line in self.cdata:
            total += line[5]
        total = "{:.2f}".format(float(total))
        total = "Cash: $" + total
        self.label07 = tk.Label(self.frame, text=total, bg="teal", fg="yellow", font='Helvetica 10 bold')
        self.label07.grid(row=1, column=18, padx=4, pady=4, sticky=tk.NW)

    def showPerson(self, name, args):
        #print(name)
        self.newWindow = tk.Toplevel(self.master)
        self.app = checkDisplay03(self.newWindow, name)

    def showDeposit(self, name, args):
        #print(name)
        self.newWindow = tk.Toplevel(self.master)
        self.app = checkDisplay04(self.newWindow, name)

    # link function to change change_dropdown
    def showData(self):
        #frame = tk.Frame(self.master, width=700, height=300)
        self.searchObj = SearchNames(34, self.ds)

        total = 0
        self.label01 = []
        self.label02 = []
        self.label03 = []
        self.label04 = []
        self.label05 = []
        self.label06 = []
        self.button01 = []

        self.peepFunctions = []
        self.peepArgs = []

        fileNames = []

        # Dictionary with options
        self.tkvar = tk.StringVar(self.master)
        self.tkvar2 = tk.StringVar(self.master)

        self.tkvar.trace('w', self.change_dropdown)
        self.tkvar2.trace('w', self.change_dropdown2)

        #choices = []
        self.tkvar.set(self.people[0]) # set the default option
        self.tkvar2.set(self.pages[0]) # set the default option

        #peoplePopup = tk.OptionMenu(self.frame, self.tkvar, *self.people)
        pagesPopup = tk.OptionMenu(self.frame, self.tkvar2, *self.pages)
        #tk.Label(self.frame, text="Choose a dish").grid(row = 1, column = 1)
        #peoplePopup.grid(row = 1, column =2)
        pagesPopup.grid(row = 1, column =4)
        self.button02 = tk.Button(self.frame, text="New Sheet", command=partial(self.openNewSheet))
        self.button02.grid(row=1, column=3, columnspan=1, padx=4, pady=4, sticky=tk.EW)

        self.responseLabel = tk.Label(self.frame, text="", bg="teal", fg="yellow")
        self.responseLabel.grid(row=1, rowspan=2, column=1, columnspan=4, padx=4, pady=4, sticky=tk.W)

        self.OutLabel = tk.Label(self.frame, text="", bg="teal", fg="yellow")
        self.OutLabel.grid(row=2, rowspan=1, column=1, columnspan=4, padx=4, pady=4, sticky=tk.W)

        row_num = 6
        self.headline01 = tk.Label(self.frame, text="Name", bg="teal", fg="yellow")
        self.headline01.grid(row=3, column=2, padx=4, pady=2, sticky=tk.W)

        self.headline02 = tk.Label(self.frame, text=" Date ", bg="teal", fg="yellow")
        self.headline02.grid(row=3, column=4, padx=4, pady=2, sticky=tk.W)

        self.headline03 = tk.Label(self.frame, text=" Deposit ", bg="teal", fg="yellow")
        self.headline03.grid(row=3, column=6, padx=4, pady=2, sticky=tk.W)

        self.headline04 = tk.Label(self.frame, text="Check #", bg="teal", fg="yellow")
        self.headline04.grid(row=3, column=8, padx=4, pady=2, sticky=tk.W)

        self.headline05 = tk.Label(self.frame, text="Amount", bg="teal", fg="yellow")
        self.headline05.grid(row=3, column=10, padx=4, pady=2, sticky=tk.W)

        self.headline06 = tk.Label(self.frame, text="Sheet Total", bg="teal", fg="yellow")
        self.headline06.grid(row=3, column=14, padx=4, pady=2, sticky=tk.W)

        self.headline07 = tk.Label(self.frame, text="Image", bg="teal", fg="yellow")
        self.headline07.grid(row=3, column=16, padx=4, pady=2, sticky=tk.W)

        self.showCash()
        sortedKeys = []
        for key in self.ds:
            sortedKeys.append(key)
        sortedKeys.sort()
        lastName = ''
        pEnt = ''
        totals = dict()

        for ent in sortedKeys:
            for e in self.ds[ent]:
                pEnt = ent
                if lastName == ent:
                    pEnt = ''
                lastName = ent
                self.label01.append(tk.Label(self.frame, text=pEnt, bg="teal", fg="yellow"))
                self.label01[len(self.label01)-1].grid(row=row_num, column=2, padx=4, pady=4, sticky=tk.NW)
                self.label01[len(self.label01)-1].bind("<Button-1>", partial(self.showPerson, pEnt))

                self.label02.append(tk.Label(self.frame, text=e[3], bg="teal", fg="yellow"))
                self.label02[len(self.label02)-1].grid(row=row_num, column=4, padx=4, pady=4, sticky=tk.NW)

                self.label03.append(tk.Label(self.frame, text=e[1], bg="teal", fg="yellow"))
                self.label03[len(self.label03)-1].grid(row=row_num, column=6, padx=4, pady=4, sticky=tk.NW)
                self.label03[len(self.label03)-1].bind("<Button-1>", partial(self.showDeposit, e[1]))

                self.label04.append(tk.Label(self.frame, text=e[0], bg="teal", fg="yellow"))
                self.label04[len(self.label04)-1].grid(row=row_num, column=8, padx=4, pady=4, sticky=tk.NW)

                fAmt = "{:.2f}".format(float(e[5]))
                self.label05.append(tk.Label(self.frame, text=fAmt, bg="teal", fg="yellow"))
                self.label05[len(self.label05)-1].grid(row=row_num, column=10, padx=4, pady=4, sticky=tk.NW)

                #self.label05.append(tk.Label(self.frame, text=e[6], bg="teal", fg="yellow"))
                #self.label05[len(self.label05)-1].grid(row=row_num, column=10, padx=4, pady=4, sticky=tk.NW)

                if e[7] in totals:
                    totals[e[7]] = totals[e[7]] + e[5]
                else:
                    totals[e[7]] = e[5]

                fAmt2 = "{:.2f}".format(float(totals[e[7]]))
                self.label06.append(tk.Label(self.frame, text="$" + str(fAmt2), bg="teal", fg="yellow"))
                self.label06[len(self.label06)-1].grid(row=row_num, column=14, padx=4, pady=4, sticky=tk.NW)

                self.button01.append(tk.Button(self.frame, text="View", command=partial(self.show_image, e[6])))
                self.button01[len(self.button01)-1].grid(row=row_num, column=16, columnspan=2, padx=4, pady=4, sticky=tk.EW)
                total = total + e[5]
                row_num += 1
                line = ''

    def sheetExists(self, sheet):
        dailyLog = self.openDailyLog()
        for name in dailyLog.sheetnames:
            if name == sheet:
                return True
        return False

    def runProcess(self):
        dailyLog = self.openDailyLog()
        for name in dailyLog.sheetnames:
            self.total = 0
            self.getSheet(name, dailyLog[name])
        self.showData()
        #print(self.pdata)
