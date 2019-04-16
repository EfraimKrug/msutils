#!/usr/bin/env python
# -*- coding: utf-8 -*-
import subprocess
import os
import sys
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
from functools import partial

import smtplib
from Profile import *
from mishAdd01 import *
####################################################################################################
class mishDisplay01:
    def __init__(self, master, ds, callback):
        self.ds = ds

        self.master = master
        self.master.configure(bg="teal", pady=34, padx=17)
        self.master.geometry('500x800')
        self.master.title('MishBerech')

        self.frame = tk.Frame(self.master, width=460, height=360)
        self.frame.configure(bg="teal", pady=2, padx=2)
        self.frame.grid(row=1, column=1)

        self.tkvar = ''
        self.runProcess()
        self.callback = callback

# on change dropdown value
    def change_dropdown(self, *args):
        print( self.tkvar.get() )

    def refresh(self):
        self.clearData()
        self.showHeadlines()
        self.showData()
        self.callback()

    def mishAdd(self, *args):
        self.newWindow = tk.Toplevel(self.master)
        self.app = mishAdd01(self.newWindow, self.ds, self.refresh)

    def showHeadlines(self):
        self.headline01 = tk.Label(self.frame, text="Member", bg="teal", fg="yellow")
        self.headline01.grid(row=3, column=2, padx=4, pady=2, sticky=tk.W)

        self.headline02 = tk.Label(self.frame, text="Choleh", bg="teal", fg="yellow")
        self.headline02.grid(row=3, column=4, columnspan=3, padx=4, pady=2, sticky=tk.W)

        self.headline03 = tk.Label(self.frame, text="Date", bg="teal", fg="yellow")
        self.headline03.grid(row=3, column=9, padx=4, pady=2, sticky=tk.W)

        self.button01 = tk.Button(self.frame, text="Add", command=self.mishAdd)
        self.button01.grid(row=3, column=24, padx=4, pady=4, sticky=tk.EW)

    def clearData(self):
        for widget in self.frame.winfo_children():
            widget.destroy()

    def showData(self):

        total = 0
        self.label01 = []
        self.label02 = []
        self.label03 = []
        row_num = 6
        self.showHeadlines()
        last = ""

        for d in self.ds:
            for member in d:
                for a in d[member]:
                    memberOut = member
                    if last == member:
                        memberOut = ""

                    last = member
                    self.label01.append(tk.Label(self.frame, text=str(memberOut), bg="teal", fg="yellow"))
                    self.label01[len(self.label01)-1].grid(row=row_num, column=2, padx=4, pady=4, sticky=tk.NW)

                    self.label02.append(tk.Label(self.frame, text=str(a[0]), bg="teal", fg="yellow"))
                    self.label02[len(self.label02)-1].grid(row=row_num, column=4,  columnspan=3, padx=4, pady=4, sticky=tk.NW)

                    self.label03.append(tk.Label(self.frame, text=str(a[1])[0:10], bg="teal", fg="yellow"))
                    self.label03[len(self.label03)-1].grid(row=row_num, column=9, padx=4, pady=4, sticky=tk.NW)

                    row_num += 1

    def printData(self):
        for d in self.ds:
            print("#"*50)
            for x in d:
                print("Member: " + x)
                for a in d[x]:
                    print("\tName: " + str(a[0]))
                    if len(a) > 0:
                        print("\tDate: " + str(a[1]))
                    if len(a) > 1:
                        print("\tNotes: " + str(a[2]))
                    if len(a) > 2:
                        print("\tEmail: " + str(a[3]))

        #print(self.ds)

    def runProcess(self):
        self.showData()
