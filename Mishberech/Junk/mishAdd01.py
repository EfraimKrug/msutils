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
####################################################################################################
class mishAdd01:
    def __init__(self, master, ds, callback):
        self.ds = ds

        self.master = master
        self.master.configure(bg="teal", pady=34, padx=17)
        self.master.geometry('500x500')
        self.master.title('Add a MishBerech')

        self.frame = tk.Frame(self.master, width=460, height=360)
        self.frame.configure(bg="teal", pady=2, padx=2)
        self.frame.grid(row=1, column=1)

        self.callback = callback

        #self.tkvar = ''
        self.runProcess()

    def showData(self):
        row_num = 6

        self.l0 = tk.Label(self.frame, text="Member Name: ", bg="teal", fg="yellow")
        self.l0.grid(row=3, column=1, padx=4, pady=4, sticky=tk.W)

        self.txt0 = tk.Entry(self.frame, width=25,  borderwidth=2, relief="sunken")
        self.txt0.focus()
        self.txt0.grid(row=3, column=3, columnspan=2, padx=4, pady=4, sticky=tk.E)

        self.l1 = tk.Label(self.frame, text="Choleh's Hebrew Name: ", bg="teal", fg="yellow")
        self.l1.grid(row=5, column=1, padx=4, pady=4, sticky=tk.W)

        self.txt1 = tk.Entry(self.frame, width=25,  borderwidth=2, relief="sunken")
        self.txt1.focus()
        self.txt1.grid(row=5, column=3, columnspan=2, padx=4, pady=4, sticky=tk.E)

        self.l2 = tk.Label(self.frame, text="Member Email: ", bg="teal", fg="yellow")
        self.l2.grid(row=7, column=1, padx=4, pady=4, sticky=tk.W)

        self.txt2 = tk.Entry(self.frame, width=25,  borderwidth=2, relief="sunken")
        self.txt2.focus()
        self.txt2.grid(row=7, column=3, columnspan=2, padx=4, pady=4, sticky=tk.E)

        self.button01 = tk.Button(self.frame, text="Save", command=self.getData)
        self.button01.grid(row=2, column=7, padx=4, pady=4, sticky=tk.EW)

        #self.showHeadlines()
        last = ""

    def getData(self):
        hold = dict()
        nameIn = self.txt0.get()
        if len(nameIn.strip()) < 5:
            return

        email = ''
        f = False
        target = []
        for dct in self.ds:
            if nameIn in dct.keys():
                email = dct[nameIn][0][3]
                target = dct[nameIn]
                f = True

        if not f:
            self.ds.append(dict())
            self.ds[len(self.ds)-1][nameIn] = []
            self.ds[len(self.ds)-1][nameIn].append([self.txt1.get()])
            self.ds[len(self.ds)-1][nameIn][len(self.ds[len(self.ds)-1][nameIn])-1].append(datetime.today().strftime('%Y-%m-%d'))
            self.ds[len(self.ds)-1][nameIn][len(self.ds[len(self.ds)-1][nameIn])-1].append('')
            txt = self.txt2.get()
            if txt.find('@') < 0:
                self.ds[len(self.ds)-1][nameIn][len(self.ds[len(self.ds)-1][nameIn])-1].append(email)
            else:
                self.ds[len(self.ds)-1][nameIn][len(self.ds[len(self.ds)-1][nameIn])-1].append(txt)
        else:
            #print(self.ds[len(self.ds)-1])
            target.append([self.txt1.get()])
            target[len(target)-1].append(datetime.today().strftime('%Y-%m-%d'))
            target[len(target)-1].append('')
            txt = self.txt2.get()
            if txt.find('@') < 0:
                target[len(target)-1].append(email)
            else:
                target[len(target)-1].append(self.txt2.get())

        self.callback()

    def printData(self):
        for d in self.ds:
            print("#"*50)
            for x in d:
                print("Member: " + x)
                for a in d[x]:
                    print("\tName: " + str(a[0]))
                    print("\tDate: " + str(a[1]))
                    print("\tNotes: " + str(a[2]))
                    print("\tEmail: " + str(a[3]))



    def runProcess(self):
        self.showData()
        self.getData()
