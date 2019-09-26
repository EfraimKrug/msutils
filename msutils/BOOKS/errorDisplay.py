#!/usr/bin/env python
# -*- coding: utf-8 -*-
# import subprocess
# import os
# import sys
# import winreg
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
# import csv
# import shutil
# import requests

from datetime import datetime
from datetime import time
from datetime import date

# from openpyxl import load_workbook
# from openpyxl.styles import colors
# from openpyxl.styles import Font, Color
# from openpyxl.utils import get_column_letter
#
# from openpyxl.styles.borders import Border, Side
# from openpyxl.styles import Alignment

from functools import partial

# import smtplib
from Profile import *
#from AlefBet import *
#from periodProcess import *
####################################################################################################
class errorDisplay:
    def __init__(self, master, error):
        self.master = master
        self.master.configure(bg="teal", pady=34, padx=17)
        self.master.geometry('500x250')
        self.master.title('Ouch! Not good...')

        self.frame = tk.Frame(self.master, width=260, height=260)
        self.frame.configure(bg="teal", pady=2, padx=2)
        self.frame.grid(row=1, column=1)
        self.error = error

        self.runProcess()

    def doSomething(self):
        self.headline03.destroy()
        self.headline03 = tk.Label(self.frame, text="But not that.", bg="teal", fg="red", font='Helvetica 14 bold')
        self.headline03.grid(row=3, column=3,  columnspan=4, padx=4, pady=2, sticky=tk.W)

    def showData(self):
        self.button02 = tk.Button(self.frame, text="Do Something", command=partial(self.doSomething))
        self.button02.grid(row=1, column=3, columnspan=1, padx=10, pady=10, sticky=tk.EW)

        self.headline03 = tk.Label(self.frame, text=self.error, bg="teal", fg="red", font='Helvetica 14 bold')
        self.headline03.grid(row=3, column=3,  columnspan=4, padx=4, pady=2, sticky=tk.W)

    def runProcess(self):
        self.showData()
