#!/usr/bin/env python
# -*- coding: utf-8 -*-
import os
import sys
import winreg

#########################################################
# get parent directory...
sys.path.append(os.getcwd())
sys.path.append(os.getcwd()[0:os.getcwd().rfind('\\')])

import csv
import shutil
import requests

from datetime import datetime
from datetime import time
from datetime import date

import smtplib
from Profile import *


def downloadXLSX(fileName):
    print("download")
    url = "https://images.shulcloud.com/616/uploads/mishberech/MishBerech.xlsx"

    response = requests.get(url, stream=True)
    with open(fileName, 'wb') as out_file:
        shutil.copyfileobj(response.raw, out_file)
    del response

def getExcel():
    print("getExcel")
    global EXCELEXE
    handle = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")

    num_values = winreg.QueryInfoKey(handle)[1]
    for i in range(num_values):
        for x in winreg.EnumValue(handle, i):
            if(str(x).find("EXCEL") > -1):
                EXCELEXE = x

def openFile(fileName):
    print("openFile")
    getExcel()
    print(fileName)
    if fileName.find('Mish') > -1:
        print("opening now...")
        os.system("start  \"" + EXCELEXE + "\" \"" + fileName + "\"")

def runProcess():
    print("runProcess")
    fileName = basedir + "\\MishBerech.xlsx"
    downloadXLSX(fileName)
    openFile(fileName)

runProcess()
