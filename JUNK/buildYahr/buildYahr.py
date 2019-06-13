#
# buildYahr
#   get directories
#   build yahr.bat
#
import os
import sys
#########################################################
# get parent directory...
sys.path.append(os.getcwd())
sys.path.append(os.getcwd()[0:os.getcwd().rfind('\\')])

from Profile import *
import winreg

EXCELEXE = r'C:\Program Files\Microsoft Office\root\Office\EXCEL.EXE'
WORDEXE = r'C:\Program Files\Microsoft Office\root\Office\WINWORD.EXE'

def getExcel():
    global EXCELEXE
    handle = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")

    num_values = winreg.QueryInfoKey(handle)[1]
    for i in range(num_values):
        for x in winreg.EnumValue(handle, i):
            if(str(x).find("EXCEL") > -1):
                EXCELEXE = x

def getWord():
    global WORDEXE
    handle = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\winword.exe")

    num_values = winreg.QueryInfoKey(handle)[1]
    for i in range(num_values):
        for x in winreg.EnumValue(handle, i):
            if(str(x).find("WINWORD") > -1):
                WORDEXE = x

def writeAddCheckBat(fout):
    fout.write("@echo off\n")
    fout.write("REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n")
    fout.write("REM  This file checks for the yahrzeit download from shulcloud  ;;\n")
    fout.write("REM  then, either prompts the user to download it, or reformats ;;\n")
    fout.write("REM  the file to make it printable and usable by the gabbai     ;;\n")
    fout.write("REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n")
    #fout.write("cd c:\Users\KTM\python\msutils\CheckEntry")
    fout.write("cd " + basedir + "\\CheckEntry\n")
    fout.write("\n")
    fout.write(":getfilename\n")
    fout.write("echo processing 'accounts.xlsx'\n")
    fout.write("if EXIST accounts.xlsx (\n")
    fout.write("echo found accounts.xlsx\n")
    fout.write("goto getfilename2\n")
    fout.write("\n")
    fout.write("\n")
    fout.write("echo You need to create an accounts file...\n")
    fout.write("type .\\readme.me\n")
    fout.write("goto enderror\n")
    fout.write("\n")
    fout.write(":getfilename2\n")
    fout.write("echo processing 'checks.xlsx'\n")
    fout.write("if EXIST checks.xlsx (\n")
    fout.write("echo found checks.xlsx\n")
    fout.write(".\\dist\\CheckEntry\\CheckEntry\n")
    fout.write("goto endall\n")
    fout.write("\n")
    fout.write("\n")
    fout.write("echo You need to create a checks file...\n")
    fout.write("type .\\readme.me\n")
    fout.write("goto enderror\n")
    fout.write("\n")
    fout.write("echo yahrzeits.xlsx file does not exist\n")
    fout.write("goto enderror\n")
    fout.write("\n")
    fout.write(":endall\n")
    fout.write("\n")
    fout.write("goto finalexit\n")
    fout.write(":enderror\n")
    fout.write("echo nope! try again...\n")
    fout.write("\n")
    fout.write(":finalexit\n")
    fout.write("echo thanks!\n")
    fout.write("\n")


def writeYahrBat(fout):
    fout.write("@echo off\n")
    fout.write("REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n")
    fout.write("REM  This file checks for the yahrzeit download from shulcloud  ;;\n")
    fout.write("REM  then, either prompts the user to download it, or reformats ;;\n")
    fout.write("REM  the file to make it printable and usable by the gabbai     ;;\n")
    fout.write("REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n")
    fout.write("cd " + basedir + "\n")
    print("Base Directory: " + basedir)
    fout.write("del new.xlsx\n")
    fout.write("\n")
    fout.write("echo checking for file...\n")
    fout.write("if \"%~1\"==\"\" GOTO getfilename\n")
    fout.write("echo .\\yahr\\dist\\yahr\\yahr \"%~1\"\n")
    fout.write(".\\yahr\\dist\\yahr\\yahr \"%~1\"\n")
    fout.write("goto endall\n")
    fout.write("\n")
    fout.write(":getfilename\n")
    fout.write("echo no filename given, processing 'yahrzeits.xlsx'\n")
    fout.write("if EXIST yahrzeits.xlsx (\n")
    fout.write(".\\yahr\\dist\\yahr\\yahr yahrzeits.xlsx\n")
    fout.write("goto endall\n")
    fout.write(")\n")
    fout.write("\n")
    fout.write("echo yahrzeits.xlsx file does not exist\n")
    fout.write("goto enderror\n")
    fout.write("\n")
    fout.write(":endall\n")
    fout.write("\"" + EXCELEXE + "\" new.xlsx\n")
    fout.write("\n")
    fout.write(":enderror\n")
    fout.write("echo nope! try again...\n")

def writeReconcile001Bat(fout):
    fout.write("@echo off\n")
    fout.write("REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n")
    fout.write("REM  This file checks for the yahrzeit download from shulcloud  ;;\n")
    fout.write("REM  then, either prompts the user to download it, or reformats ;;\n")
    fout.write("REM  the file to make it printable and usable by the gabbai     ;;\n")
    fout.write("REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n")
    fout.write("cd " + basedir + "\n")
    print("Base Directory: " + basedir)
    fout.write("del .\\Reconcile\\out001\n")
    fout.write("\n")
    fout.write(":getfilename1\n")
    fout.write("echo for PayPal processing 'PPTrans.xlsx'\n")
    fout.write("if EXIST .\\Reconcile\\PPTrans.xlsx (\n")
    fout.write("goto getfilename2\n")
    fout.write(")\n")
    fout.write("\n")
    fout.write("type .\\Reconcile\\ReadMe.Me\n")
    fout.write("goto enderror\n")
    fout.write("\n")
    fout.write(":getfilename2\n")
    fout.write("echo for ShulCloud processing 'SCTrans.xlsx'\n")
    fout.write("if EXIST .\\Reconcile\\SCTrans.xlsx (\n")
    #fout.write("python .\\Reconcile\\Recon001.py > .\\Reconcile\\out001\n")
    fout.write(".\\Reconcile\\dist\\Recon001 > .\\Reconcile\\out001\n")
    fout.write("goto endall\n")
    fout.write(")\n")
    fout.write("\n")
    fout.write("type .\\Reconcile\\ReadMe.Me\n")
    fout.write("goto enderror\n")
    fout.write("\n")
    fout.write(":endall\n")
    fout.write("\"" + WORDEXE + "\" .\\Reconcile\\out001\n")
    fout.write("\n")
    fout.write(":enderror\n")
    fout.write("echo nope! try again...\n")

def writeReconcile002Bat(fout):
    fout.write("@echo off\n")
    fout.write("REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n")
    fout.write("REM  This file checks for the yahrzeit download from shulcloud  ;;\n")
    fout.write("REM  then, either prompts the user to download it, or reformats ;;\n")
    fout.write("REM  the file to make it printable and usable by the gabbai     ;;\n")
    fout.write("REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n")
    fout.write("cd " + basedir + "\n")
    print("Base Directory: " + basedir)
    fout.write("del .\\Reconcile\\out002\n")
    fout.write("\n")
    fout.write(":getfilename1\n")
    fout.write("echo for PayPal processing 'PeoplePP.xlsx'\n")
    fout.write("if EXIST .\\Reconcile\\PeoplePP.xlsx (\n")
    fout.write("goto getfilename2\n")
    fout.write(")\n")
    fout.write("\n")
    fout.write("type .\\Reconcile\\ReadMe.Me\n")
    fout.write("goto enderror\n")
    fout.write("\n")
    fout.write(":getfilename2\n")
    fout.write("echo for ShulCloud processing 'PeopleSC.xlsx'\n")
    fout.write("if EXIST .\\Reconcile\\PeopleSC.xlsx (\n")
    #fout.write("python .\\Reconcile\\Recon002.py > .\\Reconcile\\out002\n")
    fout.write(".\\Reconcile\\dist\\Recon002 > .\\Reconcile\\out002\n")
    fout.write("goto endall\n")
    fout.write(")\n")
    fout.write("\n")
    fout.write("type .\\Reconcile\\ReadMe.Me\n")
    fout.write("goto enderror\n")
    fout.write("\n")
    fout.write(":endall\n")
    fout.write("\"" + WORDEXE + "\" .\\Reconcile\\out002\n")
    fout.write("\n")
    fout.write(":enderror\n")
    fout.write("echo nope! try again...\n")

getExcel()
getWord()
f = open("..\\bat\\yahr.bat", "w")
writeYahrBat(f)
f.close()

f = open("..\\bat\\addChecks.bat", "w")
writeAddCheckBat(f)
f.close()

f = open("..\\bat\\recon001.bat", "w")
writeReconcile001Bat(f)
f.close()

f = open("..\\bat\\recon002.bat", "w")
writeReconcile002Bat(f)
f.close()
