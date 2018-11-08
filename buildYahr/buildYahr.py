#
# buildYahr
#   get directories
#   build yahr.bat
#
import os
from profile import *
import winreg

#EXCELEXE = r'C:\Program Files\Microsoft Office\root\Office\EXCEL.EXE'

def getExcel():
    global EXCELEXE
    handle = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")

    num_values = winreg.QueryInfoKey(handle)[1]
    for i in range(num_values):
        for x in winreg.EnumValue(handle, i):
            if(str(x).find("EXCEL") > -1):
                EXCELEXE = x

def writeBat(fout):
    fout.write("@echo off\n")
    fout.write("REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n")
    fout.write("REM  This file checks for the yahrzeit download from shulcloud  ;;\n")
    fout.write("REM  then, either prompts the user to download it, or reformats ;;\n")
    fout.write("REM  the file to make it printable and usable by the gabbai     ;;\n")
    fout.write("REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;\n")
    fout.write("cd " + basedir + "\n")
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

getExcel()
f = open("..\\bat\\yahr.bat", "w")
writeBat(f)
f.close()
