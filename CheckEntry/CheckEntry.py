
#CheckProcess
import sys
import webbrowser
import csv
from openpyxl import load_workbook

ACCOUNTS = dict()
CHECKS = dict()
CHROME_PATH = 'C:/Users/KTM/AppData/Local/Google/Chrome/Application/chrome.exe %s'
URLS = []

def openwb():
  wb = load_workbook('./Accounts.xlsx')
  return wb

def loadACCOUNTS(wbook):
    _key = ""
    _key2 = ""
    _val = ""

    sheet = wbook[wbook.sheetnames[0]]
    for r in range(1, sheet.max_row+1):
        _key = sheet.cell(row=r,column=1).value[0:str(sheet.cell(row=r,column=1).value).find(',')]
        init = str(sheet.cell(row=r,column=1).value).find(',') + 1
        _key2 = sheet.cell(row=r,column=1).value[init:]
        _key2 = _key2.replace(' ','')

        key2_idx = 0
        while _key in ACCOUNTS:
            _key = _key + _key2[key2_idx]
            key2_idx = key2_idx + 1

        ACCOUNTS[_key] = sheet.cell(row=r,column=2).value

def loadURLS(checks):
    _key = ""
    _key2 = ""
    _val = ""

    sheet = checks[checks.sheetnames[0]]
    for r in range(1, sheet.max_row+1):
        init = str(sheet.cell(row=r,column=1).value).find(',') + 1
        if init < 1:
            _key = sheet.cell(row=r,column=1).value
        else:
            _key = sheet.cell(row=r,column=1).value[0:init-1]
        _key2 = sheet.cell(row=r,column=1).value[init:]
        _key2 = _key2.replace(' ','')

        key2_idx = 0
        while _key in ACCOUNTS:
            _prev_key = _key
            _key = _key + _key2[key2_idx]
            key2_idx = key2_idx + 1

        _key = _prev_key
        if _key in ACCOUNTS:
            print ("KEY: " + _key + "," + _key2)
            ID = ACCOUNTS[_key]
            URLS.append("https://www.kadimahtorasmoshe.org/admin/transaction_add.php?account_id=" + str(ID) + "&return=account")

        print("==================================")

def printURLS():
    for line in URLS:
        print(line)

def processURLS():
    for url in URLS:
        print(url)
        webbrowser.get(CHROME_PATH).open(url)
        print("Press enter to continue: ")
        x = raw_input()

def openChecksWB():
  wb = load_workbook('./Checks.xlsx')
  return wb

def printFile(wbook):
    sheet = wbook[wbook.sheetnames[0]]
    for r in range(1, sheet.max_row+1):
        print(sheet.cell(row=r,column=1).value + " :: " + str(sheet.cell(row=r,column=2).value))

accounts = openwb()
checks = openChecksWB()

loadACCOUNTS(accounts)
loadURLS(checks)
processURLS()
