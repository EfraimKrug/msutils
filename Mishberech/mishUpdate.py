#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
import csv
from datetime import datetime
from datetime import time
from datetime import date

from openpyxl import load_workbook
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl.utils import get_column_letter

from openpyxl.styles.borders import Border, Side

import smtplib
from profile import *

#fromaddr = 'KadimahTorasMoshe@gmail.com'
toaddrs  = 'EfraimMKrug@gmail.com'

####################################################
#username = 'KadimahTorasMoshe@gmail.com'
#password = 'August7Brachas'
####################################################

last = 0
accountArray = []
accountArray.append(dict())

def openTrx():
    wb = load_workbook('MishBerech.xlsx')
    return wb

def getTrx(sheet):
    global last
    for r in range(7, sheet.max_row + 1):
        if(type(sheet.cell(row=r, column=6).value) != unicode):
            continue
        #print(type(sheet.cell(row=r, column=6).value))
        if sheet.cell(row=r, column=6).value.find("@") > -1:
            accountArray.append(dict())
            last += 1

        accountArray[last][r] = []
        accountArray[last][r].append(sheet.cell(row=r, column=2).value)  #name
        accountArray[last][r].append(unicode(sheet.cell(row=r, column=3).value))  #choleh
        x = sheet.cell(row=r, column=3).value
        y = x.encode('UTF-8')
        print(y.decode('UTF-8'))
        accountArray[last][r].append(sheet.cell(row=r, column=4).value)  #date
        accountArray[last][r].append(sheet.cell(row=r, column=5).value)  #comment
        accountArray[last][r].append(sheet.cell(row=r, column=6).value)  #email

def getEmail():
    today_dt = datetime.now()

    for accounts in accountArray:
        for id in accounts:
            #for r in range(2, sheet.max_row + 1):
                print("Name: " + (accounts[id][1]))
                line_dt = datetime.strptime(str(accounts[id][2])[0:10], "%Y-%m-%d")
                dayCount = (today_dt - line_dt).days
                if dayCount > 45:
                    #print(writeEmail(accounts[id]))
                    print ("> 45")
                else:
                    #print ("< 45 so no email..." + str(accounts[id][2]))
                    print ("< 45")

def writeEmail(account):
    s = ""
    s = "Dear " + account[0] + ",\n\n"
    s += "We have been saying a mish'berech for " + account[1].encode("utf8")
    s += "\nsince " + str(account[2])[0:10] + ".\n\n"
    s += "Please email us if you would like us to continue mentioning this name in the shul"
    s += "\nmish'berech, otherwise we will be removing the name after this coming Shabbos."
    s += "\n\nThanks so much!\n\nBest and Good Shabbos! "
    #print (s)
    return s

def sendTo(account):
    arr = []
    arr2 = []
    for email in account[3]:
        if(isinstance(email, unicode) == True):
            if (email.find("@") > 0):
                arr.append(email)

    for email in arr:
        if (email not in arr2):
            arr2.append(email)
    return arr2

def sendItAll():
    global server
    global fromaddr
    global toaddrs

    usedAddressList = []
    for accounts in accountArray:
        for a in accounts:
            msgTxt = writeEmail(accounts[a])
            list = sendTo(accounts[a])
            usedAddressList = []
            if(len(list) < 1):
                print("Not Sent: " + str(a))
            for email in list:
                toaddrs = email
                if email not in usedAddressList:
                    msg = "\r\n".join([
                      "From: " + fromaddr,
                      "To: " + toaddrs,
                      "Subject: MishBerech list",
                      "",
                      msgTxt
                      ])

                    #server.sendmail(fromaddr, toaddrs, msg)
                    print("Sending to: " + toaddrs)
                    usedAddressList.append(email)
                print("Not sending to: " + toaddrs)

#######################################################

trx = openTrx()
getTrx(trx[trx.sheetnames[0]])
getEmail()
#getEmail(ppl[ppl.sheetnames[0]])
#server = smtplib.SMTP(smtpvar)
#server.ehlo()
#server.starttls()
#server.login(username,password)
#sendItAll()
#server.quit()
