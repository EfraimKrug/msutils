import os
import subprocess
import signal

cmd = 'WMIC PROCESS get Caption,Processid'
proc = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE)
processArray = []
for line in proc.stdout:
    holdArr = []
    arr = line.split(' ')
    for a in arr:
        if(a.strip() != ''):
            holdArr.append(a)
    processArray.append(holdArr)

for entry in processArray:
    if(len(entry) < 2): continue
    if(entry[0].find('MSPUB') > -1 or
        entry[0].find('atom') > -1 or
        entry[0].find('Browser.exe') > -1 or
        entry[0].find('Dropbox') > -1 or
        entry[0].find('EXCEL') > -1 or
        entry[0].find('Skype') > -1 or
        entry[0].find('OneDrive') > -1 or
        entry[0].find('LMS') > -1 or
        entry[0].find('UNS') > -1 or
        entry[0].find('WORD') > -1):
        try:
            subprocess.check_output("Taskkill /PID %d /F" % int(entry[1]))
            print(entry[0] + "==>" + entry[1])
        except:
            print("Could not kill " + entry[0])
    else:
        print(entry[0])
