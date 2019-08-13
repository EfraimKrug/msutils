from CDCommonCode import *
from checkDisplay04 import *
####################################################################################################
### This is where the entire application opens - with a list of all deposits tracked so far -
### in all workbooks in the directory
####################################################################################################
class DepositDisplay01:
    def __init__(self, master):
        self.cashcheckSwitch = ''
        self.ds = dict()        # {check_name: [check_number, memo, check_date, arrival_date, check_amount, check_image],
        self.sdata = dict()     # sheet by sheet...
        self.pdata = dict()
        self.cdata = []     # cash
        #self.searchObj = ''
        self.depositName = []

        self.master = master
        self.master.configure(bg="teal", pady=34, padx=17)
        self.master.geometry('700x700')
        self.master.title('Kadima Toras-Moshe Deposit Tracking')

        self.frame = tk.Frame(self.master, width=700, height=700)
        self.frame.configure(bg="teal", pady=2, padx=2)
        self.frame.grid(row=1, column=1)

        self.allHistory = True

        self.people = []
        self.pages = []
        self.files = []
        self.workbooks = dict()
        self.workingFile = 'Deposits'
        self.sheetNames = []

        self.dropInit = True
        self.CDCommonCode = CDCommonCode(self.master)
        self.runProcess()

    def getSheet(self, sheet):
        arr = ['a', 'b', 'c', 'd', 'e', 'f', 'g']
        arrptr = 0
        prev = ""
        curr = ""
        prevDate = ""
        currDate = ""

        for r in range(5, sheet.max_row+1):
            if (sheet.cell(row=r, column=2).value):
                curr = sheet.cell(row=r, column=2).value
                currDate = sheet.cell(row=r, column=4).value
                prev = curr
                prevDate = currDate
                arrptr = 0
            else:
                curr = prev + "@@" + str(arr[arrptr])
                currDate = prevDate
                arrptr += 1

            # print(prev + "/" + curr)
            self.ds[curr] = dict()
            self.ds[curr]['row'] = r
            # self.ds[curr]['date'] = sheet.cell(row=r, column=4).value
            self.ds[curr]['date'] = currDate
            self.ds[curr]['cash'] = sheet.cell(row=r, column=5).value
            self.ds[curr]['checks'] = sheet.cell(row=r, column=6).value
            self.ds[curr]['total'] = sheet.cell(row=r, column=7).value
            self.ds[curr]['image'] = sheet.cell(row=r, column=9).value
            # print(self.ds[sheet.cell(row=r, column=2).value]['image'])

    def depositExists(self, sheet, depositName):
        # print(depositName, sheet.title)
        for r in range(5, sheet.max_row+1):
            if str(depositName).find(str(sheet.cell(row=r, column=2).value)) > -1:
                return True
        return False

    def getDepositNextLine(self, sheet):
        return sheet.max_row + 1

    #
    # get display for specific excel page...
    #
    #def change_dropdown2(self, *args):
    def change_dropdown2(self, event):
        wb = ""
        name = self.tkvar2.get()
        for a in self.pages:
            if a[0] == name:
                wb = a[1]

        self.newWindow = tk.Toplevel(self.master)
        self.app = checkDisplay02(self.newWindow, name, wb, self.CDCommonCode)

    def showCash(self):
        total = 0
        for line in self.cdata:
            total += float(line[5])
        total = "{:.2f}".format(float(total))
        total = "Cash: $" + total
        self.label07 = tk.Label(self.frame, text=total, bg="teal", fg="yellow", font='Helvetica 10 bold')
        self.label07.grid(row=1, column=18, padx=4, pady=4, sticky=tk.NW)

    def showPerson(self, name, args):
        self.newWindow = tk.Toplevel(self.master)
        self.app = checkDisplay03(self.newWindow, name, self.CDCommonCode)

    def showDeposit(self, name, args):
        wb = ""
        dName = name
        sName = ""
        self.newWindow = tk.Toplevel(self.master)
        self.app = checkDisplay04(self.newWindow, dName, '', wb, self.CDCommonCode)

    # link function to change change_dropdown
    def createDeposit(self):
        sheetList = self.getCheckFiles(self.getDepositFromCheckFiles())
        newDepName = self.accumulateDeposit(self.workbooks['DepositList'], sheetList)
        vals = self.getFinancialTotals(newDepName[0])
        self.updateDeposit(self.workbooks['DepositList'], vals, newDepName)
        recordArray = self.buildTrackingSheet(newDepName)
        self.createNewSheet(self.workbooks, newDepName, recordArray)
        self.workbooks.save(self.workingFile + ".xlsx")

    def mySort(self, inDict):
        arrToSort = []
        outDict1 = dict()

        for line in inDict:
            arrToSort.append(str(inDict[line]['date']) + "_$_" + line)

        arrToSort = sorted(arrToSort)
        for e in arrToSort:
            sp = e.split('_$_')

        print(arrToSort)

        targetVal = -1
        outDict = dict()
        return inDict

    def showData(self):
        label00 = []
        row_num = 2

        sheetList = self.getCheckFiles(self.getDepositFromCheckFiles())
        newDepName = self.accumulateDeposit(self.workbooks['DepositList'], sheetList)

        if newDepName[0]:
            buttonText = "Create: " + newDepName[0]
            button01 = tk.Button(self.frame, text=buttonText, command=partial(self.createDeposit))
            button01.grid(row=1, column=2, columnspan=3, padx=20, pady=10, sticky=tk.EW)

        label00.append(tk.Label(self.frame, text='Deposit', bg="teal", fg="blue"))
        label00[len(label00)-1].grid(row=row_num, column=1, padx=4, pady=4, sticky=tk.NW)

        label00.append(tk.Label(self.frame, text='Date', bg="teal", fg="blue"))
        label00[len(label00)-1].grid(row=row_num, column=3, padx=4, pady=4, sticky=tk.NW)

        label00.append(tk.Label(self.frame, text='Cash', bg="teal", fg="blue"))
        label00[len(label00)-1].grid(row=row_num, column=5, padx=4, pady=4, sticky=tk.NW)

        label00.append(tk.Label(self.frame, text='Checks', bg="teal", fg="blue"))
        label00[len(label00)-1].grid(row=row_num, column=7, padx=4, pady=4, sticky=tk.NW)

        label00.append(tk.Label(self.frame, text='Total', bg="teal", fg="blue"))
        label00[len(label00)-1].grid(row=row_num, column=9, padx=4, pady=4, sticky=tk.NW)

        label00.append(tk.Label(self.frame, text='Image', bg="teal", fg="blue"))
        label00[len(label00)-1].grid(row=row_num, column=11, padx=4, pady=4, sticky=tk.NW)

        row_num = 3
        total_lines = len(self.ds.keys())
        myThing = self.mySort(self.ds)

        # for line in self.ds.keys():
        for line in sorted(myThing.keys()):
            if (total_lines - row_num > 15):
                row_num += 1
                continue

            lab = ''
            if '@@' in line:
                label = ''
            else:
                lab = line

            label00.append(tk.Label(self.frame, text=lab, bg="teal", fg="yellow"))
            label00[len(label00)-1].grid(row=row_num, column=1, padx=4, pady=4, sticky=tk.NW)
            label00[len(label00)-1].bind("<Button-1>", partial(self.showDeposit, line))

            label00.append(tk.Label(self.frame, text=str(myThing[line]['date'])[0:10], bg="teal", fg="yellow"))
            label00[len(label00)-1].grid(row=row_num, column=3, padx=4, pady=4, sticky=tk.NW)

            x = ""
            if(myThing[line]['cash']):
                x = "${:0,.2f}".format(myThing[line]['cash'])
            label00.append(tk.Label(self.frame, text=x, bg="teal", fg="yellow"))
            label00[len(label00)-1].grid(row=row_num, column=5, padx=4, pady=4, sticky=tk.NW)

            x = ""
            if(myThing[line]['checks']):
                x = "${:0,.2f}".format(myThing[line]['checks'])
            label00.append(tk.Label(self.frame, text=x, bg="teal", fg="yellow"))
            label00[len(label00)-1].grid(row=row_num, column=7, padx=4, pady=4, sticky=tk.NW)

            x = ""
            if(myThing[line]['total']):
                x = "${:0,.2f}".format(myThing[line]['total'])
            label00.append(tk.Label(self.frame, text=x, bg="teal", fg="yellow"))
            label00[len(label00)-1].grid(row=row_num, column=9, padx=4, pady=4, sticky=tk.NW)

            label00.append(tk.Label(self.frame, text=myThing[line]['image'], bg="teal", fg="yellow"))
            label00[len(label00)-1].grid(row=row_num, column=11, padx=4, pady=4, sticky=tk.NW)
            label00[len(label00)-1].bind("<Button-1>", partial(self.CDCommonCode.showDepositImage, myThing[line]['image']))

            row_num+=1



    #
    # looks through all the check files, and compares the deposit names
    # to the depositList file. If the name is not there, it is returned.
    #
    def getDepositFromCheckFiles(self):
        files = []
        self.CDCommonCode.setAllHistory(self.allHistory)
        files = self.CDCommonCode.getFiles(files)
        workbooks = self.CDCommonCode.openDailyLog(files)
        for wb in workbooks:
            for name in workbooks[wb].sheetnames:
                if(not self.depositExists(self.workbooks['DepositList'], workbooks[wb][name].cell(row=2, column=7).value)):
                    return workbooks[wb][name].cell(row=2, column=7).value

    def getFinancialTotals(self, depName):
        files = []
        vals = dict()
        vals['cash'] = 0
        vals['checks'] = 0

        files = self.CDCommonCode.getFiles(files)
        workbooks = self.CDCommonCode.openDailyLog(files)
        for wb in workbooks:
            for name in workbooks[wb].sheetnames:
                if(workbooks[wb][name].cell(row=2, column=7).value == depName):
                    for r in range(5, workbooks[wb][name].max_row):
                        if (str(workbooks[wb][name].cell(row=r, column=4).value).find('Total:') > -1):
                            if(workbooks[wb][name].cell(row=r, column=5).value):
                                vals['cash'] = vals['cash'] + float(workbooks[wb][name].cell(row=r, column=5).value)
                        if (str(workbooks[wb][name].cell(row=r, column=3).value).find('Sub Total:') > -1):
                            if(workbooks[wb][name].cell(row=r, column=5).value):
                                vals['checks'] = vals['checks'] + float(workbooks[wb][name].cell(row=r, column=5).value)

        return vals

    ########################################################################
    #   @param records is a dictionary returned from
    #   Return a dictionary with two entries: Cash and checks
    #   Each entry is an array of Records
    ########################################################################
    def createHeadLines(self, newSheet, lineNum):
        if lineNum == 0:
            newSheet.cell(row=1, column=2).value = "Name on Check"
            newSheet.cell(row=1, column=3).value = "Memo"
            newSheet.cell(row=1, column=4).value = "Date on Check"
            newSheet.cell(row=1, column=5).value = "Amount"
            newSheet.cell(row=1, column=6).value = "Image"
            newSheet.cell(row=2, column=1).value = "Check No."
            return

        newSheet.cell(row=lineNum, column=1).value = "Cash"

    def createNewSheet(self, wb, depName, records):
        counter = 3
        checkTotal = 0
        cashTotal = 0
        newSheet = wb.create_sheet(title = depName[0])
        self.createHeadLines(newSheet, counter - 3)
        #for type in records:
        for line in records['checks']:

            if(records['checks'][line][0] == 'None' and records['checks'][line][1] == 'None'):
                continue
            if(records['checks'][line][3] == 'None' and records['checks'][line][4] == 'None'):
                continue
            checkTotal += float(records['checks'][line][4])
            for arraySpot in range(0, 6):
                cellSpot = arraySpot + 1
                if(arraySpot == 4):
                    newSheet.cell(row=counter, column=cellSpot).value = float(records['checks'][line][arraySpot])
                else:
                    if(arraySpot == 3):
                        newSheet.cell(row=counter, column=cellSpot).value = records['checks'][line][arraySpot][0:10]
                    else:
                        newSheet.cell(row=counter, column=cellSpot).value = str(records['checks'][line][arraySpot])

            counter += 1

        counter += 2
        newSheet.cell(row=counter, column=4).value = "Checks: "
        newSheet.cell(row=counter, column=5).value = checkTotal
        counter += 3

        self.createHeadLines(newSheet, counter)
        counter += 1
        for line in records['cash']:
            if(records['cash'][line][0] == 'None' and records['cash'][line][1] == 'None'):
                continue
            if(records['cash'][line][3] == 'None' and records['cash'][line][4] == 'None'):
                continue
            cashTotal += float(records['cash'][line][4])
            for arraySpot in range(0, 6):
                cellSpot = arraySpot + 1
                if(arraySpot == 4):
                    newSheet.cell(row=counter, column=cellSpot).value = float(records['cash'][line][arraySpot])
                else:
                    if(arraySpot == 3):
                        newSheet.cell(row=counter, column=cellSpot).value = records['cash'][line][arraySpot][0:10]
                    else:
                        newSheet.cell(row=counter, column=cellSpot).value = str(records['cash'][line][arraySpot])

            counter += 1

        counter += 3
        newSheet.cell(row=counter, column=4).value = "Cash: "
        newSheet.cell(row=counter, column=5).value = cashTotal
        newSheet.cell(row=counter+1, column=4).value = "Total: "
        newSheet.cell(row=counter+1, column=5).value = cashTotal + checkTotal

    #######################################################################
    # This function looks at all the sheets in the workbooks
    # If the sheet is assigned to our deposit (depName)
    #   Then get all the Cash Records,
    #   And get all the Check Records
    #
    #   Return a dictionary with two entries: Cash and checks
    #   Each entry is an array of Records
    ########################################################################
    def checkCashSwitch(self, switch):
        if switch.find('cash') > -1:
            return True
        return False

    def buildTrackingSheet(self, depName):
        recordArray = dict()
        recordArray['cash'] = dict()
        recordArray['checks'] = dict()
        switch = 'cash'
        counter = 0

        files = []
        files = self.CDCommonCode.getFiles(files)
        workbooks = self.CDCommonCode.openDailyLog(files)
        pageGuard = []

        for wb in workbooks:
            for name in workbooks[wb].sheetnames:
                if(workbooks[wb][name].cell(row=2, column=7).value == depName[0]):
                    switch = 'cash'
                    if workbooks[wb][name].title in pageGuard:
                        continue
                    pageGuard.append(workbooks[wb][name].title)

                    for r in range(3, workbooks[wb][name].max_row):
                        if (str(workbooks[wb][name].cell(row=r, column=1).value).find('Check') > -1):
                            switch = 'check'

                        if (self.checkCashSwitch(switch)):
                            recordArray['cash'][counter] = []
                            for f in range(1,7):
                                recordArray['cash'][counter].append(str(workbooks[wb][name].cell(row=r, column=f).value))
                        else:
                             recordArray['checks'][counter] = []
                             for f in range(1,7):
                                 recordArray['checks'][counter].append(str(workbooks[wb][name].cell(row=r, column=f).value))
                        counter += 1
        return recordArray

    def getCheckFiles(self, depositName):
        files = []
        sheets = []
        files = self.CDCommonCode.getFiles(files)
        workbooks = self.CDCommonCode.openDailyLog(files)
        for wb in workbooks:
            for name in workbooks[wb].sheetnames:
                if(workbooks[wb][name].cell(row=2, column=7).value == depositName):
                    sheets.append(workbooks[wb][name])

        return sheets

    #
    # this updates the deposit list?
    def accumulateDeposit(self, depSheet, checkList):
        ret = []
        newDepName = self.getDepositFromCheckFiles()
        ret.append(newDepName)
        nextLine = self.getDepositNextLine(depSheet)
        ret.append(nextLine)
        depSheet.cell(row=nextLine, column=2).value = newDepName
        depSheet.cell(row=nextLine, column=4).value = datetime.today().strftime('%m/%d/%Y')
        return ret

    def updateDeposit(self, depSheet, vals, newDepName):
        depSheet.cell(row=newDepName[1], column=5).value = vals['cash']
        depSheet.cell(row=newDepName[1], column=6).value = vals['checks']
        depSheet.cell(row=newDepName[1], column=7).value = vals['cash'] + vals['checks']

    def runProcess(self):
        self.workbooks = self.CDCommonCode.openDepositLog(self.workingFile)
        for sheet in self.workbooks:
            self.sheetNames.append(sheet.title)

        if 'DepositList' in self.sheetNames:
            self.getSheet(self.workbooks['DepositList'])

        self.showData()
