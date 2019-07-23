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

        self.people = []
        self.pages = []
        self.files = []
        self.workbooks = dict()
        self.workingFile = 'Deposits'
        self.sheetNames = []

        #self.tkvar = ''
        #self.EXCELEXE = ''

        self.dropInit = True
        self.CDCommonCode = CDCommonCode(self.master)
        self.runProcess()


    def loadRow(self, month, day, sheet, current_row):
        arr = []
        if(sheet.cell(row=current_row, column=2).value in self.ds):
            arr = self.ds[sheet.cell(row=current_row, column=2).value]


        name = month+day

        dName = ''
        for depArr in self.depositName:
            if name == depArr[1]:
                dName = depArr[0]

        peep = sheet.cell(row=current_row, column=2).value

        newRow = [sheet.cell(row=current_row, column=1).value,
                  dName,
                  sheet.cell(row=current_row, column=3).value,
                  str(sheet.cell(row=current_row, column=4).value)[0:10],
                  str(month) + "-" + str(day),
                  sheet.cell(row=current_row, column=5).value,
                  sheet.cell(row=current_row, column=6).value,
                  name]

        if not peep in self.people:
            self.people.append(peep)

        arr.append(newRow)
        self.ds[sheet.cell(row=current_row, column=2).value] = arr
        if not name in self.sdata:
            self.sdata[name] = dict()

        self.sdata[name][sheet.cell(row=current_row, column=2).value] = arr
        if not peep in self.pdata:
            self.pdata[peep] = dict()

        self.pdata[peep][sheet.cell(row=current_row, column=2).value] = arr


    def loadRowCash(self, month, day, sheet, current_row):
        arr = []
        if(sheet.cell(row=current_row, column=2).value in self.ds):
            arr = self.ds[sheet.cell(row=current_row, column=2).value]

        name = month+day

        dName = ''
        for depArr in self.depositName:
            if name == depArr[1]:
                dName = depArr[0]

        peep = sheet.cell(row=current_row, column=2).value

        newRow = [sheet.cell(row=current_row, column=1).value,
                  dName,
                  sheet.cell(row=current_row, column=3).value,
                  str(sheet.cell(row=current_row, column=4).value)[0:10],
                  str(month) + "-" + str(day),
                  sheet.cell(row=current_row, column=5).value,
                  sheet.cell(row=current_row, column=6).value,
                  name]

        self.cdata.append(newRow)

    def getSheet(self, sheet):
        for r in range(5, sheet.max_row+1):
            self.ds[sheet.cell(row=r, column=2).value] = dict()
            self.ds[sheet.cell(row=r, column=2).value]['row'] = r
            self.ds[sheet.cell(row=r, column=2).value]['date'] = sheet.cell(row=r, column=4).value
            self.ds[sheet.cell(row=r, column=2).value]['cash'] = sheet.cell(row=r, column=5).value
            self.ds[sheet.cell(row=r, column=2).value]['checks'] = sheet.cell(row=r, column=6).value
            self.ds[sheet.cell(row=r, column=2).value]['total'] = sheet.cell(row=r, column=7).value
            self.ds[sheet.cell(row=r, column=2).value]['image'] = sheet.cell(row=r, column=9).value

    def depositExists(self, sheet, depositName):
        for r in range(5, sheet.max_row):
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
        # print("change_dropdown2")
        # if self.dropInit:
        #     self.dropInit = False
        #     return
        wb = ""
        name = self.tkvar2.get()
        for a in self.pages:
            if a[0] == name:
                wb = a[1]

        self.newWindow = tk.Toplevel(self.master)
        self.app = checkDisplay02(self.newWindow, name, wb)

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
        self.app = checkDisplay03(self.newWindow, name)

    def showDeposit(self, name, args):
        #self.CDCommonCode.showDepositImage(name)
        print (name)

        wb = ""
        dName = name
        sName = ""
        # for depArr in self.depositName:
        #     print(depArr)
        #     if name == depArr[0]:
        #         dName = depArr[0]
        #         sName = depArr[1]
        #         wb = depArr[2]
        #
        self.newWindow = tk.Toplevel(self.master)
        #self.app = checkDisplay04(self.newWindow, dName, sName, wb)
        # self.app = checkDisplay04(self.newWindow, dName, '', wb)
        self.app = checkDisplay04(self.newWindow, dName, '', wb)

    # link function to change change_dropdown
    def createDeposit(self):
        # print("creating deposit")
        sheetList = self.getCheckFiles(self.getDepositFromCheckFiles())
        newDepName = self.accumulateDeposit(self.workbooks['DepositList'], sheetList)
        vals = self.getFinancialTotals(newDepName[0])
        self.updateDeposit(self.workbooks['DepositList'], vals, newDepName)
        recordArray = self.buildTrackingSheet(newDepName)
        self.createNewSheet(self.workbooks, newDepName, recordArray)
        # print("Saving: " + self.workingFile)
        self.workbooks.save(self.workingFile + ".xlsx")


    # def mySort(self, inDict):
    #     targetVal = -1
    #     outDict = dict()
    #     keysHold1 = inDict.keys()
    #     keysHold2 = keysHold1
    #     for line in keysHold1:
    #         for line1 in keysHold2:
    #             if inDict[line1]['row'] == targetVal:
    #                 outDict[line1] = inDict[line]
    #         targetVal += 1
    #     return outDict

    def showData(self):
        label00 = []
        row_num = 2

        sheetList = self.getCheckFiles(self.getDepositFromCheckFiles())
        newDepName = self.accumulateDeposit(self.workbooks['DepositList'], sheetList)

        #button01 = tk.Button(self.frame, text="Create", command=partial(self.createDeposit, self.files, self.workingFile))
        if newDepName[0]:
            buttonText = "Create: " + newDepName[0]
            button01 = tk.Button(self.frame, text=buttonText, command=partial(self.createDeposit))
            button01.grid(row=1, column=2, columnspan=3, padx=20, pady=10, sticky=tk.EW)

        # entryBox = tk.Entry(self.frame)
        # entryBox.insert(0, newDepName[0])
        # entryBox.grid(row=1, column=3, columnspan=1, padx=10, pady=10, sticky=tk.EW)

        label00.append(tk.Label(self.frame, text='Deposit', bg="teal", fg="yellow"))
        label00[len(label00)-1].grid(row=row_num, column=1, padx=4, pady=4, sticky=tk.NW)

        label00.append(tk.Label(self.frame, text='Date', bg="teal", fg="yellow"))
        label00[len(label00)-1].grid(row=row_num, column=3, padx=4, pady=4, sticky=tk.NW)

        label00.append(tk.Label(self.frame, text='Cash', bg="teal", fg="yellow"))
        label00[len(label00)-1].grid(row=row_num, column=5, padx=4, pady=4, sticky=tk.NW)

        label00.append(tk.Label(self.frame, text='Checks', bg="teal", fg="yellow"))
        label00[len(label00)-1].grid(row=row_num, column=7, padx=4, pady=4, sticky=tk.NW)

        label00.append(tk.Label(self.frame, text='Total', bg="teal", fg="yellow"))
        label00[len(label00)-1].grid(row=row_num, column=9, padx=4, pady=4, sticky=tk.NW)

        label00.append(tk.Label(self.frame, text='Image', bg="teal", fg="yellow"))
        label00[len(label00)-1].grid(row=row_num, column=11, padx=4, pady=4, sticky=tk.NW)

        row_num = 3
        total_lines = len(self.ds.keys())
        # myThing = self.mySort(self.ds)
        for line in self.ds.keys():
            if (total_lines - row_num > 7):
                row_num += 1
                continue

            label00.append(tk.Label(self.frame, text=line, bg="teal", fg="yellow"))
            label00[len(label00)-1].grid(row=row_num, column=1, padx=4, pady=4, sticky=tk.NW)
            label00[len(label00)-1].bind("<Button-1>", partial(self.showDeposit, line))

            label00.append(tk.Label(self.frame, text=str(self.ds[line]['date'])[0:10], bg="teal", fg="yellow"))
            label00[len(label00)-1].grid(row=row_num, column=3, padx=4, pady=4, sticky=tk.NW)

            #x = "${:0,.2f}".format(self.ds[line]['cash'])
            x = ""
            if(self.ds[line]['cash']):
                x = "${:0,.2f}".format(self.ds[line]['cash'])
            label00.append(tk.Label(self.frame, text=x, bg="teal", fg="yellow"))
            label00[len(label00)-1].grid(row=row_num, column=5, padx=4, pady=4, sticky=tk.NW)

            x = ""
            if(self.ds[line]['checks']):
                x = "${:0,.2f}".format(self.ds[line]['checks'])
            label00.append(tk.Label(self.frame, text=x, bg="teal", fg="yellow"))
            label00[len(label00)-1].grid(row=row_num, column=7, padx=4, pady=4, sticky=tk.NW)

            x = ""
            if(self.ds[line]['total']):
                x = "${:0,.2f}".format(self.ds[line]['total'])
            label00.append(tk.Label(self.frame, text=x, bg="teal", fg="yellow"))
            label00[len(label00)-1].grid(row=row_num, column=9, padx=4, pady=4, sticky=tk.NW)

            label00.append(tk.Label(self.frame, text=self.ds[line]['image'], bg="teal", fg="yellow"))
            label00[len(label00)-1].grid(row=row_num, column=11, padx=4, pady=4, sticky=tk.NW)
            label00[len(label00)-1].bind("<Button-1>", partial(self.CDCommonCode.showDepositImage, self.ds[line]['image']))

            row_num+=1



    #
    # looks through all the check files, and compares the deposit names
    # to the depositList file. If the name is not there, it is returned.
    #
    def getDepositFromCheckFiles(self):
        #print("getDepositFromCheckFiles")
        files = []
        files = self.CDCommonCode.getFiles(files)
        workbooks = self.CDCommonCode.openDailyLog(files)
        for wb in workbooks:
            for name in workbooks[wb].sheetnames:
                if(not self.depositExists(self.workbooks['DepositList'], workbooks[wb][name].cell(row=2, column=7).value)):
                    return workbooks[wb][name].cell(row=2, column=7).value
                #print(name)

    def getFinancialTotals(self, depName):
        files = []
        vals = dict()
        vals['cash'] = 0
        vals['checks'] = 0

        files = self.CDCommonCode.getFiles(files)
        workbooks = self.CDCommonCode.openDailyLog(files)
        # print("++>>")
        for wb in workbooks:
            for name in workbooks[wb].sheetnames:
                # print(name + "==>" +  workbooks[wb][name].cell(row=2, column=7).value)
                if(workbooks[wb][name].cell(row=2, column=7).value == depName):
                    for r in range(5, workbooks[wb][name].max_row):
                        # print(str(workbooks[wb][name].cell(row=r, column=4).value))
                        if (str(workbooks[wb][name].cell(row=r, column=4).value).find('Total:') > -1):
                            if(workbooks[wb][name].cell(row=r, column=5).value):
                                # print(workbooks[wb][name].cell(row=r, column=5).value)
                                vals['cash'] = vals['cash'] + float(workbooks[wb][name].cell(row=r, column=5).value)
                                # print("cash")
                                # print(int(workbooks[wb][name].cell(row=r, column=5).value))
                        if (str(workbooks[wb][name].cell(row=r, column=3).value).find('Sub Total:') > -1):
                            if(workbooks[wb][name].cell(row=r, column=5).value):
                                # print(workbooks[wb][name].cell(row=r, column=5).value)
                                vals['checks'] = vals['checks'] + float(workbooks[wb][name].cell(row=r, column=5).value)
                                # print("Check")
                                # print(int(workbooks[wb][name].cell(row=r, column=5).value))

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
        #print(depName)
        newSheet = wb.create_sheet(title = depName[0])
        self.createHeadLines(newSheet, counter - 3)
        #for type in records:
        for line in records['checks']:

            if(records['checks'][line][0] == 'None' and records['checks'][line][1] == 'None'):
                continue
            if(records['checks'][line][3] == 'None' and records['checks'][line][4] == 'None'):
                continue
            checkTotal += float(records['checks'][line][4])
            # print(checkTotal)
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
                #print("HERE: " + workbooks[wb][name].cell(row=2, column=7).value)
                #print(str(workbooks[wb][name].cell(row=2, column=7).value) + ":" + depName[0])
                if(workbooks[wb][name].cell(row=2, column=7).value == depName[0]):
                    switch = 'cash'
                    if workbooks[wb][name].title in pageGuard:
                        continue
                    pageGuard.append(workbooks[wb][name].title)

                    for r in range(3, workbooks[wb][name].max_row):
                        #print(r)
                        if (str(workbooks[wb][name].cell(row=r, column=1).value).find('Check') > -1):
                            switch = 'check'

                        if (self.checkCashSwitch(switch)):
                            #print("CASH")
                            recordArray['cash'][counter] = []
                            for f in range(1,7):
                                recordArray['cash'][counter].append(str(workbooks[wb][name].cell(row=r, column=f).value))
                                #print(recordArray['cash'][counter])
                            #counter += 1
                        else:
                             #print("CHECKS")
                             recordArray['checks'][counter] = []
                             for f in range(1,7):
                                 recordArray['checks'][counter].append(str(workbooks[wb][name].cell(row=r, column=f).value))
                        counter += 1
        #print("HERE")
        #print(recordArray)
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
        # self.files = self.CDCommonCode.openDepositLog(self.files)
        self.workbooks = self.CDCommonCode.openDepositLog(self.workingFile)
        for sheet in self.workbooks:
            self.sheetNames.append(sheet.title)

        if 'DepositList' in self.sheetNames:
            self.getSheet(self.workbooks['DepositList'])

        # if self.depositExists(self.workbooks['DepositList'], 'YadaYada'):
        #     print("YadaYada Exists")
        # else:
        #     print("YadaYada Doesn't Exist")
        #
        # if self.depositExists(self.workbooks['DepositList'], 'Acharei Mot'):
        #     print("Kedoshim Exists")
        # else:
        #     print("Kedoshim Doesn't Exist")

        #print(self.getDepositNextLine(self.workbooks['DepositList']))
        # print(self.getDepositFromCheckFiles())
        # sheetList = self.getCheckFiles(self.getDepositFromCheckFiles())
        # # for sheet in sheetList:
        # #     print("--> " + sheet.title)
        # newDepName = self.accumulateDeposit(self.workbooks['DepositList'], sheetList)
        # vals = self.getFinancialTotals(newDepName[0])
        # self.updateDeposit(self.workbooks['DepositList'], vals, newDepName)
        # recordArray = self.buildTrackingSheet(newDepName)
        # #for lineNum in recordArray['cash']:
        # #    print(recordArray['cash'][lineNum])
        # self.createNewSheet(self.workbooks, newDepName, recordArray)
        # print("Saving: " + self.workingFile)
        # self.workbooks.save(self.workingFile + ".xlsx")
        #print(self.ds)
        # for name in self.workbooks[wb].sheetnames:
        #     self.total = 0
        #     #self.getSheet(name, self.workbooks[wb][name], wb)
        #     print("==>" + name)
        self.showData()
