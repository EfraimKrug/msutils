from CDCommonCode import *
from checkDisplay02 import *
import tkinter.font as tkFont
####################################################################################################
### Showing the individual deposit - that will probably span across spreadsheets -
####################################################################################################
class checkDisplay04:
    def __init__(self, master, depositName, sheet, wb, CDCommonCode):
        self.cashcheckSwitch = ''
        self.ds = dict()
        self.depositName = depositName
        self.sheetName = sheet
        self.wb = wb
        self.sdata = dict()     # sheet by sheet...
        self.pdata = dict()
        self.cdata = []

        self.master = master
        self.fullHeight = self.master.winfo_screenheight()
        self.fullWidth = self.master.winfo_screenwidth()

        self.master.configure(bg="teal", pady=3, padx=1)
        self.master.geometry('530x700')
        self.master.title(depositName)


        self.frame = tk.Frame(self.master, width=310, height=self.fullHeight)
        self.frame.configure(bg="teal", pady=2, padx=2)
        self.frame.grid(row=1, column=1)

        self.people = []
        self.pages = []
        # self.tkvar = ''
        self.checkTotal = 0
        self.cashTotal = 0

        self.depositWB = ""
        self.newDepositSheet = ""
        self.CDCommonCode = CDCommonCode
        self.runProcess(self.depositName)

    ########################################################################################
    # Gathering all the cash data for a new row in our new spread sheet
    # The data is placed into and array,
    # which is placed in the array: self.cdata
    # @param month, day: date from the work sheet we are working with...
    # @param sheet: name of the sheet - for convenience (actually the same as month-day)
    # @param current_row: the row we are reading from...
    # @sideEffect: loading the data onto self.cdata (cash data)
    ########################################################################################
    def loadRowCash(self, month, day, sheet, current_row):
        arr = []
        if(sheet.cell(row=current_row, column=2).value in self.ds):
            arr = self.ds[sheet.cell(row=current_row, column=2).value]

        name = month+day
        peep = sheet.cell(row=current_row, column=2).value

        newRow = [sheet.cell(row=current_row, column=1).value,
                  sheet.cell(row=current_row, column=3).value,
                  str(sheet.cell(row=current_row, column=4).value)[0:10],
                  str(month) + "-" + str(day),
                  sheet.cell(row=current_row, column=5).value,
                  sheet.cell(row=current_row, column=6).value,
                  name]

        self.cdata.append(newRow)

    ########################################################################################
    # Gathering all the check data for a new row in our new spread sheet
    # The data is placed into and array,
    # which is placed in the array: self.cdata
    # @param month, day: date from the work sheet we are working with...
    # @param sheet: name of the sheet - for convenience (actually the same as month-day)
    # @param current_row: the row we are reading from...
    # @sideEffect: loading the data onto self.cdata (cash data)
    ########################################################################################
    def loadRow(self, month, day, sheet, current_row):
        # new array for each row
        arr = []
        # if we have seen this name before, get the data we have already gathered
        if(sheet.cell(row=current_row, column=2).value in self.ds):
            arr = self.ds[sheet.cell(row=current_row, column=2).value]

        name = month+day
        # person's name (on check)
        peep = sheet.cell(row=current_row, column=2).value

        newRow = [sheet.cell(row=current_row, column=1).value,
                  sheet.cell(row=current_row, column=3).value,
                  str(sheet.cell(row=current_row, column=4).value)[0:10],
                  str(month) + "-" + str(day),
                  sheet.cell(row=current_row, column=5).value,
                  sheet.cell(row=current_row, column=6).value,
                  name]
        # gather a list of all the unique people names
        if not peep in self.people:
            self.people.append(peep)

        # this array will now be all the data for this name we have seen so far
        arr.append(newRow)

        # ds is a dictionary key: person's name/value list of arrays one array / row
        self.ds[sheet.cell(row=current_row, column=2).value] = arr
        if not name in self.sdata:
            self.sdata[name] = dict()

        self.sdata[name][sheet.cell(row=current_row, column=2).value] = arr
        if not peep in self.pdata:
            self.pdata[peep] = dict()

        self.pdata[peep][sheet.cell(row=current_row, column=2).value] = arr


    def getSheet(self, name, sheet, depositName):

        if(not sheet.cell(row=2, column=7).value == depositName):
            return False

        (day, month) = self.CDCommonCode.parseName(name)
        if not name in self.pages:
            self.pages.append(name)

        for r in range(3, sheet.max_row):
            if(str(sheet.cell(row=r,column=1).value).lower() == 'cash'):
                self.cashcheckSwitch = 'cash'
            if(str(sheet.cell(row=r,column=1).value).lower().find('check') > -1):
                self.cashcheckSwitch = 'check'

            if(sheet.cell(row=r, column=2).value and self.cashcheckSwitch.find('check') > -1):
                self.loadRow(month, day, sheet, r)

            if(sheet.cell(row=r, column=2).value and self.cashcheckSwitch.find('cash') > -1):
                self.loadRowCash(month, day, sheet, r)

        return True

    def showPerson(self, name, args):
        self.newWindow = tk.Toplevel(self.master)
        #self.app = checkDisplay03(self.newWindow, name, '', self.wb)
        self.app = checkDisplay03(self.newWindow, name, self.CDCommonCode)

    def showCash(self):
        total = 0
        for line in self.cdata:
            total += line[4]

        self.cashTotal = total
        total = "{:.2f}".format(float(total))
        total = "Cash: $" + total
        self.label07 = tk.Label(self.frame, text=total, bg="teal", fg="yellow", font='Helvetica 10 bold')
        self.label07.grid(row=1, column=18, padx=4, pady=4, sticky=tk.NW)

    def openDepositWorkBook(self, depositName):
        self.depositWB = load_workbook(depositDir + '\\Deposits.xlsx')
        for name in self.depositWB.sheetnames:
            if name == depositName:
                return False

        self.newDepositSheet = self.depositWB.create_sheet(title = depositName)
        return True

    def copyData(self, oldSheet, newSheet):
            newCheckRow = 3
            newCashRow = 50

            newSheet.cell(row=1, column=2).value = "Name on check"
            newSheet.cell(row=1, column=3).value = "Memo"
            newSheet.cell(row=1, column=4).value = "Date on check"
            newSheet.cell(row=1, column=5).value = "Amount"
            newSheet.cell(row=1, column=6).value = "Image"
            newSheet.cell(row=2, column=1).value = "Check No."
            newSheet.cell(row=49, column=1).value = "Cash"

            for r in range(3, oldSheet.max_row):
                if(str(oldSheet.cell(row=r,column=1).value).lower() == 'cash'):
                    self.cashcheckSwitch = 'cash'
                if(str(oldSheet.cell(row=r,column=1).value).lower().find('check') > -1):
                    self.cashcheckSwitch = 'check'

                if(oldSheet.cell(row=r, column=2).value and self.cashcheckSwitch.find('check') > -1):
                    for c in range(1, 8):
                        newSheet.cell(row=newCheckRow, column=c).value = oldSheet.cell(row=r,column=c).value
                    newCheckRow += 1

                if(oldSheet.cell(row=r, column=2).value and self.cashcheckSwitch.find('cash') > -1):
                    for c in range(1, 8):
                        newSheet.cell(row=newCashRow, column=c).value = oldSheet.cell(row=r,column=c).value
                    newCashRow += 1

    def makeDeposit(self, name, args):
        if(self.openDepositWorkBook(name)):
            dailyLog = self.CDCommonCode.openOneDailyLog(self.wb)
            for sheetName in dailyLog.sheetnames:
                if(name == dailyLog[sheetName].cell(row=2,column=7).value):
                    self.copyData(dailyLog[sheetName], self.newDepositSheet)

        self.depositWB.save(depositDir + '\\Deposits.xlsx')

    def showData(self, rNum):
        row_num = rNum
        self.label01 = []
        self.label02 = []
        self.label03 = []
        self.label04 = []
        self.label05 = []
        self.label06 = []
        self.button01 = []
        fileNames = []

        self.title = tk.Label(self.frame, text=self.depositName, bg="teal", fg="blue", font='Helvetica 10 bold')
        self.title.grid(row=1, column=1, padx=1, pady=4, sticky=tk.NW)
        self.title.bind("<Button-1>", partial(self.makeDeposit, self.depositName))

        self.headline01 = tk.Label(self.frame, text="Name", bg="teal", fg="blue")
        self.headline01.grid(row=3, column=2, padx=1, pady=2, sticky=tk.W)

        self.headline02 = tk.Label(self.frame, text=" Date ", bg="teal", fg="blue")
        self.headline02.grid(row=3, column=4, padx=1, pady=2, sticky=tk.W)

        self.headline03 = tk.Label(self.frame, text="Check #", bg="teal", fg="blue")
        self.headline03.grid(row=3, column=6, padx=1, pady=2, sticky=tk.W)

        self.headline04 = tk.Label(self.frame, text="Amount", bg="teal", fg="blue")
        self.headline04.grid(row=3, column=8, padx=1, pady=2, sticky=tk.W)

        self.headline07 = tk.Label(self.frame, text="Image", bg="teal", fg="blue")
        self.headline07.grid(row=3, column=14, padx=1, pady=2, sticky=tk.W)


        sortedKeys = []
        for key in self.ds:
            sortedKeys.append(key)
        sortedKeys.sort()
        lastName = ''
        pEnt = ''

        self.master.geometry('530x700')
        if len(sortedKeys) > 15:
            self.master.geometry('530x%s' % (self.fullHeight))

        self.showCash()
        label_font = tkFont.Font(family='Arial', size=8)

        for ent in sortedKeys:
            for e in self.ds[ent]:
                pEnt = ent
                if lastName == ent:
                    pEnt = ''
                lastName = ent
                self.label01.append(tk.Label(self.frame, text=pEnt, font=label_font, bg="teal", fg="yellow"))
                self.label01[len(self.label01)-1].grid(row=row_num, column=2, padx=1, pady=1, sticky=tk.NW)
                self.label01[len(self.label01)-1].bind("<Button-1>", partial(self.showPerson, pEnt))

                self.label02.append(tk.Label(self.frame, text=e[3], font=label_font, bg="teal", fg="yellow"))
                self.label02[len(self.label02)-1].grid(row=row_num, column=4, padx=1, pady=1, sticky=tk.NW)

                self.label03.append(tk.Label(self.frame, text=e[0], font=label_font, bg="teal", fg="yellow"))
                self.label03[len(self.label03)-1].grid(row=row_num, column=6, padx=1, pady=1, sticky=tk.NW)

                fAmt = "{:.2f}".format(float(e[4]))
                self.label05.append(tk.Label(self.frame, text=fAmt, font=label_font, bg="teal", fg="yellow"))
                self.label05[len(self.label04)-1].grid(row=row_num, column=8, padx=1, pady=1, sticky=tk.NW)

                self.checkTotal = self.checkTotal + e[4]

                self.button01.append(tk.Button(self.frame, text="View", font=label_font, command=partial(self.CDCommonCode.show_image, e[5]),height=1))
                self.button01[len(self.button01)-1].grid(row=row_num, column=14, columnspan=2, padx=1, pady=1, sticky=tk.EW)

                row_num += 1
                line = ''
        return row_num

    def runProcess(self, depositName):
        fileList = self.CDCommonCode.getFiles([])
        workbooks = self.CDCommonCode.openDailyLog(fileList)
        for wb in workbooks:
            for name in workbooks[wb].sheetnames:
                self.total = 0
                self.getSheet(name, workbooks[wb][name], depositName)
        ###########################################################
        # dailyLog = self.CDCommonCode.openOneDailyLog(self.wb)
        self.total = 0
        row_num = 6
        # for name in dailyLog.sheetnames:
        #     self.getSheet(name, dailyLog[name], depositName)

        row_num = self.showData(row_num)
        row_num = row_num + 3

        fAmt = "Checks: $" + "{:.2f}".format(float(self.checkTotal))
        fAmtLabel = tk.Label(self.frame, text=fAmt, bg="teal", fg="yellow", font='Helvetica 10 bold')
        fAmtLabel.grid(row=row_num, column=18, padx=4, pady=4, sticky=tk.NW)

        row_num = row_num + 1
        fAmt = "Total: $" + "{:.2f}".format(float(self.checkTotal + self.cashTotal))
        fAmtLabel2 = tk.Label(self.frame, text=fAmt, bg="teal", fg="yellow", font='Helvetica 10 bold')
        fAmtLabel2.grid(row=row_num, column=18, padx=4, pady=4, sticky=tk.NW)
