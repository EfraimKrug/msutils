from CDCommonCode import *
from checkDisplay02 import *
####################################################################################################
### This is where the entire application opens - with a list of all deposits tracked so far -
### in all workbooks in the directory
####################################################################################################
class checkDisplay01:
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
        self.master.geometry('400x300')
        self.master.title('Kadima Toras-Moshe Check Tracking')

        self.frame = tk.Frame(self.master, width=360, height=260)
        self.frame.configure(bg="teal", pady=2, padx=2)
        self.frame.grid(row=1, column=1)

        self.people = []
        self.pages = []
        self.files = []
        self.workbooks = dict()
        self.workingFile = 'DailyLog'

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

    def getSheet(self, name, sheet, wb):
        (day, month) = self.CDCommonCode.parseName(name)
        if not name in self.pages:
            self.pages.append([name, wb])

        self.depositName.append([str(sheet.cell(row=2,column=7).value), name, wb])

        for r in range(3, sheet.max_row):
            if(str(sheet.cell(row=r,column=1).value).lower() == 'cash'):
                self.cashcheckSwitch = 'cash'
            if(str(sheet.cell(row=r,column=1).value).lower().find('check') > -1):
                self.cashcheckSwitch = 'check'

            if(sheet.cell(row=r, column=2).value and self.cashcheckSwitch.find('check') > -1):
                self.loadRow(month, day, sheet, r)

            if(sheet.cell(row=r, column=2).value and self.cashcheckSwitch.find('cash') > -1):
                self.loadRowCash(month, day, sheet, r)


# on change dropdown value
    # def change_dropdown(self, *args):
    #     self.workingFile = self.tkvar.get()


    #
    # get display for specific excel page...
    #
    #def change_dropdown2(self, *args):
    def change_dropdown2(self, event):
        #print("change_dropdown2")
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
            total += line[5]
        total = "{:.2f}".format(float(total))
        total = "Cash: $" + total
        self.label07 = tk.Label(self.frame, text=total, bg="teal", fg="yellow", font='Helvetica 10 bold')
        self.label07.grid(row=1, column=18, padx=4, pady=4, sticky=tk.NW)

    def showPerson(self, name, args):
        self.newWindow = tk.Toplevel(self.master)
        self.app = checkDisplay03(self.newWindow, name)

    def showDeposit(self, name, args):
        wb = ""
        dName = ""
        sName = ""
        for depArr in self.depositName:
            if name == depArr[0]:
                dName = depArr[0]
                sName = depArr[1]
                wb = depArr[2]

        self.newWindow = tk.Toplevel(self.master)
        #self.app = checkDisplay04(self.newWindow, dName, sName, wb)
        self.app = checkDisplay04(self.newWindow, dName, '', wb)

    # link function to change change_dropdown
    def showData(self):
        total = 0
        self.label03 = []
        #self.button01 = []

        self.peepFunctions = []
        self.deposits = []

        fileNames = []

        self.files = self.CDCommonCode.getFiles(self.files)
        fileList = self.files
        # self.tkvar = tk.StringVar(self.master)
        # self.tkvar.trace('w', self.change_dropdown)
        # self.tkvar.set(fileList[0]) # set the default option
        #

        # pagesPopup = tk.OptionMenu(self.frame, self.tkvar, *fileList)
        # pagesPopup.grid(row = 1, column =6, padx=10, pady=10, sticky=tk.EW)
        pages = []
        for a in self.pages:
            pages.append(a[0])

        pages = sorted(pages, key=self.CDCommonCode.compareMonths)
###################################################################################
        # try sorting pages better...
        tempDict = dict()
        mth = []
        #print(pages)
        for nm in pages:
            fnm = nm[:-2]
            num = nm[-2:]
            if fnm not in mth:
                tempDict[fnm] = []
                mth.append(fnm)

            tempDict[fnm].append(num)

        pages = []
        mth = sorted(mth, key=self.CDCommonCode.compareMonths)
        # print(mth)
        for nm in mth:
            tempDict[nm].sort()

        for nm in mth:
            for num in tempDict[nm]:
                x = str(nm) + str(num)
                if x not in pages:
                    pages.append(str(nm) + str(num))




###################################################################################

        #pages = sorted(pages)

        self.tkvar2 = tk.StringVar()
        # self.tkvar2 = tk.StringVar(self.master)
        # self.tkvar2.trace('w', self.change_dropdown2)
        # self.tkvar2.set(pages[len(pages)-1]) # set the default option
        #
        # pagesPopup2 = tk.OptionMenu(self.frame, self.tkvar2, *pages)
        # pagesPopup2.grid(row = 1, column =5, padx=10, pady=10, sticky=tk.EW)
        # print(pages)
        pagesPopup2 = ttk.Combobox(self.frame, textvariable=self.tkvar2, values=pages)
        pagesPopup2.grid(row = 1, column =5, padx=10, pady=10, sticky=tk.EW)
        pagesPopup2.current(1)
        pagesPopup2.bind("<<ComboboxSelected>>", self.change_dropdown2)

        #self.button01 = tk.Button(self.frame, text="Shift", command=partial(self.CDCommonCode.shiftWBook, self.files, self.workingFile))
        self.button01 = tk.Button(self.frame, text="Shift", command=partial(self.CDCommonCode.cleanUp, self.files))
        self.button01.grid(row=1, column=2, columnspan=1, padx=10, pady=10, sticky=tk.EW)

        self.button02 = tk.Button(self.frame, text="Open Excel", command=partial(self.CDCommonCode.openNewSheet, self.workingFile))
        self.button02.grid(row=1, column=3, columnspan=1, padx=10, pady=10, sticky=tk.EW)

        row_num = 6
        self.headline03 = tk.Label(self.frame, text=" Deposits ", bg="teal", fg="yellow")
        self.headline03.grid(row=3, column=3, padx=4, pady=2, sticky=tk.W)

        for e1 in self.ds:
            for e2 in self.ds[e1]:
                if not e2[1] in self.deposits:
                    self.deposits.append(e2[1])

        for d in self.deposits:
            self.label03.append(tk.Label(self.frame, text=d, bg="teal", fg="yellow"))
            self.label03[len(self.label03)-1].grid(row=row_num, column=3, padx=4, pady=4, sticky=tk.NW)
            self.label03[len(self.label03)-1].bind("<Button-1>", partial(self.showDeposit, d))
            row_num = row_num + 1

    # def sheetExists(self, sheet):
    #     dailyLog = self.CDCommonCode.openDailyLog(self.files)
    #     for name in dailyLog.sheetnames:
    #         if name == sheet:
    #             return True
    #     return False

    def runProcess(self):
        self.files = self.CDCommonCode.getFiles(self.files)
        self.workbooks = self.CDCommonCode.openDailyLog(self.files)
        for wb in self.workbooks:
            for name in self.workbooks[wb].sheetnames:
                self.total = 0
                self.getSheet(name, self.workbooks[wb][name], wb)
        self.showData()
