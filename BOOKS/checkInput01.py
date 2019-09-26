from CDCommonCode import *

class checkInput01:
    def __init__(self, master):
        self.depositName = []
        self.workingFile = 'DailyLog'
        self.master = master
        self.master.configure(bg="teal", pady=34, padx=17)

        # else:
        self.master.geometry('400x400')
        self.frame = tk.Frame(self.master, width=360, height=360)

        self.master.title('Kadima Toras-Moshe Check Input')

        self.frame.configure(bg="teal", pady=2, padx=2)
        self.frame.grid(row=1, column=1)

        # self.people = []
        self.pages = []
        self.files = []
        self.workbooks = dict()
        self.workingFile = 'DailyLog'

        # self.dropInit = True
        self.CDCommonCode = CDCommonCode(self.master)
        self.runProcess()

    def getSheet(self, name, sheet, wb):
        (day, month) = self.CDCommonCode.parseName(name)
        if not name in self.pages:
            self.pages.append([name, wb])

        self.depositName.append([str(sheet.cell(row=2,column=7).value), name, wb])

    def getNextRow(self, newSheet, rowNumber):
        r = rowNumber
        while not newSheet.cell(row=r, column=5).value is None:
            r += 1
            if r > 200:
                break
        return r

    def updateCash(self, newSheet, dataWidgets):
        newSheet.cell(row=2, column=7).value = dataWidgets['depositName'].get()
        nextRow = self.getNextRow(newSheet, 3)
        newSheet.cell(row=nextRow, column=1).value = "Cash"
        newSheet.cell(row=nextRow, column=2).value = dataWidgets['checkName'].get()
        newSheet.cell(row=nextRow, column=3).value = dataWidgets['notes'].get()
        newSheet.cell(row=nextRow, column=4).value = dataWidgets['checkDate'].get()
        newSheet.cell(row=nextRow, column=5).value = float(dataWidgets['amount'].get())
        newSheet.cell(row=11,column=5).value = "=SUM(E3:E10)"
        newSheet.cell(row=31,column=5).value = "=SUM(E13:E30)"
        newSheet.cell(row=32,column=5).value = "=SUM(E11,E31)"

    def updateChecks(self, newSheet, dataWidgets):
        newSheet.cell(row=2, column=7).value = dataWidgets['depositName'].get()
        nextRow = self.getNextRow(newSheet, 13)
        newSheet.cell(row=nextRow, column=1).value = dataWidgets['checkNumber'].get()
        newSheet.cell(row=nextRow, column=2).value = dataWidgets['checkName'].get()
        newSheet.cell(row=nextRow, column=3).value = dataWidgets['notes'].get()
        newSheet.cell(row=nextRow, column=4).value = dataWidgets['checkDate'].get()
        newSheet.cell(row=nextRow, column=5).value = float(dataWidgets['amount'].get())
        newSheet.cell(row=11,column=5).value = "=SUM(E3:E10)"
        newSheet.cell(row=31,column=5).value = "=SUM(E13:E30)"
        newSheet.cell(row=32,column=5).value = "=SUM(E11,E31)"

    def change_dropdown2(self, event):
        wb = ""
        name = self.tkvar2.get()
        # print(name)

    def clearValues(self, dataWidgets):
        dataWidgets['checkNumber'].delete(0, tk.END)
        dataWidgets['checkNumber'].insert(0, "Cash")
        dataWidgets['checkName'].delete(0, tk.END)
        dataWidgets['checkName'].insert(0, "")
        dataWidgets['notes'].delete(0, tk.END)
        dataWidgets['notes'].insert(0, "")
        dataWidgets['amount'].delete(0, tk.END)
        dataWidgets['amount'].insert(0, "0.00")


    def saveData(self, dataWidgets):

        self.CDCommonCode.setWorkbooks(self.workbooks)
        self.CDCommonCode.createSheet(self.workingFile)

        newSheet = self.CDCommonCode.getCurrentWorkbook(self.workingFile)[self.CDCommonCode.getNewSheetName()]
        if dataWidgets['checkNumber'].get() in "Cash":
            self.updateCash(newSheet, dataWidgets)
        else:
            self.updateChecks(newSheet, dataWidgets)

        self.clearValues(dataWidgets)
        self.CDCommonCode.saveWorkbook(self.workingFile)

    def showData(self, lastDeposit):
        total = 0
        self.label03 = []

        row_num = 1
        dataWidgets = dict()

        self.files = self.CDCommonCode.getFiles(self.files)
        fileList = self.files
        self.tkvar2 = tk.StringVar()
        pages = ["New", lastDeposit]

        self.label03.append(tk.Label(self.frame, text="Deposit", bg="teal", fg="yellow"))
        self.label03[len(self.label03)-1].grid(row=row_num, column=1, padx=4, pady=4, sticky=tk.NW)

        depositName = ttk.Entry(self.frame)
        depositName.insert(0, lastDeposit)
        depositName.grid(row = row_num, column=2, padx=10, pady=10, sticky=tk.EW)
        dataWidgets['depositName'] = depositName

        # check number
        row_num = row_num + 1
        self.label03.append(tk.Label(self.frame, text="Check Number", bg="teal", fg="yellow"))
        self.label03[len(self.label03)-1].grid(row=row_num, column=1, padx=4, pady=4, sticky=tk.NW)
        # row_num = row_num + 1

        checkNumber = ttk.Entry(self.frame)
        checkNumber.grid(row = row_num, column=2, padx=10, pady=10, sticky=tk.EW)
        checkNumber.insert(0, "Cash")
        dataWidgets['checkNumber'] = checkNumber

        # name on check
        row_num = row_num + 1
        self.label03.append(tk.Label(self.frame, text="Last Name", bg="teal", fg="yellow"))
        self.label03[len(self.label03)-1].grid(row=row_num, column=1, padx=4, pady=4, sticky=tk.NW)
        # row_num = row_num + 1
        checkName = ttk.Entry(self.frame)
        checkName.grid(row = row_num, column=2, padx=10, pady=10, sticky=tk.EW)
        dataWidgets['checkName'] = checkName

        # memo/notes
        row_num = row_num + 1
        self.label03.append(tk.Label(self.frame, text="Notes", bg="teal", fg="yellow"))
        self.label03[len(self.label03)-1].grid(row=row_num, column=1, padx=4, pady=4, sticky=tk.NW)
        # row_num = row_num + 1
        notes = ttk.Entry(self.frame)
        notes.grid(row = row_num, column=2, padx=10, pady=10, sticky=tk.EW)
        dataWidgets['notes'] = notes

        # date
        row_num = row_num + 1
        self.label03.append(tk.Label(self.frame, text="Date", bg="teal", fg="yellow"))
        self.label03[len(self.label03)-1].grid(row=row_num, column=1, padx=4, pady=4, sticky=tk.NW)
        # row_num = row_num + 1
        checkDate = ttk.Entry(self.frame)
        checkDate.insert(0, datetime.today().strftime('%Y-%m-%d'))
        checkDate.grid(row = row_num, column=2, padx=10, pady=10, sticky=tk.EW)
        dataWidgets['checkDate'] = checkDate

        # amount
        row_num = row_num + 1
        self.label03.append(tk.Label(self.frame, text="Amount", bg="teal", fg="yellow"))
        self.label03[len(self.label03)-1].grid(row=row_num, column=1, padx=4, pady=4, sticky=tk.NW)
        # row_num = row_num + 1
        amount = ttk.Entry(self.frame)
        amount.insert(0, "0.00")
        amount.grid(row = row_num, column=2, padx=10, pady=10, sticky=tk.EW)
        dataWidgets['amount'] = amount

        row_num = row_num + 7
        self.button01 = tk.Button(self.frame, text="Save", command=partial(self.saveData, dataWidgets))
        self.button01.grid(row=row_num, column=2, columnspan=1, padx=10, pady=10, sticky=tk.EW)

    def runProcess(self):
        self.files = self.CDCommonCode.getFiles(self.files)
        self.workbooks = self.CDCommonCode.openDailyLog(self.files)
        for wb in self.workbooks:
            for name in self.workbooks[wb].sheetnames:
                self.total = 0
                self.getSheet(name, self.workbooks[wb][name], wb)

        lastDeposit = self.depositName[len(self.depositName)-1][0]
        self.showData(lastDeposit)
