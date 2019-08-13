from CDCommonCode import *
from checkDisplay03 import *
from checkDisplay04 import *
#import PyPDF2
####################################################################################################
class checkDisplay02:
    def __init__(self, master, oSheetName, oWBName, CDCommonCode):
        self.cashcheckSwitch = ''
        self.ds = dict()
        self.oSheetName = oSheetName
        self.oWBName = oWBName
        self.sdata = dict()     # sheet by sheet...
        self.pdata = dict()
        self.cdata = []
        self.depositName = ''

        self.master = master
        self.master.configure(bg="teal", pady=34, padx=17)
        self.master.geometry('700x700')
        self.master.title('Edit Your Display Data')

        self.frame = tk.Frame(self.master, width=460, height=360)
        self.frame.configure(bg="teal", pady=2, padx=2)
        self.frame.grid(row=1, column=1)

        self.people = []
        self.pages = []
        self.tkvar = ''

        # self.CDCommonCode = CDCommonCode(self.master)
        self.CDCommonCode = CDCommonCode
        self.runProcess(self.oSheetName)

    def loadRowCash(self, month, day, sheet, current_row):
        arr = []
        if(sheet.cell(row=current_row, column=2).value in self.ds):
            arr = self.ds[sheet.cell(row=current_row, column=2).value]

        name = month+day
        peep = sheet.cell(row=current_row, column=2).value

        newRow = [sheet.cell(row=current_row, column=1).value,
                  self.depositName,
                  sheet.cell(row=current_row, column=3).value,
                  str(sheet.cell(row=current_row, column=4).value)[0:10],
                  str(month) + "-" + str(day),
                  sheet.cell(row=current_row, column=5).value,
                  sheet.cell(row=current_row, column=6).value,
                  name]

        self.cdata.append(newRow)


    def loadRow(self, month, day, sheet, current_row):
        arr = []
        if(sheet.cell(row=current_row, column=2).value in self.ds):
            arr = self.ds[sheet.cell(row=current_row, column=2).value]

        name = month+day
        peep = sheet.cell(row=current_row, column=2).value

        newRow = [sheet.cell(row=current_row, column=1).value,
                  self.depositName,
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


    def getSheet(self, name, sheet):
        (day, month) = self.CDCommonCode.parseName(name)
        if not name in self.pages:
            self.pages.append(name)

        self.depositName = str(sheet.cell(row=2,column=7).value)

        for r in range(3, sheet.max_row):
            if(str(sheet.cell(row=r,column=1).value).lower() == 'cash'):
                self.cashcheckSwitch = 'cash'
            if(str(sheet.cell(row=r,column=1).value).lower().find('check') > -1):
                self.cashcheckSwitch = 'check'

            if(sheet.cell(row=r, column=2).value and self.cashcheckSwitch.find('check') > -1):
                self.loadRow(month, day, sheet, r)

            if(sheet.cell(row=r, column=2).value and self.cashcheckSwitch.find('cash') > -1):
                self.loadRowCash(month, day, sheet, r)


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
        self.app = checkDisplay03(self.newWindow, name, self.CDCommonCode)

    def showDeposit(self, name, args):
        self.newWindow = tk.Toplevel(self.master)
        self.app = checkDisplay04(self.newWindow, name, '', self.oWBName, self.CDCommonCode)

    def printSheet(self, sheetName, args):
        fPath = dailyLogDir + "\\\\print.xlsx"
        try:
            os.remove(fPath)
        except:
            print("could not find " + fPath)

        dailyLog = self.CDCommonCode.openOneDailyLog(self.oWBName)
        daySheet = dailyLog[sheetName]
        wb = Workbook()
        #printSheet = wb.create_sheet(title = 'printSheet')
        printSheet = wb['Sheet']
        printSheet.column_dimensions['C'].width = 32
        printSheet.column_dimensions['D'].width = 20
        printSheet.column_dimensions['F'].width = 20
        printSheet.page_setup.orientation = printSheet.ORIENTATION_LANDSCAPE
        printSheet.page_setup.fitToWidth = True

        for r in range(1, daySheet.max_row + 1):
            for c in range(1, 15):
                printSheet.cell(row=r,column=c).value = daySheet.cell(row=r,column=c).value

        wb.save(fPath)
        #print("saving: " + fPath)
        os.startfile(fPath, "print")

        # Theoretically this would print the check images... but it's not
        # there yet... whatever.

        # for i in range (1, 50):
        #     if printSheet.cell(row=i, column=6).value:
        #         if printSheet.cell(row=i, column=6).value.find("Checks") > -1:
        #             cPath = checkDir + printSheet.cell(row=i, column=6).value + ".pdf"
        #             self.CDCommonCode.show_image(cPath)


    def showData(self):
        total = 0
        self.label01 = []
        self.label02 = []
        self.label03 = []
        self.label04 = []
        self.label05 = []
        self.label06 = []
        self.button01 = []
        fileNames = []

        row_num = 6
        self.title = tk.Label(self.frame, text=self.oSheetName, bg="teal", fg="yellow", font='Helvetica 10 bold')
        self.title.grid(row=1, column=1, padx=4, pady=4, sticky=tk.NW)

        self.headline01 = tk.Label(self.frame, text="Name", bg="teal", fg="yellow")
        self.headline01.grid(row=3, column=2, padx=4, pady=2, sticky=tk.W)

        self.headline02 = tk.Label(self.frame, text=" Date ", bg="teal", fg="yellow")
        self.headline02.grid(row=3, column=4, padx=4, pady=2, sticky=tk.W)

        self.headline03 = tk.Label(self.frame, text=" Deposit ", bg="teal", fg="yellow")
        self.headline03.grid(row=3, column=6, padx=4, pady=2, sticky=tk.W)

        self.headline04 = tk.Label(self.frame, text="Check #", bg="teal", fg="yellow")
        self.headline04.grid(row=3, column=8, padx=4, pady=2, sticky=tk.W)

        self.headline05 = tk.Label(self.frame, text="Amount", bg="teal", fg="yellow")
        self.headline05.grid(row=3, column=10, padx=4, pady=2, sticky=tk.W)

        self.headline06 = tk.Label(self.frame, text="Sheet Total", bg="teal", fg="yellow")
        self.headline06.grid(row=3, column=14, padx=4, pady=2, sticky=tk.W)

        self.headline07 = tk.Label(self.frame, text="Image", bg="teal", fg="yellow")
        self.headline07.grid(row=3, column=16, padx=4, pady=2, sticky=tk.W)

        self.headline08 = tk.Label(self.frame, text="Print", bg="yellow", fg="teal")
        self.headline08.grid(row=3, column=18, padx=4, pady=4, sticky=tk.W)
        self.headline08.bind("<Button-1>", partial(self.printSheet, self.oSheetName))


        sortedKeys = []
        for key in self.ds:
            sortedKeys.append(key)
        sortedKeys.sort()
        lastName = ''
        pEnt = ''
        totals = dict()

        self.showCash()

        for ent in sortedKeys:
            for e in self.ds[ent]:
                pEnt = ent
                if lastName == ent:
                    pEnt = ''
                lastName = ent
                self.label01.append(tk.Label(self.frame, text=pEnt, bg="teal", fg="yellow"))
                self.label01[len(self.label01)-1].grid(row=row_num, column=2, padx=4, pady=4, sticky=tk.NW)
                self.label01[len(self.label01)-1].bind("<Button-1>", partial(self.showPerson, pEnt))

                self.label02.append(tk.Label(self.frame, text=e[3], bg="teal", fg="yellow"))
                self.label02[len(self.label02)-1].grid(row=row_num, column=4, padx=4, pady=4, sticky=tk.NW)

                self.label03.append(tk.Label(self.frame, text=e[1], bg="teal", fg="yellow"))
                self.label03[len(self.label03)-1].grid(row=row_num, column=6, padx=4, pady=4, sticky=tk.NW)
                self.label03[len(self.label03)-1].bind("<Button-1>", partial(self.showDeposit, e[1]))

                self.label04.append(tk.Label(self.frame, text=e[0], bg="teal", fg="yellow"))
                self.label04[len(self.label03)-1].grid(row=row_num, column=8, padx=4, pady=4, sticky=tk.NW)

                fAmt = "{:.2f}".format(float(e[5]))
                self.label05.append(tk.Label(self.frame, text=fAmt, bg="teal", fg="yellow"))
                self.label05[len(self.label04)-1].grid(row=row_num, column=10, padx=4, pady=4, sticky=tk.NW)

                if e[7] in totals:
                    totals[e[7]] = totals[e[7]] + e[5]
                else:
                    totals[e[7]] = e[5]

                fAmt2 = "{:.2f}".format(float(totals[e[7]]))
                self.label06.append(tk.Label(self.frame, text="$" + str(fAmt2), bg="teal", fg="yellow"))
                self.label06[len(self.label06)-1].grid(row=row_num, column=14, padx=4, pady=4, sticky=tk.NW)

                self.button01.append(tk.Button(self.frame, text="View", command=partial(self.CDCommonCode.show_image, e[6])))
                self.button01[len(self.button01)-1].grid(row=row_num, column=16, columnspan=2, padx=4, pady=4, sticky=tk.EW)
                total = total + e[5]
                row_num += 1
                line = ''

    def sheetExists(self, sheet):
        dailyLog = self.CDCommonCode.openOneDailyLog(self.oWBName)
        for name in dailyLog.sheetnames:
            if name == sheet:
                return True
        return False

    def runProcess(self, name):
        dailyLog = self.CDCommonCode.openOneDailyLog(self.oWBName)
        self.total = 0
        self.getSheet(name, dailyLog[name])
        self.showData()
