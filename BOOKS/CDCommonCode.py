from CDCommon import *

class CDCommonCode:
    def __init__(self, master):
        self.master = master
        self.cashcheckSwitch = ''
        self.ds = dict()        # {check_name: [check_number, memo, check_date, arrival_date, check_amount, check_image],
        self.workbooks = dict()

    ############################################
    # error display...
    ############################################
    def error_window(self, message):
        self.newWindow = tk.Toplevel(self.master)
        self.app = errorDisplay(self.newWindow, "Crash & Burn: " + message)


    ############################################
    # windows: find the excel program
    ############################################
    def getExcel(self):
        handle = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")

        num_values = winreg.QueryInfoKey(handle)[1]
        for i in range(num_values):
            for x in winreg.EnumValue(handle, i):
                if(str(x).find("EXCEL") > -1):
                    self.EXCELEXE = x

    #############################################
    # dealing with files and directories
    #############################################
    def getFiles(self, files):
        files = []
        file_list=os.listdir(dailyLogDir)
        for  fileN in file_list:
            if fileN.find('print') == 0:
                continue
            if fileN.find('xlsx') > 0:
                files.append(fileN[0:-5])

        return files

    def show_image(self, img):
        try:
            fileName = checkDir + img + ".pdf"
            path_to_pdf = os.path.abspath(fileName)
            path_to_acrobat = os.path.abspath(AcrobatPath)
            process = subprocess.Popen([path_to_acrobat, '/A', 'page=1', path_to_pdf], shell=False, stdout=subprocess.PIPE)
            process.wait()
        except:
            self.error_window("Sorry, that file can not be found!")

    def showDepositImage(self, img, args):
        try:
            fileName = depositDir + '\\' + img + ".pdf"
            path_to_pdf = os.path.abspath(fileName)
            path_to_acrobat = os.path.abspath(AcrobatPath)
            process = subprocess.Popen([path_to_acrobat, '/A', 'page=1', path_to_pdf], shell=False, stdout=subprocess.PIPE)
            process.wait()
        except:
            self.error_window("Sorry, that file can not be found!")

####################################################
# Working with internally with excel files
####################################################

    def buildPage(self, newSheet):
        al = Alignment(horizontal='center', vertical='center')
        newSheet.cell(row=1,column=1).value = datetime.today().strftime('%d-%B-%Y')
        newSheet.cell(row=1,column=7).value = "Deposit: "
        newSheet.cell(row=1,column=7).alignment = al

        newSheet.cell(row=2,column=2).value = "Name on Check"
        newSheet.cell(row=2,column=2).alignment = al
        newSheet.cell(row=2,column=3).value = "Memo"
        newSheet.cell(row=2,column=3).alignment = al
        newSheet.cell(row=2,column=4).value = "Date on Check"
        newSheet.cell(row=2,column=4).alignment = al
        newSheet.cell(row=2,column=5).value = "Amount"
        newSheet.cell(row=2,column=5).alignment = al
        newSheet.cell(row=2,column=6).value = "Image"
        newSheet.cell(row=2,column=6).alignment = al

        newSheet.cell(row=3,column=1).value = "Cash"
        newSheet.cell(row=3,column=1).alignment = al
        newSheet.cell(row=11,column=4).value = "Total: "
        newSheet.cell(row=11,column=5).value = "=SUM(E4:E10)"

        newSheet.cell(row=12,column=1).value = "Check No."
        newSheet.cell(row=12,column=1).alignment = al
        newSheet.cell(row=31,column=3).value = "Sub Total: "
        newSheet.cell(row=31,column=5).value = "=SUM(E13:E30)"
        newSheet.cell(row=32,column=3).value = "Grand Total: "
        newSheet.cell(row=32,column=5).value = "=SUM(E11,E31)"

        #newSheet.column_dimensions[0].width = 20.71
        newSheet.column_dimensions['A'].width = 20
        newSheet.column_dimensions['B'].width = 33
        newSheet.column_dimensions['C'].width = 51
        newSheet.column_dimensions['D'].width = 23
        newSheet.column_dimensions['E'].width = 12
        newSheet.column_dimensions['F'].width = 19
        newSheet.column_dimensions['G'].width = 11

####################################################
# Working with externally with excel files
####################################################
    def createSheet(self, workingFile):
        sheetNameNew = True
        dt = datetime.today().strftime('%B-%d')
        da = dt.split('-')
        sheetName = da[0] + str(da[1])

        # search for the new sheet name...
        for wb in self.workbooks:
            for name in self.workbooks[wb].sheetnames:
                if name == sheetName:
                    sheetNameNew = False
        # new sheet name is not there, create new one...
        if sheetNameNew:
            newSheet = self.getCurrentWorkbook(workingFile).create_sheet(title = sheetName)
            newSheet = self.getCurrentWorkbook(workingFile)[sheetName]
            self.buildPage(newSheet)
            self.getCurrentWorkbook(workingFile).save(dailyLogDir + '\\' + workingFile + '.xlsx')


    def parseName(self, name):
        day = name[-2:]
        month = name[0:-2]
        return (day, month)

    def openNewSheet(self, workingFile):
        self.createSheet(workingFile)
        self.getExcel()
        os.system("start  \"" + self.EXCELEXE + "\" \"" + dailyLogDir + "\\" + workingFile + ".xlsx\"")

    def shiftWBook(self, files, workingFile):
        sheetNames = []
        lastThree = []
        wbList = self.openDailyLog(files)

        dailyLog = wbList[workingFile]
        newBookName = ''

        for name in dailyLog.sheetnames:
            sheetNames.append(name)

        if len(sheetNames) > 10:
            #print("length > 10")
            newBookName = sheetNames[10][0:-2]

        if len(newBookName) < 3:
            #print("length < 3 - crash and burn: " + newBookName)
            return

        newFileName = dailyLogDir + "\\" + newBookName + '.xlsx'
        workingFile = dailyLogDir + "\\" + workingFile + '.xlsx'

        try:
            fh = open(newFileName, 'r')
            print("Sorry - we have already cycled the files: " + newFileName)
            return
        except FileNotFoundError:
            print("Processing new file...")

        for i in range(-3, 0):
            lastThree.append(sheetNames[i])

        wb = Workbook()

        for sheet in lastThree:
            newSheet = wb.create_sheet(title = sheet)
            self.buildPage(newSheet)
            oldSheet = dailyLog[sheet]
            for colN in range(1,20):
                for rowN in range(1,35):
                    newSheet.cell(row=rowN, column=colN).value = oldSheet.cell(row=rowN, column=colN).value

        firstSheet = wb['Sheet']
        wb.remove_sheet(firstSheet)
        dailyLog.save(newFileName)
        wb.save(workingFile)

    # open up the excel files
    # side effect - sets CDCommonCode.workbooks to workbook list
    # param: files array of file names from directory
    # return workbook list
    # emk
    def openDailyLog(self, files):
        if len(files) == 0:
            self.workbooks['DailyLog'] = Workbook()
            dt = datetime.today().strftime('%B-%d')
            da = dt.split('-')
            sheetName = da[0] + str(da[1])
            #newSheet = self.workbooks['dailyLog'].create_sheet(title = sheetName)
            newSheet = self.workbooks['DailyLog']['Sheet']
            self.workbooks['DailyLog']['Sheet'].title = sheetName
            self.buildPage(newSheet)
            #self.workbooks['dailyLog'].remove_sheet('Sheet')
            print(dailyLogDir + '\\DailyLog.xlsx')
            self.workbooks['DailyLog'].save(filename = dailyLogDir + '\\DailyLog.xlsx')
        else:
            for file in files:
                print(dailyLogDir + '\\' + file + '.xlsx')
                self.workbooks[file] = load_workbook(dailyLogDir + '\\' + file + '.xlsx', data_only=True)

        return self.workbooks

    def openOneDailyLog(self, fileName):
        print(dailyLogDir + '\\' + fileName + '.xlsx')
        wb = load_workbook(dailyLogDir + '\\' + fileName + '.xlsx', data_only=True)
        return wb

    def openDepositLog(self, fileName):
        print(depositDir + '\\' + fileName + '.xlsx')
        wb = load_workbook(depositDir + '\\' + fileName + '.xlsx', data_only=True)
        return wb

    def getCurrentWorkbook(self, workingFile):
        return self.workbooks[workingFile]

    def compareMonths(self, m):
        monthOrder = {'january': 1, 'february':2, 'march':3, 'april':4, 'may':5, 'june':6, 'july':7, 'august':8, 'september':9, 'october':10, 'november':11, 'december':12}
        if m.lower()[:-2] in monthOrder:
            return monthOrder[m.lower()[:-2]]
        return 0
        #return monthOrder[m.lower()]
