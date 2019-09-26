from CDCommonCode import *
####################################################################################################
### This is where the entire application opens - with a list of all deposits tracked so far -
### in all workbooks in the directory
####################################################################################################
class checkMove01:
    def __init__(self):
        self.CDCommonCode = CDCommonCode()
        self.runProcess()

    def runProcess(self):
        # print("Running check Move...")
        self.CDCommonCode.getCheckScan()

def main():
    CM01 = checkMove01()

if __name__ == '__main__':
    main()
