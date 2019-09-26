###########################################################
# Reading Deposits SpreadSheet -
# Linking it to all the pieces:
#   List of Deposits (Spread Sheet)
#     Deposit Summary - Spread sheet
#     Deposit Summary - Scanned Document
#     Deposit Slip - Scanned Document
#     Deposit Slip from Bank - Scanned Document
###########################################################
from CDCommon import *
from DepositDisplay01 import *
#######################################################################
def main():
    root = tk.Tk()
    app = DepositDisplay01(root)
    root.mainloop()

if __name__ == '__main__':
    main()
