from CDCommon import *
from checkDisplay01 import *
#######################################################################
def main():
    root = tk.Tk()
    allHistory = False
    if len(sys.argv) > 1:
        if(sys.argv[1] == "All"):
            allHistory = True
    app = checkDisplay01(root, allHistory)
    root.mainloop()

if __name__ == '__main__':
    main()
