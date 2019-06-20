#######################################################
#inside file: filename - insideFile.py
#######################################################

class insideFile:
    def __init__(self, master):
        self.ds = dict()

        self.master = master
        self.master.configure(bg="teal", pady=34, padx=17)
        self.master.geometry('400x300')
        self.master.title('Some Title')

        self.frame = tk.Frame(self.master, width=360, height=260)
        self.frame.configure(bg="teal", pady=2, padx=2)
        self.frame.grid(row=1, column=1)

#######################################################
#outside file
#######################################################
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import tkinter.scrolledtext as tkst
#######################################################
from insideFile import *
#######################################################################
def main():
    root = tk.Tk()
    app = insideFile(root)
    root.mainloop()

if __name__ == '__main__':
    main()
