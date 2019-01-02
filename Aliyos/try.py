import os
import sys
#########################################################
# get parent directory...
sys.path.append(os.getcwd())
sys.path.append(os.getcwd()[0:os.getcwd().rfind('\\')])

from periodProcess import *

def runProcess():
    print("Running... hah!")

runMonthly(runProcess)
