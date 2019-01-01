import sys
import os
#########################################################
# get parent directory...
sys.path.append(os.getcwd()[0:os.getcwd().rfind('\\')])
print(sys.path)
from Profile import *
print(batdir)
