#!/usr/bin/env python
# -*- coding: utf-8 -*-
import subprocess
import os
import sys
import winreg
#########################################################
# get parent directory...
sys.path.append(os.getcwd())
sys.path.append(os.getcwd()[0:os.getcwd().rfind('\\')])
print(sys.path)
#######################################################
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import tkinter.scrolledtext as tkst
#######################################################
import csv
import shutil
import requests

from datetime import datetime
from datetime import timedelta
from datetime import time
from datetime import date

from openpyxl import load_workbook
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl.utils import get_column_letter
from openpyxl import Workbook

from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Alignment

from functools import partial

import smtplib
import Profile
from SearchNames import *
from checkDisplay02 import *
from checkDisplay03 import *
# print('importing checkDisplay04')
from checkDisplay04 import *
from errorDisplay import *
#from AlefBet import *
#from periodProcess import *
