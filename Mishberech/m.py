#!/usr/bin/env python
# -*- coding: utf-8 -*-
import sys
import os
# add parent directory to system path
sys.path.append(os.getcwd()[0:os.getcwd().rfind('\\')])

from AlefBet import *

print(sys.path)
word = u'שלום'

print(getUsableWord(word))
# wordE = 'az'
# #wordR = word[::-1]
# #print(wordR)
# print(ord(word[0]))
# print(ord(word[1]))
# print(ord(word[2]))
# print(ord(word[3]))
#
# print(ord(wordE[0]))
# #print(wordE)
# #print(ord(wordE[1]))
