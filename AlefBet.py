#!/usr/bin/env python
# -*- coding: utf-8 -*-

#between "0X5BE and "0X5F4
AlefBet =  {"0X5D0":"Alef",
  "0X5D1":"Bet",
  "0X5D2":"Gimel",
  "0X5D3":"Dalet",
  "0X5D4":"He",
  "0X5D5":"Vav",
  "0X5D6":"Zayin",
  "0X5D7":"Het",
  "0X5D8":"Tet",
  "0X5D9":"Yod",
  "0X5DA":"Final-Kaf",
  "0X5DB":"Kaf",
  "0X5DC":"Lamed",
  "0X5DD":"Final-Mem",
  "0X5DE":"Mem",
  "0X5DF":"Final-Nun",
  "0X5E0":"Nun",
  "0X5E1":"Samekh",
  "0X5E2":"Ayin",
  "0X5E3":"Final-Pe",
  "0X5E4":"Pe",
  "0X5E5":"Final-Tsadi",
  "0X5E6":"Tsadi",
  "0X5E7":"Qof",
  "0X5E8":"Resh",
  "0X5E9":"Shin",
  "0X5EA":"Tav",
  "0X20":" "}

# notice: we are going to turn the word around...
Sounds =  {"Ayin":"A",
    "Bet":"B",
    "Tsadi":"ZT",
    "Final-Tsadi":"ZT",
    "Dalet":"D",
    "Gimel":"G",
    "Final-Pe":"F",
    "He":"H",
    "Qof":"K",
    "Kaf":"K",
    "Final-Kaf":"HC",
    "Final-Mem":"M",
    "Mem":"M",
    "Lamed":"L",
    "Nun":"N",
    "Final-Nun":"N",
    "Pe":"P",
    "Samekh":"S",
    "Ayin":"A",
    "Het":"HC",
    "Vav":"U",
    "Tav":"T",
    "Yod":"Y",
    "Resh":"R",
    "Zayin":"Z",
    "Shin":"HS",
    "Tet": "T",
    " ":" "}

def soundOut(word):
    wordE = ""
    for r in range(len(word)):
        x = (hex(ord(word[r]))).upper()
        if str(x) in AlefBet:
            wordE += Sounds[AlefBet[str(x)]]
        else:
            wordE += "'"
    return wordE

def isHebrew(letters):
    if ord(letters.strip()[0]) > 0X5BE and ord(letters.strip()[0]) < 0X5F4:
        return True
    return False

def getUsableWord(word):
    wordReturn = ""
    if isHebrew(word):
        wordReturn = soundOut(word)
        wordReturn = wordReturn[::-1]
    else:
        wordReturn = word
    print ("getUsableWord: " + __name__)
    return wordReturn

###################################################################
# usage:
# word1 = u' של  שלום '
# word2 = " This is weird "
# print(getUsableWord(word1))
# print(getUsableWord(word2))
###################################################################
