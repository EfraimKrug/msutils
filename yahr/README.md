# yahr

This is a simple utility for synagogue administrators. 

Some of us use ShulCloud.

On ShulCloud - it is possible for the admin to download an excel
spreadsheet of Yahrzeit data. This spreadsheet has many columns.

This utility takes the downloaded spreadsheet, copies the data to
a new spreadsheet (called new.xlsx). The new spread sheet is split 
up by gender, sorted by hebrew yahrzeit date*, and then formatted
for easier readability.

In my shul, I then print the 'male' and 'female' spreadsheets out
for the gabbaim (I do this each month).

This code does nothing secret, impressive or fancy. It simply manipulates
the spreadsheet a bit in order to save a bit of time. I found myself
going through the same motions every month. If you have any use for 
this - please - take it! 

Instructions (how to use this code):
1) Copy this Repo to a windows machine
2) Go to ShulCloud - as admin
3) Admin Menu -> Yahrzeits
4) Fill in dates and download information
5) Press "Export this view"
6) Copy the downloaded file to the folder with yahr.exe
7) On the command line type >yahr file_name
8) Open "new.xlsx" in excel
9) click the tabs on the bottom (Male, Female)
10) Print those worksheets...

* the sort code assumes that all the dates are in the same month, 
and sorts based on day only. I will change this when I get the 
chance.
