@echo off
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
REM  This file checks for the yahrzeit download from shulcloud  ;;
REM  then, either prompts the user to download it, or reformats ;;
REM  the file to make it printable and usable by the gabbai     ;;
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
cd C:\Users\KTM\python\CheckEntry

:getfilename
echo processing 'accounts.xlsx'
if EXIST accounts.xlsx (
echo found accounts.xlsx
goto getfilename2


echo You need to create an accounts file...
type .\readme.me
goto enderror

:getfilename2
echo processing 'checks.xlsx'
if EXIST checks.xlsx (
echo found checks.xlsx
.\dist\CheckEntry\CheckEntry
goto endall


echo You need to create a checks file...
type .\readme.me
goto enderror

echo yahrzeits.xlsx file does not exist
goto enderror

:endall

goto finalexit
:enderror
echo nope! try again...

:finalexit
echo thanks!

