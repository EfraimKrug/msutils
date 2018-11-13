@echo off
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
REM  This file checks for the yahrzeit download from shulcloud  ;;
REM  then, either prompts the user to download it, or reformats ;;
REM  the file to make it printable and usable by the gabbai     ;;
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
cd C:\Users\KTM\python
del new.xlsx

echo checking for file...
if "%~1"=="" GOTO getfilename
echo .\yahr\dist\yahr\yahr "%~1"
.\yahr\dist\yahr\yahr "%~1"
goto endall

:getfilename
echo no filename given, processing 'yahrzeits.xlsx'
if EXIST yahrzeits.xlsx (
.\yahr\dist\yahr\yahr yahrzeits.xlsx
goto endall
)

echo yahrzeits.xlsx file does not exist
goto enderror

:endall
"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE" new.xlsx

:enderror
echo nope! try again...
