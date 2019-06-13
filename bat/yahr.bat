@echo off
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
REM  This file checks for the yahrzeit download from shulcloud  ;;
REM  then, either prompts the user to download it, or reformats ;;
REM  the file to make it printable and usable by the gabbai     ;;
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
cd C:\Users\KTM\python\msutils
del .\shulCloud\new.xlsx

echo processing 'yahrzeits.xlsx'
if EXIST .\shulCloud\yahrzeits.xlsx (
echo Processing...
cd yahr
REM python yahr.py ..\shulCloud\yahrzeits.xlsx
..\yahr\dist\yahr\yahr ..\shulCloud\yahrzeits.xlsx
goto endall
)

echo yahrzeits.xlsx file does not exist in .\shulCloud
goto enderror

:endall
"C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE" ..\shulCloud\new.xlsx
goto endgood

:enderror
echo nope! try again...

:endgood
echo Thanks! 
