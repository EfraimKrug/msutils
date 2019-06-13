@echo off
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
REM  This file checks for the yahrzeit download from shulcloud  ;;
REM  then, creates a list of yahrzeit names for the week in a   ;;
REM  file called "temp.txt" - and we bring that up in word.     ;;
REM  As an added benefit, we also bring up CHROME with the page ;;
REM  for latest sh'ma for next shabbos.                         ;;
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
cd C:\Users\KTM\python\msutils\yahr
del temp.txt

echo processing 'yahrzeits.xlsx'
if EXIST yahrzeits.xlsx (
python getNames.py
goto endall
)

echo yahrzeits.xlsx file does not exist
goto enderror

:endall
"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE" C:\Users\KTM\python\msutils\yahr\temp.txt

:enderror
echo nope! try again...
