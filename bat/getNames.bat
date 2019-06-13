@echo off
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
REM  This file checks for the yahrzeit download from shulcloud  ;;
REM  then, either prompts the user to download it, or reformats ;;
REM  the file to make it printable and usable by the gabbai     ;;
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
cd C:\Users\KTM\python\msutils
del .\newFiles\temp.txt

echo processing 'yahrzeits.xlsx'
if EXIST .\shulCloud\yahrzeits.xlsx (
cd GetNames
REM python getNames.py
..\GetNames\dist\getNames\getNames
goto endall
)

echo yahrzeits.xlsx file does not exist in .\shulCloud
goto enderror

:endall
"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE" ..\newFiles\NamesWeek01.txt
REM "C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE" ..\newFiles\NamesWeek02.txt
goto endclean

:enderror
echo nope! try again...

:endclean
echo Thanks so much, from the author!
