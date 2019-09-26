@echo off
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
REM  This file checks for the yahrzeit download from shulcloud  ;;
REM  then, either prompts the user to download it, or reformats ;;
REM  the file to make it printable and usable by the gabbai     ;;
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
cd ..
REM del .\newFiles\temp.txt

echo processing 'transactions.xlsx'
if EXIST .\shulCloud\transactions.xlsx (
goto skiperror1
)

echo There is no 'transactions.xlsx' file - please download one!
exit 0

:skiperror1
echo processing 'people.xlsx'
if EXIST .\shulCloud\people.xlsx (
goto skiperror2
)

echo There is no 'people.xlsx' file - please download one!
exit 0

:skiperror2
echo Running aliyos and sending email...
cd Aliyos
REM python aliyos.py
.\dist\aliyos\aliyos
