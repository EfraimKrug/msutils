@echo off
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
REM  This file checks for the yahrzeit download from shulcloud  ;;
REM  then, either prompts the user to download it, or reformats ;;
REM  the file to make it printable and usable by the gabbai     ;;
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
cd C:\Users\KTM\python\msutils

echo opening books, checks and deposits
cd .\BOOKS
REM C:\Users\KTM\python\msutils\BOOKS\dist\checkDisplay\checkDisplay.exe
.\dist\checkDisplay\checkDisplay.exe
REM python3 checkDisplay.py
goto endclean

:endclean
echo Thanks so much, from the author!
