@echo off
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
REM  This file checks for the yahrzeit download from shulcloud  ;;
REM  then, either prompts the user to download it, or reformats ;;
REM  the file to make it printable and usable by the gabbai     ;;
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
echo opening books, checks and deposits
cd .\BOOKS
.\dist\checkMove01\checkMove01.exe
goto endclean

:endclean
echo Thanks so much, from the author!
