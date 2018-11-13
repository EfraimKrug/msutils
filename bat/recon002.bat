@echo off
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
REM  This file checks for the yahrzeit download from shulcloud  ;;
REM  then, either prompts the user to download it, or reformats ;;
REM  the file to make it printable and usable by the gabbai     ;;
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
cd C:\Users\KTM\python
del .\Reconcile\out002

:getfilename1
echo for PayPal processing 'PeoplePP.xlsx'
if EXIST .\Reconcile\PeoplePP.xlsx (
goto getfilename2
)

type .\Reconcile\ReadMe.Me
goto enderror

:getfilename2
echo for ShulCloud processing 'PeopleSC.xlsx'
if EXIST .\Reconcile\PeopleSC.xlsx (
.\Reconcile\dist\Recon002 > .\Reconcile\out002
goto endall
)

type .\Reconcile\ReadMe.Me
goto enderror

:endall
"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE" .\Reconcile\out002

:enderror
echo nope! try again...
