@echo off
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
REM  This file checks for the yahrzeit download from shulcloud  ;;
REM  then, either prompts the user to download it, or reformats ;;
REM  the file to make it printable and usable by the gabbai     ;;
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
cd C:\Users\KTM\python\msutils
del .\newFiles\emailList

:getfilename1
echo for PayPal processing 'peoplePayPal.xlsx'
if EXIST .\ShulCloud\peoplePayPal.xlsx (
goto getfilename2
)

type .\Reconcile\ReadMe.Me
goto enderror

:getfilename2
echo for ShulCloud processing 'people.xlsx'
if EXIST .\ShulCloud\people.xlsx (
.\py\dist\recon001\recon001 > .\newFiles\reconReport
goto endall
)

type .\Reconcile\ReadMe.Me
goto enderror

:endall
"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE" .\newFiles\reconReport

:enderror
echo nope! try again...
