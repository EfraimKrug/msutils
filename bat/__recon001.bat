@echo off
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
REM  This file checks for the yahrzeit download from shulcloud  ;;
REM  then, either prompts the user to download it, or reformats ;;
REM  the file to make it printable and usable by the gabbai     ;;
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
cd C:\Users\KTM\python
del .\Reconcile\out001

:getfilename1
echo for PayPal processing 'PPTrans.xlsx'
if EXIST .\Reconcile\PPTrans.xlsx (
goto getfilename2
)

type .\Reconcile\ReadMe.Me
goto enderror

:getfilename2
echo for ShulCloud processing 'SCTrans.xlsx'
if EXIST .\Reconcile\SCTrans.xlsx (
.\Reconcile\dist\Recon001 > .\Reconcile\out001
goto endall
)

type .\Reconcile\ReadMe.Me
goto enderror

:endall
"C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE" .\Reconcile\out001

:enderror
echo nope! try again...
