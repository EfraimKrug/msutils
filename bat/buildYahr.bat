@echo off

REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
REM ;; buildYahr - builds the system
REM ;;  1) download the whole shebang from git hub as a zip file
REM ;;  2) open up your file explorer and navigate to the directory
REM ;;  3) right click on the zip file
REM ;;  4) 'extract all' (you can stay in the same folder if you like, or give another place)
REM ;;  5) follow the menus to extract the application
REM ;;  6) note: your base directory may be called msutils or msutils-master
REM ;;  7) navigate to .\msutils\msutils\bat
REM ;;  8) double click on buildYahr.bat
REM ;;  9) now you are ready to use these utilities... enjoy!
REM ;;  ** if you want, you can do the following:
REM ;;      a) right click on mBox.bat
REM ;;      b) 'send to' desktop - this puts the icon on your desktop for easy access
REM ;;
REM ;;  10) Build your yahrzeits spreadsheet:
REM ;;      a) download your yahrzeit spreadsheet from shulcloud
REM ;;      b) copy it from Downloads to .\msutils\msutils\yahrzeits.xlsx
REM ;;      c) doubleclick on your mBox icon
REM ;;      d) click on the 'build yahrzeits' button
REM ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

echo batdir = r'%cd%' > ..\profile.py
cd ..
echo basedir = r'%cd%' >> profile.py
echo yahrdir = r'\yahr' >> profile.py

copy profile.py buildYahr\profile.py
copy profile.py yahr\profile.py
copy profile.py CheckEntry\profile.py
copy profile.py MainBox\profile.py
copy profile.py Reconcile\profile.py

cd buildYahr
dist\buildYahr\buildYahr.exe
