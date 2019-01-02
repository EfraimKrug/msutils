#Recompile everything...
cd C:\Users\KTM\python\msutils
pyinstaller -y Aliyos\aliyos.py
pyinstaller -y buildYahr\buildYahr.py
pyinstaller -y Mishberech\mishUpdate.py
pyinstaller -y CheckEntry\CheckEntry.py
pyinstaller -y Reconcile\Recon001.py
pyinstaller -y yahr\yahr.py
pyinstaller -y MainBox\mainBox.py
bat\mainBox.bat
