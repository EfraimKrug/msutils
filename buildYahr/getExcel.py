import winreg
handle = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe")

num_values = winreg.QueryInfoKey(handle)[1]
vals = []
for i in range(num_values):
    for x in winreg.EnumValue(handle, i):
            vals.append(x)
for j in vals:
    if(str(j).find("EXCEL") > -1):
        print(j)
