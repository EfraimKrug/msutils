import os
for path, dirnames, filenames in os.walk('C:\\Program Files\\Microsoft Office\\'):
    if("EXCEL.EXE" in filenames):
        if path.find("Download\\Package") < 0:
            print((path))
