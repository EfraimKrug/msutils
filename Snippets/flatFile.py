
def storeThisRun(thisRun):
    fd = open(FILE_NAME, 'a')
    fd.write(thisRun + '\n')
    fd.close()

def getCreateTime(thisRun):
    fileName = thisRun
    print("%s" % time.ctime(os.path.getmtime(fileName)))
    str = "%s" % time.ctime(os.path.getmtime(fileName))
    arr = str.split(' ')
    # ['Tue', 'Jul', '30', '12:56:34', '2019']

storeThisRun("MyRun")
getCreateTime('flatFile.py')
