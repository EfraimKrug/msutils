
def storeThisRun(thisRun):
    fd = open(FILE_NAME, 'a')
    fd.write(thisRun + '\n')
    fd.close()

storeThisRun("MyRun")
