import subprocess
def readFile():
    fd = open(".count", "r")
    count = fd.read()
    fd.close()
    return count

def writeFile(count):
    fd = open(".count", "w")
    fd.write(str(count))

def incFile():
    count = readFile()
    count = int(count) + 1
    writeFile(count)
    return count

incFile()
print(readFile())
incFile()
print(readFile())
#subprocess.call('firefox file:///home/efraiim/code/msutils/minyanCount/MinyanCount.html', shell=True)

subprocess.call('pkill -f firefox', shell=True)
