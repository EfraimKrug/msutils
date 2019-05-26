from flask import Flask
app = Flask(__name__)
fName = '/home/efraiim/code/msutils/minyanCount/WEB/.count'
linkString =  "<a href='http://127.0.0.1:5000/inc'>Click</a>"
countString = "<div id=count><font size=26>XXXX</font></div>"
@app.route("/")
def index():
    return linkString

@app.route("/hello")
def hello():
    return "Hello World!"

@app.route("/inc")
def members():
    count = incFile()
    return linkString + "<br>" + countString.replace("XXXX", str(count))

def readFile():
    global fName
    fd = open(fName, "r")
    count = fd.read()
    fd.close()
    return count

def writeFile(count):
    global fName
    fd = open(fName, "w")
    fd.write(str(count))

def incFile():
    count = readFile()
    count = int(count) + 1
    writeFile(count)
    return count

if __name__ == "__main__":
    app.run()
