import subprocess
from flask import Flask, request, redirect
from twilio.twiml.messaging_response import MessagingResponse

app = Flask(__name__)

html01 = "<html><head><title>Minyan Count</title></head><body>"
html02 = "<div id=count>"
html03 = "</div></body></html>"

@app.route("/sms", methods=['GET', 'POST'])
def sms_reply():
    """Respond to incoming calls with a simple text message."""
    # Start our TwiML response
    minyanCount = readFile()

    number = request.form['From']
    message_body = request.form['Body']

    if message_body.find('y') > -1:
        minyanCount = incFile()
    elif message_body.find('Y') > -1:
        minyanCount = incFile()

    writeHTML(minyanCount)
    subprocess.call('pkill -f firefox', shell=True)
    subprocess.call('firefox file:///home/efraiim/code/msutils/minyanCount/MinyanCount.html', shell=True)

    #print(message_body)
    #print(readFile())
    resp = MessagingResponse()

    # Add a message
    resp.message("The Robots are coming! Head for the hills!")

    return str(resp)

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

def writeHTML(count):
    fd = open("MinyanCount.html", "w")
    html = html01 + html02 + str(count) + html03
    fd.write(html)

if __name__ == "__main__":
    app.run(debug=True)
