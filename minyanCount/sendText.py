#aliyos.py
import smtplib

fromaddr = 'KadimahTorasMoshe@gmail.com'
toaddrs  = '6177806984@tmomail.net'
smtpvar = 'smtp.gmail.com:587'
####################################################
username = 'KadimahTorasMoshe@gmail.com'
password = 'August7Brachas'
####################################################


def writeEmail():
    s = ""
    s += "Can you come to minyan?"
    s += "\nThanks so much!\nYou're saving the world!"
    return s

def sendItAll():
    msgTxt = writeEmail()
    msg = "\r\n".join([
       "From: " + fromaddr,
       "To: " + toaddrs,
       "Subject: Minyan",
       "",
       msgTxt
     ])
    return msg

server = smtplib.SMTP(smtpvar)
server.ehlo()
server.starttls()
server.login(username,password)
server.sendmail(fromaddr, toaddrs, sendItAll())
server.quit()
