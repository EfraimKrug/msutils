#aliyos.py
import smtplib

fromaddr = 'KadimahTorasMoshe@gmail.com'
toaddrs  = 'EfraimMKrug@gmail.com'
smtpvar = 'smtp.gmail.com:587'
####################################################
username = 'KadimahTorasMoshe@gmail.com'
password = 'August7Brachas'
####################################################


def writeEmail():
    s = ""
    s += "We have, in our records, that you pledged "
    s += "  We very much appreciate your pledge!\n\nYou can pay online, or, if "
    s += "you prefer, you can send a check to the office at\n\nKadima Toras Moshe\n113 Washington "
    s += "Street\nBrighton, MA 02135\n\nIf this email is in error, please let us know so we can "
    s += "\ncorrect our records."
    s += "\n\nThanks so much!\n\nBest and Good Shabbos! "
    return s

def sendItAll():
    msgTxt = writeEmail()
    msg = "\r\n".join([
       "From: " + fromaddr,
       "To: " + toaddrs,
       "Subject: Aliyah Matanah",
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
