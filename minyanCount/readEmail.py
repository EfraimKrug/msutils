import smtplib
import time
import imaplib
import email
from datetime import datetime

ORG_EMAIL   = "@gmail.com"
FROM_EMAIL  = "KadimahTorasMoshe" + ORG_EMAIL
FROM_PWD    = "August7Brachas"
SMTP_SERVER = "imap.gmail.com"
SMTP_PORT   = 993
TODAY = datetime.utcnow()

def extractattachements(message):
    if message.get_content_maintype() == 'multipart':
        for part in message.walk():
            if part.get_content_maintype() == 'multipart': continue
            if part.get('Content-Disposition') is None: continue
            #print("HERE1")
            filename = part.get_filename()
            #print("HERE2")
            print filename
            if filename is None: continue
            fb = open(filename,'wb')
            fb.write(part.get_payload(decode=True))
            fb.close()
    else:
        print(message.get_payload())


def getDate(s_dt):
    mths = {'Jan':1, 'Feb':2, 'Mar':3, 'Apr':4, 'May':5, 'Jun':6, 'Jul':7, 'Aug':8, 'Sep':9, 'Oct':10, 'Nov':11, 'Dec':12}
    holdDt = s_dt.split()
    tm = holdDt[4]
    holdTm = holdDt[4].split(":")
    mailDt = datetime(int(holdDt[3]), int(mths[holdDt[2]]), int(holdDt[1]), int(holdTm[0]), int(holdTm[1]))
    return mailDt

def readmail():
    try:
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        mail.select('inbox')

        type, data = mail.search(None, 'ALL')
        mail_ids = data[0]
        id_list = mail_ids.split()

        first_email_id = int(id_list[0])
        latest_email_id = int(id_list[-1])

        for i in range(latest_email_id,first_email_id, -1):
            typ, data = mail.fetch(i, '(RFC822)' )

            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_string(response_part[1])
                    msg_date = getDate(msg['date'])
                    print (TODAY)
                    #if (TODAY - msg_date).seconds > 7200:
                    #    exit()
                    email_subject = msg['subject']
                    email_from = msg['from']

                    #email_body = msg.get_payload()
                    print 'From : ' + email_from + '\n'
                    print 'Subject : ' + str(email_subject) + '\n'
                    extractattachements(msg)
                    #for pl in msg.get_payload():
                    #    print 'Body : ' + pl.get_payload() + '\n'
    except Exception, e:
        print str(e)

readmail()
