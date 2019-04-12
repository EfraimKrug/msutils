import smtplib
import time
import imaplib
import email
from datetime import datetime

mail = imaplib.IMAP4_SSL('imap.gmail.com')
(retcode, capabilities) = mail.login('KadimahTorasMoshe@gmail.com','August7Brachas')
mail.list()
mail.select('inbox')

n=0
(retcode, messages) = mail.search(None, '(UNSEEN)')
if retcode == 'OK':

   for num in messages[0].split() :
      print 'Processing '
      n=n+1
      typ, data = mail.fetch(num,'(RFC822)')
      for response_part in data:
         #print(response_part)
         if isinstance(response_part, tuple):
             original = email.message_from_string(response_part[1])

             print original['From']
             print original['Subject']
             typ, data = mail.store(num,'+FLAGS','\\Seen')

print n
