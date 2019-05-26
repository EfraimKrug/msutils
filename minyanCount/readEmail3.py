import imaplib
from email.parser import HeaderParser

user = "email"
password =  "pass"

mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login('KadimahTorasMoshe@gmail.com','August7Brachas')
mail.list()
mail.select('inbox')

result, data = mail.search(None, "ALL")

ids = data[0] # data is a list.
id_list = ids.split() # ids is a space separated string
latest_email_id = id_list[-1] # get the latest
result, data = mail.fetch(latest_email_id, '(BODY.PEEK[TEXT])') # fetch the email body for the given ID

header_data = data[0][1] # here's the body, which is raw text of the whole email

parser = HeaderParser()
msg = parser.parsestr(header_data)


print msg
