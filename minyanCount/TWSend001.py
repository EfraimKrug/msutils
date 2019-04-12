# Download the helper library from https://www.twilio.com/docs/python/install
from twilio.rest import Client


# Your Account Sid and Auth Token from twilio.com/console
account_sid = 'ACc8f50575262d8e2a86a936e9cb31b945'
auth_token = '89a4ca08b8bacbf420e93e0b6ff60cca'
client = Client(account_sid, auth_token)

message = client.messages \
                .create(
                     body="Join Earth's mightiest heroes. Like Kevin Beefle.",
                     from_='+17817867591',
                     to='+16177806984'
                 )

print(message.sid)
