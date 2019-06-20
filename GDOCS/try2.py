import json
import gspread
from oauth2client import SignedJwtAssertionCredentials

json_key = json.load(open('creds2.json')) # json credentials you downloaded earlier
scope = ['https://spreadsheets.google.com/feeds']

credentials = SignedJwtAssertionCredentials(json_key['client_email'], json_key['private_key'].encode(), scope) # get email and key from creds

file = gspread.authorize(credentials) # authenticate with Google
sheet = file.open("KTMPayroll").sheet1 # open sheet
T
