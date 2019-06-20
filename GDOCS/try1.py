#import library
import gspread
#Service client credential from oauth2client
from oauth2client.service_account import ServiceAccountCredentials
# Print nicely
import pprint
#Create scope
scope = ['https://spreadsheets.google.com/feeds']
#create some credential using that scope and content of startup_funding.json
creds = ServiceAccountCredentials.from_json_keyfile_name('creds2.json',scope)
#create gspread authorize using that credential
client = gspread.authorize(creds)
#Now will can access our google sheets we call client.open on StartupName
wb = client.open('KTMPayroll')
#sheet = client.open('KTMPayroll').sheet1
# pp = pprint.PrettyPrinter()
# #Access all of the record inside that
# result = sheet.get_all_record()
#
# result = sheet.row_values(5) #See individual row
# # result = sheet.col.values(5) #See individual column
# #result = sheet.cell(5,2) # See particular cell
# pp = pprint.PrettyPrinter()
#
# #update values
# sheet.update_cell(2,9,'500000')  #Change value at cell(2,9) in the sheet
# result = sheet.cell(2,9)
# pp.pprint(result)
