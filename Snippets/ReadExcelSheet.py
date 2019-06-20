
##########################################################
# @param sheet: an excel spread sheet (not workbook)
# @return page: an array of dict - one dict for each row
##########################################################
def getValuesFromExcel(sheet):
  page = []
  for r in range(3, sheet.max_row):
      x = dict()
      x['val1'] = str(sheet.cell(row=r,column=1).value).lower()
      x['val2'] = str(sheet.cell(row=r,column=2).value).lower()
      page.append(x)
  return page
