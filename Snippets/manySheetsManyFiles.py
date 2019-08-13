import...

for wb in files:
    workbooks[wb] = load_workbook(dailyLogDir + wb + '.xlsx', data_only=True)
    # print(wb + " ==> " + wb[wb.rfind('\\')+1:])
    for name in workbooks[wb].sheetnames:
        if not wb[wb.rfind('\\')+1:] in name:
            print(wb[wb.rfind('\\')+1:-4] + "||" + name + " :: NOPE")
