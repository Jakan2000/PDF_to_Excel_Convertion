import openpyxl


def addAlignmentData(data):
    path = "C:/Users/Admin/Desktop/KSV/Python/AlignmentData.xlsx"
    wb = openpyxl.load_workbook(path)
    sheet = wb.active
    row = sheet.max_row+1
    column = 65
    count = 0
    columnCount = len(data)
    for i in range(0, len(data)):
        sheet[f"{chr(column)}{row}"].value = data[i]
        column += 1
    wb.save("C:/Users/Admin/Desktop/KSV/Python/AlignmentData.xlsx")

# addAlignmentData(nestedData=['Kotak1._Apr-22_637102__06-09-2023-14-01-34.xlsx', ' 051112485', 'Undefined', '01-Apr-22 to 30-Apr-22', ['15-Apr-22', 'NEFT FDRLH22105092182 T V THANGARAJFDRL0001789', 'None', 'None', 'NEFTINW-0396757917', '91439', '336,042.86 (Cr)']])