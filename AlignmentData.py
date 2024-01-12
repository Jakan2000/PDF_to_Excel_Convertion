import openpyxl


def addAlignmentData(data):
    """
        Adds misaligned data to an existing Excel file.

        Parameters:
        - data (list): A list containing the misaligned data to be added to the Excel file.

        Notes:
        - The function appends the provided misaligned data to an existing Excel file.
        - The Excel file path is hardcoded in the function. Modify the 'path' variable to match your file path.
        - The data is added to the next available row in the active sheet.
        - The function uses ASCII values to determine the column in which the data will be added.
        - After adding the data, the Excel file is saved with the new changes.

    """
    path = "C:/Users/Admin/Desktop/KSV/Python/AlignmentData.xlsx"  # path of the excel file
    wb = openpyxl.load_workbook(path)  # load the work book
    sheet = wb.active  # get the active sheet
    row = sheet.max_row+1  # get last row new row
    column = 65  # ascii value of column A
    count = 0
    columnCount = len(data)
    for i in range(0, len(data)):  # iterate through data array length
        sheet[f"{chr(column)}{row}"].value = data[i]  # assigning the miss aligned data to excel sheet
        column += 1  # by incrementing the column by 1
    wb.save("C:/Users/Admin/Desktop/KSV/Python/AlignmentData.xlsx")  # save the excel file

# if __name__ == '__main__':
    # addAlignmentData(['Kotak1._Apr-22_637102__06-09-2023-14-01-34.xlsx', ' 051112485', 'Undefined', '01-Apr-22 to 30-Apr-22', ['15-Apr-22', 'NEFT FDRLH22105092182 T V THANGARAJFDRL0001789', 'None', 'None', 'NEFTINW-0396757917', '91439', '336,042.86 (Cr)']])