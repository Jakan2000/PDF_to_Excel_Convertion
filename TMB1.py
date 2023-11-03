from datetime import datetime

import openpyxl

from FormatingExcelFiles.AlignmentData import addAlignmentData
from FormatingExcelFiles.CommonClass import Excel


def string_in_column(wb, text):
    sheet = wb.active
    for column in range(65, sheet.max_column+65):
        for row in range(1, sheet.max_row):
            if text in str(sheet[f"{chr(column)}{row}"].value):
                return chr(column)


def dateConversion(wb, start, end, column):
    sheet = wb.active
    for i in range(start, end):
        sheet[f"{column}{i}"].value = datetime.strptime(str(sheet[f"{column}{i}"].value), "%d-%m-%Y").date()
    return wb


def alignColumns(wb, start, end, headData, refColumnToAlign):
    sheet = wb.active
    error_records = []
    for i in range(start, end):
        if sheet[f"{refColumnToAlign}{i}"].value is None:
            data = [headData[0], headData[1], headData[2]]
            data.append(i - 1)
            data.append(str(sheet[f"A{i}"].value))
            data.append(str(sheet[f"B{i}"].value))
            data.append(str(sheet[f"C{i}"].value))
            data.append(str(sheet[f"D{i}"].value))
            data.append(str(sheet[f"E{i}"].value))
            data.append(str(sheet[f"F{i}"].value))
            addAlignmentData(data)
            error_records.append(i)
    for x in error_records:
        sheet[f"B{x}"].value = "Error Record"
        sheet[f"C{x}"].value = None
        sheet[f"D{x}"].value = None
        sheet[f"E{x}"].value = None
        sheet[f"F{x}"].value = None


def aligningAllColumns(wb, start, end, refColumn):
    sheet = wb.active
    for i in range(start, end):
        if sheet[f"{refColumn}{i}"].value is None:
            sheet[f'F{i}'].value = sheet[f'E{i}'].value
            sheet[f'E{i}'].value = sheet[f'D{i}'].value
            sheet[f'D{i}'].value = sheet[f'C{i}'].value
            sheet[f'C{i}'].value = None
    return wb


def deleteHeader(wb, start):
    sheet = wb.active
    for x in range(start, 0, -1):
        sheet.delete_rows(x)
    return wb


def headerData(wb, start, end):
    sheet = wb.active
    acnum = "Undefined"
    name = "Undefined"
    period = "Undefined"
    for i in range(start, 0, -1):
        if "Name" in str(sheet[f"A{i}"].value):
            name = str(sheet[f"B{i}"].value).replace(":", "").strip()
        if "Statement for A/c" in str(sheet[f"A{i}"].value):
            spl = str(sheet[f"A{i}"].value).split("Between")
            s = spl[0].split("A/c")
            acnum = s[1].strip()
            period = str(spl[1]).replace("and", "to").strip()
    headData = [acnum, name, period]
    return headData


def deleteFooter(wb, end):
    sheet = wb.active
    for x in range(sheet.max_row, end, -1):
        sheet.delete_rows(x)
    return wb


def removeNoneRows(wb, start, end, column):
    sheet = wb.active
    for x in range(end - 1, start, -1):
        if sheet[f"{column}{x}"].value is None:
            sheet.delete_rows(x)
    return wb


def mergingRows(wb, start, end, refColumn, mergingColumn):
    sheet = wb.active
    dataToMerge = []
    for i in range(start, end):
        slno = sheet[f"{refColumn}{i}"].value
        if slno is not None:
            if len(dataToMerge) == 0:
                dataToMerge.append(f"{mergingColumn}{i}")
                dataToMerge.append(sheet[f"{mergingColumn}{i}"].value)
            else:
                s = ""
                for j in range(1, len(dataToMerge)):
                    s += str(dataToMerge[j])
                cell_address = dataToMerge[0]
                sheet[str(cell_address)].value = s
                dataToMerge = []
                dataToMerge.append(f"{mergingColumn}{i}")
                dataToMerge.append(sheet[f"{mergingColumn}{i}"].value)
        if slno is None:
            dataToMerge.append(sheet[f"{mergingColumn}{i}"].value)
    st1 = ""
    for m in range(1, len(dataToMerge)):
        st1 += str(dataToMerge[m])
    cell_address = dataToMerge[0]
    sheet[str(cell_address)].value = st1
    dataToMerge = []
    return wb


def tmb1_validation(wb):
    sheet = wb.active
    max_column = sheet.max_column
    countOfColumn = 6
    if max_column < countOfColumn or max_column > countOfColumn:
        return True
    else:
        return False


def tmb1_main(wb):
    sheet = wb.active
    if tmb1_validation(wb):
        raise Exception(f"<= INVALID FORMATE =>  <Count Of Column Mismatch>")
    else:
        column = string_in_column(wb, text="Closing Balance")
        if column == "D":
            startText = "Withdrawals"
        if column == "C":
            startText = "Chq. No."
        startEndRefColumn = column
        endText = "Closing Balance"
        deleteFlagRefText1 = "Page"
        deleteFlagRefText2 = "Date"
        deleteFlagRefColumn = "A"
        columnToMerg1 = "B"
        refColumnToMerg = "A"
        refColumnToAlignAllColumn = "F"
        refColumnToAlign = "F"
        refHeaderText1 = "Date"
        refHeaderText2 = "Particulars"
        refHeaderText3 = "Chq. No."
        refHeaderText4 = "Withdrawals"
        refHeaderText5 = "Deposits"
        refHeaderText6 = "Balance(INR)"
        headerText1 = "Transaction_Date"
        headerText2 = "Narration"
        headerText3 = "ChequeNo_RefNo"
        headerText4 = "Withdrawal"
        headerText5 = "Deposit"
        headerText6 = "Balance"
        dateConversionColumn1 = "A"
        columns = ["Transaction_Date", "Value_Date", "ChequeNo_RefNo", "Narration", "Deposit", "Withdrawal", "Balance"]
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndRefColumn)
        pageNoRemoved = Excel.remove_rows(wb, start, end, deleteFlagRefText1, deleteFlagRefColumn)
        start, end = Excel.get_start_end_row_index(pageNoRemoved, startText, endText, startEndRefColumn)
        dupHeaderRemoved = Excel.remove_rows(pageNoRemoved, start, end, deleteFlagRefText2, deleteFlagRefColumn)
        start, end = Excel.get_start_end_row_index(dupHeaderRemoved, startText, endText, startEndRefColumn)
        mergedColumnB = mergingRows(dupHeaderRemoved, start, end, refColumnToMerg, columnToMerg1)
        noneRowsRemoved = removeNoneRows(mergedColumnB, start, end, refColumnToMerg)
        start, end = Excel.get_start_end_row_index(noneRowsRemoved, startText, endText, startEndRefColumn)
        footerDeleted = deleteFooter(noneRowsRemoved, end - 1)  # end-1 to Include End Footer
        start, end = Excel.get_start_end_row_index(footerDeleted, startText, endText, startEndRefColumn)
        headData = headerData(footerDeleted, start, end)
        headerDeleted = deleteHeader(footerDeleted, start - 1)  # start-1 to Skip Header
        start, end = Excel.get_start_end_row_index(headerDeleted, startText, endText, startEndRefColumn)
        allColumnAligned = aligningAllColumns(headerDeleted, start, end + 1, refColumnToAlignAllColumn)
        alignColumns(allColumnAligned, start, end + 1, headData, refColumnToAlign)
        columnToCreateSlNo = 65 + Excel.column_count(wb)
        slCreated = Excel.create_slno_column(allColumnAligned, start, end + 1, chr(columnToCreateSlNo))
        lastCol = 65 + Excel.column_count(wb)
        transdate = Excel.alter_header_name(slCreated, refHeaderText1, headerText1, lastCol)
        naration = Excel.alter_header_name(transdate, refHeaderText2, headerText2, lastCol)
        chqno = Excel.alter_header_name(naration, refHeaderText3, headerText3, lastCol)
        debit = Excel.alter_header_name(chqno, refHeaderText4, headerText4, lastCol)
        credit = Excel.alter_header_name(debit, refHeaderText5, headerText5, lastCol)
        balance = Excel.alter_header_name(credit, refHeaderText6, headerText6, lastCol)
        convertedDateA = dateConversion(balance, start + 1, end + 1, dateConversionColumn1)  # start+1 to Skip Header, end+1 to Include Last Row
        columnFinalised = Excel.finalise_column(convertedDateA, columns)
        createdTransTypeColumn = Excel.transaction_type_column(columnFinalised)
        return wb


if __name__ == "__main__":
    # path = "C:/Users/Admin/Downloads/TMB_-_2333__23-09-2023-11-52-23.xlsx"
    path = "C:/Users/Admin/Desktop/tmb2.xlsx"  # muthu bro statement
    wb = openpyxl.load_workbook(path)
    result = tmb1_main(wb)
    result.save('C:/Users/Admin/Desktop/FinalOutput/TMB1output.xlsx')
