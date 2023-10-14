from datetime import datetime

import openpyxl

from KSV.FormatingExcelFiles.CommonClass import Excel


def replace_to_none(wb, start, end, refText, column):
    sheet = wb.active
    for i in range(start, end):
        if refText in str(sheet[f"{column}{i}"].value):
            sheet[f"{column}{i}"].value = None
    return wb


def deleteHeader(wb, start):
    sheet = wb.active
    for x in range(start, 0, -1):
        sheet.delete_rows(x)
    return wb


def deleteFooter(wb, end):
    sheet = wb.active
    for x in range(sheet.max_row, end, -1):
        sheet.delete_rows(x)
    return wb


def dateConversion(wb, start, end, column):
    sheet = wb.active
    for i in range(start, end):
        sheet[f"{column}{i}"].value = datetime.strptime(str(sheet[f"{column}{i}"].value), "%b %d, %Y").date()
    return wb


def indusind2_validation(wb):
    sheet = wb.active
    max_column = sheet.max_column
    countOfColumn = 6
    if max_column < countOfColumn or max_column > countOfColumn:
        return True
    else:
        return False


def indusind2_main(wb):
    sheet = wb.active
    if indusind2_validation(wb):
        raise Exception(f"<= INVALID FORMATE =>  <Count Of Column Mismatch>")
    else:
        startText = "Date"
        endText = "This is a computer generated statement"
        startEndDefColumn = "A"
        deleteFlagStartText1 = "Page"
        deleteFlagStopText1 = "Balance"
        deleteFlagRefColumn1 = "F"
        deleteFlagStartText2 = "Page"
        deleteFlagStopText2 = "Credit"
        deleteFlagRefColumn2 = "E"
        dateConversionColumn1 = "A"
        stringAlignColumn1 = "C"
        negativeValueColumnRefText1 = "Withdrawal"
        refTextToReplace = "-"
        refColumnToReplaceText1 = "C"
        refColumnToReplaceText2 = "D"
        refTextToDeleteColumn1 = "Type"
        refHeaderText1 = "Date"
        refHeaderText2 = "Description"
        refHeaderText3 = "Debit"
        refHeaderText4 = "Credit"
        refHeaderText5 = "Balance"
        headerText1 = "Transaction_Date"
        headerText2 = "Narration"
        headerText3 = "Withdrawal"
        headerText4 = "Deposit"
        headerText5 = "Balance"
        columns = ["Transaction_Date", "Value_Date", "ChequeNo_RefNo", "Narration", "Deposit", "Withdrawal", "Balance"]
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndDefColumn)
        dupHeaderRemoved1 = Excel.delete_rows_by_range(wb, start, end, deleteFlagStartText1, deleteFlagStopText1, deleteFlagRefColumn1)
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndDefColumn)
        dupHeaderRemoved2 = Excel.delete_rows_by_range(dupHeaderRemoved1, start, end, deleteFlagStartText2, deleteFlagStopText2, deleteFlagRefColumn2)
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndDefColumn)
        convertedDateA = dateConversion(dupHeaderRemoved2, start + 1, end, dateConversionColumn1)
        alignedStringC = Excel.string_align(convertedDateA, start, end, stringAlignColumn1)
        footerDeleted = deleteFooter(alignedStringC, end - 1)
        headerDeleted = deleteHeader(footerDeleted, start - 1)
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndDefColumn)
        deletedColumnTYPE = Excel.delete_column(headerDeleted, refTextToDeleteColumn1)
        lastCol = 65 + Excel.column_count(wb)  # 65 -> ASCII value
        transdate = Excel.alter_header_name(deletedColumnTYPE, refHeaderText1, headerText1, lastCol)
        narration = Excel.alter_header_name(transdate, refHeaderText2, headerText2, lastCol)
        debit = Excel.alter_header_name(narration, refHeaderText3, headerText3, lastCol)
        credit = Excel.alter_header_name(debit, refHeaderText4, headerText4, lastCol)
        balance = Excel.alter_header_name(credit, refHeaderText5, headerText5, lastCol)
        replacedNoneD = replace_to_none(balance, start, end + 1, refTextToReplace, refColumnToReplaceText1)
        replacedNoneE = replace_to_none(replacedNoneD, start, end + 1, refTextToReplace, refColumnToReplaceText2)
        columnToCreateSlNo = 65 + Excel.column_count(wb)
        slCreated = Excel.create_slno_column(replacedNoneE, start, end + 1, chr(columnToCreateSlNo))
        columnFinalised = Excel.finalise_column(slCreated, columns)
        createdTransTypeColumn = Excel.transaction_type_column(columnFinalised)
        return wb


if __name__ == "__main__":
    path = "C:/Users/Admin/Downloads/1._Indusind_-_2673__07-10-2023-12-03-48.xlsx"
    wb = openpyxl.load_workbook(path)
    indusind2_main(wb)
    wb.save("C:/Users/Admin/Desktop/FinalOutput/INDUSIND2output.xlsx")
