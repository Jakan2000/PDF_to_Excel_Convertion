import os
from datetime import datetime
from openpyxl.utils import column_index_from_string
import openpyxl

from FormatingExcelFiles.CommonClass import Excel


def dateConversion(wb, start, end, column):
    sheet = wb.active
    for i in range(start, end):
        sheet[f"{column}{i}"].value = datetime.strptime(str(sheet[f"{column}{i}"].value), "%d-%b-%Y").date()
    return wb


def removeNone(wb, start, end, column):
    sheet = wb.active
    for x in range(start, end):
        if sheet[f"{column}{x}"].value is not None and "None" in str(sheet[f"{column}{x}"].value):
            sheet[f"{column}{x}"].value = str(sheet[f"{column}{x}"].value).replace("None", "")
    return wb


def splitingDate(wb, start, end, column):
    sheet = wb.active
    valueDateColumn = "H"
    for i in range(start, end):
        spl = str(sheet[f"{column}{i}"].value).split("(")
        sheet[f"{column}{i}"].value = spl[0]
        sheet[f"{valueDateColumn}{i}"].value = spl[1].replace(")", "")
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


def iob1_validation(wb):
    sheet = wb.active
    max_column = sheet.max_column
    countOfColumn = 7
    if max_column < countOfColumn or max_column > countOfColumn:
        return True
    else:
        return False


def iob1_main(wb):
    sheet = wb.active
    if iob1_validation(wb):
        raise Exception(f"<= INVALID FORMATE =>  <Count Of Column Mismatch>")
    else:
        startText = "Date (ValueDate)"
        endText = "available balance"
        startEndRefColumn = "A"
        dateSplitColumn = "A"
        start = 1
        end = sheet.max_row
        stringAlignColumn1 = "A"
        stringAlignColumn2 = "B"
        stringAlignColumn3 = "C"
        stringAlignColumn4 = 'D'
        dateConversionColumn1 = "A"
        dateConversionColumn2 = "H"
        columnToDeleteRefText1 = "TransactionType"
        refHeaderText1 = "Date "
        refHeaderText2 = "Particulars"
        refHeaderText3 = "RefNo./ChequeNo"
        refHeaderText4 = "Debit(Rs)"
        refHeaderText5 = "Credit(Rs)"
        refHeaderText6 = "Balance(Rs)"
        refHeaderText7 = "ValueDate"
        headerText1 = "Transaction_Date"
        headerText2 = "Narration"
        headerText3 = "ChequeNo_RefNo"
        headerText4 = "Withdrawal"
        headerText5 = "Deposit"
        headerText6 = "Balance"
        headerText7 = "Value_Date"
        refTextToReplaceToNone = "-"
        columnToReplaceTextToNone1 = "D"
        columnToReplaceTextToNone2 = "E"
        start = 1
        end = sheet.max_row
        alignedStringA = Excel.string_align(wb, start, end, stringAlignColumn1)
        alignedStringB = Excel.string_align(alignedStringA, start, end, stringAlignColumn2)
        alignedStringC = Excel.string_align(alignedStringB, start, end, stringAlignColumn3)
        alignedStringD = Excel.string_align(alignedStringC, start, end, stringAlignColumn4)
        removedNoneA = removeNone(alignedStringD, start, end, stringAlignColumn1)
        removedNoneB = removeNone(removedNoneA, start, end, stringAlignColumn2)
        removedNoneC = removeNone(removedNoneB, start, end, stringAlignColumn3)
        removedNoneD = removeNone(removedNoneC, start, end, stringAlignColumn4)
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndRefColumn)
        footerDeleted = deleteFooter(wb, end - 1)  # end-1 to Include Last Row
        headerDeleted = deleteHeader(footerDeleted, start - 1)  # start-1 to Skip Header
        start, end = Excel.get_start_end_row_index(headerDeleted, startText, endText, startEndRefColumn)
        dateSplited = splitingDate(headerDeleted, start, end + 1, dateSplitColumn)  # end+1 to Include Last Row
        convertedDateA = dateConversion(dateSplited, start + 1, end + 1, dateConversionColumn1)  # start+1 to Skip Header, end+1 to include last row
        convertedDateH = dateConversion(convertedDateA, start + 1, end + 1, dateConversionColumn2)  # start+1 to Skip Header, end+1 to include last row
        transTypecolumnDeleted = Excel.delete_column(convertedDateH, columnToDeleteRefText1)
        lastCol = 65 + Excel.column_count(wb)
        transdate = Excel.alter_header_name(transTypecolumnDeleted, refHeaderText1, headerText1, lastCol)
        narration = Excel.alter_header_name(transdate, refHeaderText2, headerText2, lastCol)
        chqno = Excel.alter_header_name(narration, refHeaderText3, headerText3, lastCol)
        debit = Excel.alter_header_name(chqno, refHeaderText4, headerText4, lastCol)
        credit = Excel.alter_header_name(debit, refHeaderText5, headerText5, lastCol)
        balance = Excel.alter_header_name(credit, refHeaderText6, headerText6, lastCol)
        valuedate = Excel.alter_header_name(balance, refHeaderText7, headerText7, lastCol)
        columnToCreateSlNo = 65 + Excel.column_count(wb)  # 65 -> ASCII value
        slCreated = Excel.create_slno_column(valuedate, start, end + 1, chr(columnToCreateSlNo))
        replacedToNoneD = Excel.replace_to_none(slCreated, start, end + 1, refTextToReplaceToNone, columnToReplaceTextToNone1)
        replacedToNoneE = Excel.replace_to_none(replacedToNoneD, start, end + 1, refTextToReplaceToNone, columnToReplaceTextToNone2)
        createdTransTypeColumn = Excel.transaction_type_column(replacedToNoneE)
        return wb


if __name__ == "__main__":
    path = "C:/Users/Admin/Downloads/IOB_-_8713__23-09-2023-17-44-36.xlsx"
    wb = openpyxl.load_workbook(path)
    result = iob1_main(wb)
    result.save("C:/Users/Admin/Desktop/FinalOutput/IOB1output.xlsx")
