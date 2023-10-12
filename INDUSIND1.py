import os
from datetime import datetime
from openpyxl.utils import column_index_from_string
import openpyxl

from KSV.FormatingExcelFiles.CommonClass import Excel


def dateConversion(wb, start, end, column):
    sheet = wb.active
    for i in range(start, end):
        sheet[f"{column}{i}"].value = datetime.strptime(str(sheet[f"{column}{i}"].value), "%d-%b-%Y").date()
    return wb


def deleteHeader(wb, start):
    sheet = wb.active
    for x in range(start, 0, -1):
        sheet.delete_rows(x)
    return wb


def deleteFooter(wb, start, refColumn):
    sheet = wb.active
    for x in range(sheet.max_row, start, -1):
        if sheet[f"{refColumn}{x}"].value is None or "None" in str(sheet[f"{refColumn}{x}"].value):
            sheet.delete_rows(x)
        else:
            break
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


def deleteRowsByRange(wb, start, end, startText, stopText, startRefcolumn, stopRefColumn):
    sheet = wb.active
    delete_flag = False
    rows_to_delete = []
    for i in range(start, end):
        if startText in str(sheet[f"{startRefcolumn}{i}"].value):
            delete_flag = True
        if delete_flag:
            rows_to_delete.append(i)
        if stopText in str(sheet[f"{stopRefColumn}{i}"].value):
            delete_flag = False
    for x in reversed(rows_to_delete):
        sheet.delete_rows(x)
    return wb


def indusind1_validation(wb):
    sheet = wb.active
    max_column = sheet.max_column
    countOfColumn = 6
    if max_column < countOfColumn or max_column > countOfColumn:
        return True
    else:
        return False


def indusind1_main(wb):
    sheet = wb.active
    if indusind1_validation(wb):
        raise Exception(f"<= INVALID FORMATE =>  <Count Of Column Mismatch>")
    else:
        startText = "Date"
        endText = ""
        startEndRefColumn = "A"
        deleteFlagStartText1 = "Page"
        deleteFlagStopText1 = "Account No"
        deleteFlagStartTextRefColumn1 = "A"
        deleteFlagStopTextRefColumn1 = "B"
        columnToMerg = "B"
        refColumnToMerg = "A"
        deleteFooterRefColumn = "B"
        dateConversionColumn1 = "A"
        refHeaderText1 = "Date"
        refHeaderText2 = "Particulars"
        refHeaderText3 = "Chq./Ref. No"
        refHeaderText4 = "WithDrawal"
        refHeaderText5 = "Deposit"
        refHeaderText6 = "Balance"
        headerText1 = "Transaction_Date"
        headerText2 = "Narration"
        headerText3 = "ChequeNo_RefNo"
        headerText4 = "Withdrawal"
        headerText5 = "Deposit"
        headerText6 = "Balance"
        stringAlignColumn1 = "B"
        columns = ["Transaction_Date", "Value_Date", "ChequeNo_RefNo", "Narration", "Deposit", "Withdrawal", "Balance"]
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndRefColumn)
        end = sheet.max_row
        dupHeaderDeleted1 = deleteRowsByRange(wb, start, end, deleteFlagStartText1, deleteFlagStopText1,
                                              deleteFlagStartTextRefColumn1, deleteFlagStopTextRefColumn1)
        start, end = Excel.get_start_end_row_index(dupHeaderDeleted1, startText, endText, startEndRefColumn)
        end = sheet.max_row
        dupHeaderDeleted2 = deleteRowsByRange(dupHeaderDeleted1, start, end, deleteFlagStartText1, deleteFlagStopText1,
                                              deleteFlagStopTextRefColumn1, deleteFlagStopTextRefColumn1)
        start, end = Excel.get_start_end_row_index(dupHeaderDeleted2, startText, endText, startEndRefColumn)
        end = sheet.max_row
        mergedColumnB = mergingRows(dupHeaderDeleted2, start, end, refColumnToMerg, columnToMerg)
        noneRowsRemoved = removeNoneRows(mergedColumnB, start, end, refColumnToMerg)
        footerDeleted = deleteFooter(noneRowsRemoved, start, deleteFooterRefColumn)
        headerDeleted = deleteHeader(footerDeleted, start - 1)  # start-1 to Skip The Header
        start, end = Excel.get_start_end_row_index(dupHeaderDeleted2, startText, endText, startEndRefColumn)
        end = sheet.max_row
        convertedDateA = dateConversion(headerDeleted, start + 1, end + 1, dateConversionColumn1)
        lastCol = 65 + Excel.column_count(wb)  # 65 -> ASCII value
        transdate = Excel.alter_header_name(convertedDateA, refHeaderText1, headerText1, lastCol)
        narration = Excel.alter_header_name(transdate, refHeaderText2, headerText2, lastCol)
        chqno = Excel.alter_header_name(narration, refHeaderText3, headerText3, lastCol)
        debit = Excel.alter_header_name(chqno, refHeaderText4, headerText4, lastCol)
        credit = Excel.alter_header_name(debit, refHeaderText5, headerText5, lastCol)
        balance = Excel.alter_header_name(credit, refHeaderText6, headerText6, lastCol)
        alignedStringB = Excel.string_align(balance, start, end + 1, stringAlignColumn1)
        columnToCreateSlNo = 65 + Excel.column_count(wb)
        slCreated = Excel.create_slno_column(alignedStringB, start, end + 1, chr(columnToCreateSlNo))
        res = Excel.finalise_column(slCreated, columns)
        return res


if __name__ == "__main__":
    path = "C:/Users/Admin/Downloads/Senthil_indusind_pdf.io__23-09-2023-14-19-31.xlsx"
    wb = openpyxl.load_workbook(path)
    result = indusind1_main(wb)
    result.save("C:/Users/Admin/Desktop/FinalOutput/INDUSIND1output.xlsx")
