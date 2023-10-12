import os
from datetime import datetime
from openpyxl.utils import column_index_from_string
import openpyxl

from KSV.FormatingExcelFiles.CommonClass import Excel


def dateConvertion(wb, start, end, column):
    sheet = wb.active
    for i in range(start, end):
        sheet[f"{column}{i}"].value = datetime.strptime(str(sheet[f"{column}{i}"].value), "%d-%m-%Y").date()
    return wb


def deleteHeader(wb, start):
    sheet = wb.active
    for x in range(start, 0, -1):
        sheet.delete_rows(x)
    return (wb)


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


def icici2_validation(wb):
    sheet = wb.active
    max_column = sheet.max_column
    countOfColumn = 6
    if max_column < countOfColumn or max_column > countOfColumn:
        return True
    else:
        return False


def icici2_main(wb):
    sheet = wb.active
    if icici2_validation(wb):
        raise ValueError(f"<= INVALID FORMATE =>  <Count Of Column Mismatch>")
    else:
        startText = "PARTICULARS"
        endText = "TOTAL"
        startEndDefColumn = "C"
        deleteFlagStartText = "Page"
        deleteFlagEndText = "MODE"
        refColumn1 = "B"
        refColumnToMerg = "A"
        columnToMerg1 = "C"
        openBalRefText = "B/F"
        openBalRefColumn = "C"
        dateConversionColumn = "A"
        refTextDeleteColumn = "MODE**"
        refHeaderText1 = "DATE"
        refHeaderText2 = "PARTICULARS"
        refHeaderText3 = "DEPOSITS"
        refHeaderText4 = "WITHDRAWALS"
        refHeaderText5 = "BALANCE"
        lastCol = 65 + sheet.max_column  # 65 => ASCII value "A"
        headerText1 = "Transaction_Date"
        headerText2 = "Narration"
        headerText3 = "Deposit"
        headerText4 = "Withdrawal"
        headerText5 = "Balance"
        stringAlignColumn1 = "A"
        stringAlignColumn2 = "B"
        stringAlignColumn3 = "C"
        stringAlignColumn4 = "D"
        stringAlignColumn5 = "E"
        columns = ["Transaction_Date", "Value_Date", "ChequeNo_RefNo", "Narration", "Deposit", "Withdrawal", "Balance"]
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndDefColumn)
        rowsRemoveD = Excel.delete_rows_by_range(wb, start, end, deleteFlagStartText, deleteFlagEndText, refColumn1)
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndDefColumn)
        mergColumnC = mergingRows(rowsRemoveD, start, end, refColumnToMerg, columnToMerg1)
        removeNull = removeNoneRows(mergColumnC, start, end, refColumnToMerg)
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndDefColumn)
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndDefColumn)
        footerDeleted = deleteFooter(removeNull, end - 1)  # end-1 to Include Last Row
        headerDeleted = deleteHeader(footerDeleted, start - 1)  # start-1 to Skip Header
        start, end = Excel.get_start_end_row_index(headerDeleted, startText, endText, startEndDefColumn)
        removeOpenBal = Excel.remove_row(headerDeleted, start, end, openBalRefText, openBalRefColumn)
        start, end = Excel.get_start_end_row_index(removeOpenBal, startText, endText, startEndDefColumn)
        dateConvertedA = dateConvertion(removeOpenBal, start + 1, end + 1,
                                        dateConversionColumn)  # start+1 to Skip Header, end+1 to Include Last Row
        deletedModeColumn = Excel.delete_column(wb, refTextDeleteColumn)
        date = Excel.alter_header_name(deletedModeColumn, refHeaderText1, headerText1, lastCol - 1)
        naration = Excel.alter_header_name(date, refHeaderText2, headerText2, lastCol - 1)
        deposits = Excel.alter_header_name(naration, refHeaderText3, headerText3, lastCol - 1)
        withdrawal = Excel.alter_header_name(deposits, refHeaderText4, headerText4, lastCol - 1)
        balance = Excel.alter_header_name(withdrawal, refHeaderText5, headerText5, lastCol - 1)
        alignedA = Excel.string_align(balance, start, end, stringAlignColumn1)
        alignedB = Excel.string_align(alignedA, start, end, stringAlignColumn2)
        columnToCreateSlNo = 65 + Excel.column_count(wb)
        slCreated = Excel.create_slno_column(alignedB, start, end + 1, chr(columnToCreateSlNo))
        res = Excel.finalise_column(slCreated, columns)
        return res


if __name__ == "__main__":
    path = "C:/Users/Admin/Downloads/ICICI_-_2207PW-088601502207_unlocked__15-09-2023-12-58-00.xlsx"
    wb = openpyxl.load_workbook(path)
    result = icici2_main(wb)
    result.save('C:/Users/Admin/Desktop/FinalOutput/ICICI2output.xlsx')
