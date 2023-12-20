from datetime import datetime

import openpyxl

from AlignmentData import addAlignmentData
from CommonClass import Excel


def dateConvertion(wb, start, end, column, ref):
    sheet = wb.active
    for i in range(start, end):
        if ref in str(sheet[f"{column}{i}"].value):
            sheet[f"{column}{i}"].value = datetime.strptime(str(sheet[f"{column}{i}"].value), "%d/%m/%y").date()
    return wb


def alignColumns(wb, start, end, headData, refColumnToAlign):
    sheet = wb.active
    error_records = []
    for i in range(start, end):
        if sheet[f"{refColumnToAlign}{i}"].value is not None:
            data = [headData[0], headData[1], headData[2]]
            data.append(i - 1)
            data.append(str(sheet[f"A{i}"].value))
            data.append(str(sheet[f"B{i}"].value))
            data.append(str(sheet[f"C{i}"].value))
            data.append(str(sheet[f"D{i}"].value))
            data.append(str(sheet[f"E{i}"].value))
            data.append(str(sheet[f"F{i}"].value))
            data.append(str(sheet[f"G{i}"].value))
            data.append(str(sheet[f"H{i}"].value))
            addAlignmentData(data)
            error_records.append(i)
    for x in error_records:
        sheet[f"B{x}"].value = "Error Record"
        sheet[f"C{x}"].value = None
        sheet[f"D{x}"].value = None
        sheet[f"E{x}"].value = None
        sheet[f"F{x}"].value = None
        sheet[f"G{x}"].value = None
        sheet[f"H{x}"].value = None


def deleteHeader(wb, start, end):
    sheet = wb.active
    for x in range(start, 0, -1):
        sheet.delete_rows(x)
    return wb


def aligningAllColumns(wb, start, end, refColumn):
    sheet = wb.active
    for i in range(start, end):
        if sheet[f"{refColumn}{i}"].value is None:
            sheet[f'C{i}'].value = sheet[f'D{i}'].value
            sheet[f'D{i}'].value = sheet[f'E{i}'].value
            sheet[f'E{i}'].value = sheet[f'F{i}'].value
            sheet[f'F{i}'].value = sheet[f'G{i}'].value
            sheet[f'G{i}'].value = sheet[f'H{i}'].value
            sheet[f'H{i}'].value = None
    return wb


def deleteFooter(wb, start, end):
    sheet = wb.active
    for x in range(sheet.max_row, end, -1):
        sheet.delete_rows(x)
    return wb


def headerData(wb, start, end):
    sheet = wb.active
    s1 = "Undefined"
    s2 = "Undefined"
    s3 = "Undefined"
    s4 = "Undefined"
    s5 = "Undefined"
    s6 = "Undefined"
    s7 = "Undefined"
    s8 = "Undefined"
    for i in range(start, 0, -1):
        if sheet[f"B{i}"].value is not None and "Account No" in str(sheet[f"B{i}"].value):
            s1 = sheet[f"D{i}"].value
        if sheet[f"B{i}"].value is not None and "IFSC" in str(sheet[f"B{i}"].value):
            s2 = sheet[f"D{i}"].value
        if sheet[f"A{i}"].value is not None and "Statement From :" in str(sheet[f"A{i}"].value):
            s3 = f"{sheet[f'A{i}'].value} {sheet[f'B{i}'].value}"
    spl1 = s1.split(":")
    a = spl1[3].strip().split(" ")
    acno = a[0]
    cusid = f"Customer ID : {spl1[2]}"
    name = "Undefined"
    ifsc = f"IFSC : {s2}"
    period = s3
    openbal = s4
    closebal = s5
    debits = s6
    credits = s7
    headData = [acno, name, period]
    return headData


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


def hdfc1_validation(wb):
    sheet = wb.active
    max_column = sheet.max_column
    countOfColumn = 7
    if max_column < countOfColumn or max_column > countOfColumn:
        return True
    else:
        return False


def hdfc1_main(wb):
    sheet = wb.active
    if hdfc1_validation(wb):
        print(f"<= INVALID FORMATE : Count Of Column Mismatch =>")
        response = {"data": None,
                    "msg": "<= INVALID FORMATE : Count Of Column Mismatch =>"}
        return response
    else:
        startText = "Narration"
        endText = "STATEMENT SUMMARY :-"
        startEndRefColumn = "B"
        deleteFlagStartText = "HDFC BANK LIMITED"
        deleteFlagStopText = "Statement From :"
        deleteFlagRefColumn = "A"
        dateConversionColumn = "A"
        columnToMerg1 = "B"
        refColumnToMerg = "A"
        refColumnToAlignAllColumn = "C"
        refColumnToAlignColumn = "H"
        dateConversionColumn1 = "A"
        dateConversionColumn2 = "D"
        dateRefText = "/"
        refHeaderText1 = "Date"
        refHeaderText2 = "Narration"
        refHeaderText3 = "Chq./Ref.No."
        refHeaderText4 = "Value Dt"
        refHeaderText5 = "Withdrawal Amt."
        refHeaderText6 = "Deposit Amt."
        refHeaderText7 = "Closing Balance"
        headerText1 = "Transaction_Date"
        headerText2 = "Narration"
        headerText3 = "ChequeNo_RefNo"
        headerText4 = "Value_Date"
        headerText5 = "Withdrawal"
        headerText6 = "Deposit"
        headerText7 = "Balance"
        negativeValueColumnRefText1 = "Withdrawal"
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndRefColumn)
        dupRowsDeleted = Excel.delete_rows_by_range(wb, start, end, deleteFlagStartText, deleteFlagStopText, deleteFlagRefColumn)
        start, end = Excel.get_start_end_row_index(wb, startText, endText, startEndRefColumn)
        mergedColumnB = mergingRows(dupRowsDeleted, start, end, refColumnToMerg, columnToMerg1)
        noneRemoved = removeNoneRows(mergedColumnB, start, end, refColumnToMerg)
        start, end = Excel.get_start_end_row_index(noneRemoved, startText, endText, startEndRefColumn)
        footerDeleted = deleteFooter(noneRemoved, start, end - 1)  # end-1 to IncludeEnd Footer
        start, end = Excel.get_start_end_row_index(footerDeleted, startText, endText, startEndRefColumn)
        headData = headerData(footerDeleted, start, end)
        start, end = Excel.get_start_end_row_index(footerDeleted, startText, endText, startEndRefColumn)
        headerDeleted = deleteHeader(footerDeleted, start - 1, end)  # start-1 to Skip Header
        start, end = Excel.get_start_end_row_index(headerDeleted, startText, endText, startEndRefColumn)
        columnAligned = aligningAllColumns(headerDeleted, start, end + 1, refColumnToAlignAllColumn)  # end+1 to Include Last Row
        alignColumns(columnAligned, start, end, headData, refColumnToAlignColumn)
        convertedDateA = dateConvertion(columnAligned, start + 1, end + 1, dateConversionColumn1, dateRefText)  # start+1 to Skip Header, end+1 to Include Last Row
        convertedDateD = dateConvertion(convertedDateA, start + 1, end + 1, dateConversionColumn2, dateRefText)  # start+1 to Skip Header, end+1 to Include Last Row
        lastCol = 65 + Excel.column_count(wb)
        transdate = Excel.alter_header_name(convertedDateD, refHeaderText1, headerText1, lastCol)
        naration = Excel.alter_header_name(transdate, refHeaderText2, headerText2, lastCol)
        chqno = Excel.alter_header_name(naration, refHeaderText3, headerText3, lastCol)
        valuedate = Excel.alter_header_name(chqno, refHeaderText4, headerText4, lastCol)
        debit = Excel.alter_header_name(valuedate, refHeaderText5, headerText5, lastCol)
        credit = Excel.alter_header_name(debit, refHeaderText6, headerText6, lastCol)
        balance = Excel.alter_header_name(credit, refHeaderText7, headerText7, lastCol)
        columnToCreateSlNo = 65 + Excel.column_count(wb)
        slnoCreated = Excel.create_slno_column(balance, start, end + 1, chr(columnToCreateSlNo))
        negativeValueChecked = Excel.check_neagativeValue_by_column(slnoCreated, negativeValueColumnRefText1)
        createdTransTypeColumn = Excel.transaction_type_column(negativeValueChecked)
        response = {"data": wb,
                    "msg": None}
        return response


if __name__ == "__main__":
    path = "C:/Users/Admin/Downloads/hdfc.xlsx"
    wb = openpyxl.load_workbook(path)
    result = hdfc1_main(wb)
    result.save('C:/Users/Admin/Desktop/HDFC1output.xlsx')
