import openpyxl

from CommonClass import Excel


def deleteFooter(wb, end):
    sheet = wb.active
    for x in range(sheet.max_row, end, -1):
        sheet.delete_rows(x)
    return wb


def align_balance_column(wb, start, end, column):
    sheet = wb.active
    for row in range(start, end):
        if sheet[f"{column}{row}"].value == "None":
            sheet[f"C{row}"].value = sheet[f"D{row}"].value
            sheet[f"D{row}"].value = sheet[f"E{row}"].value
            sheet[f"E{row}"].value = sheet[f"F{row}"].value
            sheet[f"F{row}"].value = sheet[f"G{row}"].value
            sheet[f"G{row}"].value = None
    return wb


def removeNoneRows(wb, start, end, column):  # removing the unwanted rows, when the reference column cell value is none
    sheet = wb.active
    for x in range(end, start, -1):  # iterating through table data end row to table data start row
        if sheet[f"{column}{x}"].value is None:  # if reference column cell value is none
            sheet.delete_rows(x)  # delete row
    return wb


def mergingRows(wb, start, end, refColumn, mergingColumn):  # merging the rows of desired column
    sheet = wb.active
    dataToMerge = []  # array to store row data
    for i in range(start, end):  # iterating through table data start row to table data end row
        slno = sheet[f"{refColumn}{i}"].value  # getting reference column cell value
        if slno is not None:  # if reference column cell is not none then it's the starting row
            if len(dataToMerge) == 0:  # if dataToMerge is empty this is the starting row
                dataToMerge.append(f"{mergingColumn}{i}")  # store cell address in the 0 index
                dataToMerge.append(sheet[f"{mergingColumn}{i}"].value)  # store data from the 1 index
            else:  # if refColumn is not none and dataToMerge is not empty -> it is next starting row
                s = ""  # empty string to merge the row data
                for j in range(1, len(dataToMerge)):  # iterate the dataToMerge array
                    s += str(dataToMerge[j])  # concat the row data
                cell_address = dataToMerge[0]  # take current cell address from 0 index
                sheet[str(cell_address)].value = s  # assign conceited data to the cell
                dataToMerge = []  # emptying dataToMerge ot find the next row starting
                dataToMerge.append(f"{mergingColumn}{i}")  # appending next starting row address in the 0 index
                dataToMerge.append(sheet[f"{mergingColumn}{i}"].value)  # appending row data to the corresponding index
        if slno is None:  # if date is none this is not the starting row
            dataToMerge.append(sheet[f"{mergingColumn}{i}"].value)  # append data to the corresponding index
    # while iterating through the loop the last row will be skipped, so merge the last row by this set of code
    st1 = ""  # empty string to merge the row data
    for m in range(1, len(dataToMerge)):  # iterate dataToMerge array
        st1 += str(dataToMerge[m])  # concat the row data
    cell_address = dataToMerge[0]  # take current cell address from 0 index
    sheet[str(cell_address)].value = st1  # assign conceited data to the cell
    dataToMerge = []  # emptying the dataToMerge
    return wb  # return work book by merging the corresponding rows in the column


def make_cell_none_by_DateLength(wb, start, end, column):
    sheet = wb.active
    yearLength = 6  # length of the string to remove
    for x in range(end, start, -1):  # iterating through table data start row and table data end row
        if len(str(sheet[f"{column}{x}"].value)) < yearLength:  # if cell value length is < year length
            sheet[f"{column}{x}"].value = None  # assign cell value to None
    return wb


def mergingDateColumn(wb, start, end, column):  # merging date column cell values -> incomplete date string
    sheet = wb.active
    inCompleteDateLen = 10  # incomplete date length (string)
    yearLength = 4  # year length (string)
    for i in range(start, end):  # iterating through table data start row to table data end row
        if sheet[f"{column}{i}"].value is not None:  # if cell value in date column is not none
            if len(str(sheet[f"{column}{i}"].value)) < inCompleteDateLen and len(str(sheet[f"{column}{i + 1}"].value)) == yearLength:  # if length of date in cell is less than incomplete date length and length of date in next row cell is equal to year length
                s = str(sheet[f"{column}{i}"].value) + " " + str(sheet[f"{column}{i + 1}"].value)  # concat the date with year
                sheet[f"{column}{i}"].value = s  # assign it to date cell
    return wb


def deleteHeader(wb, start):
    sheet = wb.active
    for x in range(start, 0, -1):
        sheet.delete_rows(x)
    return wb


def aligning_all_column_headers(wb, start, end, refHeadertext, refColumn):
    sheet = wb.active
    for row in range(start, end):
        if sheet[f"{refColumn}{row}"].value is not None:
            if refHeadertext in str(sheet[f"{refColumn}{row}"].value):
                align_column_headers(wb, row, refColumnToMergeHeader="B")
    return wb

def align_column_headers(wb, start, refColumnToMergeHeader):
    sheet = wb.active
    if sheet[f"{refColumnToMergeHeader}{start}"].value is None:
        sheet[f"B{start}"].value = sheet[f"C{start}"].value
        sheet[f"C{start}"].value = sheet[f"D{start}"].value
        sheet[f"D{start}"].value = sheet[f"E{start}"].value
        sheet[f"E{start}"].value = sheet[f"F{start}"].value
        sheet[f"F{start}"].value = sheet[f"G{start}"].value
        sheet[f"G{start}"].value = None
    return wb


def kotak4_validation(wb):
    sheet = wb.active
    max_column = 6
    countOfColumn = Excel.column_count(wb)
    if max_column < countOfColumn or max_column > countOfColumn:
        return True
    else:
        return False


def kotak4_main(wb):
    sheet = wb.active
    startText = "Date"
    stopText = "Statement Summary"
    startEndRefColumn = "A"
    refColumnToMergeHeaderName = "B"
    refHeadertextToAlignAllColumnHeaders = "Date"
    deleteFlagStartText = "Period"
    deleteFlagEndText = "Narration"
    deleteFlagRefColumn = "B"
    columnTomergeDateRows = "A"
    refColumnToMerge = "A"
    columnToMerge1 = "B"
    columnToMerge2 = "C"
    refColumnToRemoveNoneRows = "A"
    balanceColumnToAlign = "F"
    start, end = Excel.get_start_end_row_index(wb, startText, stopText, startEndRefColumn)
    aligning_all_column_headers(wb, start, end + 1, refHeadertextToAlignAllColumnHeaders, startEndRefColumn)
    Excel.delete_rows_by_range(wb, start + 1, end + 1, deleteFlagStartText, deleteFlagEndText, deleteFlagRefColumn)
    start, end = Excel.get_start_end_row_index(wb, startText, stopText, startEndRefColumn)
    deleteHeader(wb, start - 1)
    start, end = Excel.get_start_end_row_index(wb, startText, stopText, startEndRefColumn)
    mergingDateColumn(wb, start + 1, end + 1, columnTomergeDateRows)
    make_cell_none_by_DateLength(wb, start + 1, end + 1, columnTomergeDateRows)
    mergingRows(wb, start, end + 1, refColumnToMerge, columnToMerge1)
    mergingRows(wb, start, end + 1, refColumnToMerge, columnToMerge2)
    deleteFooter(wb, end - 1)
    start, end = Excel.get_start_end_row_index(wb, startText, stopText, startEndRefColumn)
    removeNoneRows(wb, start, end + 1, refColumnToRemoveNoneRows)
    start, end = Excel.get_start_end_row_index(wb, startText, stopText, startEndRefColumn)
    align_balance_column(wb, start, end + 1, balanceColumnToAlign)

    if kotak4_validation(wb):
        print(f"<= INVALID FORMATE : Count Of Column Mismatch =>")
        response = {"data": None,
                    "msg": "<= INVALID FORMATE : Count Of Column Mismatch =>"}
        return response
    return wb


if __name__ == "__main__":
    # path = "C:/Users/Admin/Desktop/KSV/source_excel_files/sasikala_kotak__14-12-2023-10-37-39.xlsx"
    path = "C:/Users/Admin/Downloads/R_JAYARAMAN__27-12-2023-13-29-48.xlsx"
    wb = openpyxl.load_workbook(path)
    result = kotak4_main(wb)
    # result["data"].save("C:/Users/Admin/Desktop/Kotak3output.xlsx")
    result.save("C:/Users/Admin/Desktop/Kotak3output.xlsx")