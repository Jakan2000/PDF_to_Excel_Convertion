import os
from datetime import datetime

import pandas
from openpyxl.utils import column_index_from_string
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows


class Excel:

    def get_start_end_row_index(wb, startText, endText, startEndDefColumn):
        sheet = wb.active
        start = 0
        end = 0
        for cell in sheet[startEndDefColumn]:
            start += 1
            if startText in str(cell.value):
                break
        for cell in sheet[startEndDefColumn]:
            end += 1
            if endText in str(cell.value):
                break
        return start, end

    def create_slno_column(wb, start, end, column):
        sheet = wb.active
        slno = 1
        for i in range(start, end):
            if i == 1:
                sheet[f"{column}{i}"].value = "Sl.No."
            else:
                sheet[f"{column}{i}"].value = slno
                slno += 1
        return wb

    def column_count(wb):
        sheet = wb.active
        column = 65
        count = 0
        for i in range(column, column + sheet.max_column):
            if sheet[f"{chr(i)}1"].value is None:
                break
            count += 1
        return count

    def creat_column(wb, header):
        sheet = wb.active
        max_column = Excel.column_count(wb) + 1
        column = openpyxl.utils.get_column_letter(max_column)
        sheet[f"{column}1"] = header
        return wb

    def finalise_column(wb, col):
        sheet = wb.active
        missing_columns = []
        column = 65
        for h in range(0, len(col)):
            count = 0
            for i in range(column, column + sheet.max_column):
                if col[h] in str(sheet[f"{chr(i)}1"].value):
                    count += 1
            if count == 0:
                missing_columns.append(col[h])
        if len(missing_columns) != 0:
            for i in range(0, len(missing_columns)):
                Excel.creat_column(wb, missing_columns[i])
        return wb

    def string_align(wb, start, end, column):
        sheet = wb.active
        for i in range(start, end):
            sheet[f"{column}{i}"].value = str(sheet[f"{column}{i}"].value).replace('\n', '')
        return wb

    def alter_header_name(wb, refText, actualText, lastCol):
        sheet = wb.active
        column = 65
        row = 1
        while column < lastCol:
            if refText in str(sheet[f"{chr(column)}{row}"].value):
                sheet[f"{chr(column)}{row}"].value = actualText
            column += 1
        return wb

    def remove_row(wb, start, end, refText, column):
        sheet = wb.active
        for x in range(end, start, -1):
            if refText in str(sheet[f"{column}{x}"].value):
                sheet.delete_rows(x)
                break
        return wb

    def remove_rows(wb, start, end, refText, column):
        sheet = wb.active
        for x in range(end, start, -1):
            if refText in str(sheet[f"{column}{x}"].value):
                sheet.delete_rows(x)
        return wb

    def delete_rows_by_range(wb, start, end, startText, stopText, refcolumn):
        sheet = wb.active
        delete_flag = False
        rows_to_delete = []
        for i in range(start, end):
            if startText in str(sheet[f"{refcolumn}{i}"].value):
                delete_flag = True
            if delete_flag:
                rows_to_delete.append(i)
            if stopText in str(sheet[f"{refcolumn}{i}"].value):
                delete_flag = False
        for x in reversed(rows_to_delete):
            sheet.delete_rows(x)
        return wb

    def delete_column(wb, refText):
        sheet = wb.active
        column_index = None
        for col in range(1, sheet.max_column + 1):
            if refText in sheet.cell(row=1, column=col).value:
                column_index = col
                break
        if column_index is None:
            return wb
        sheet.delete_cols(column_index)
        return wb

    def get_header(wb):
        sheet = wb.active
        header = [sheet["A1"].value, sheet["B1"].value, sheet["C1"].value, sheet["D1"].value, sheet["E1"].value,
                  sheet["F1"].value, sheet["G1"].value, sheet["H1"].value]
        return header

    def find_column_index_by_header(wb, header):
        sheet = wb.active
        column_index = None
        for column in range(65, 65 + sheet.max_column):
            if header in str(sheet[f"{chr(column)}1"].value):
                column_index = column
                break
        return column_index

    def check_neagativeValue_by_column(wb, header):
        sheet = wb.active
        column = Excel.find_column_index_by_header(wb, header)
        for i in range(2, sheet.max_row + 1):
            value = sheet[f"{chr(column)}{i}"].value
            if isinstance(value, str) and value.strip() != '' and float(value.replace(',', '')) < 0.0:
                temp = str(sheet[f"{chr(column)}{i}"].value).replace(',', '')
                sheet[f"{chr(column + 1)}{i}"].value = float(temp.replace("-", ""))
                sheet[f"{chr(column)}{i}"].value = None
        return wb

    def empty_cell_to_none(wb, start, end, header):
        sheet = wb.active
        column = Excel.find_column_index_by_header(wb, header)
        for x in range(start, end):
            if len(str(sheet[f"{chr(column)}{x}"].value)) < 1:
                sheet[f"{chr(column)}{x}"].value = None
        return wb

    def remove_string(wb, start, end, refString, column):
        sheet = wb.active
        for x in range(start, end):
            if refString in str(sheet[f"{column}{x}"].value):
                sheet[f"{column}{x}"].value = str(sheet[f"{column}{x}"].value).replace(refString, "")
        return wb

    def replace_to_none(wb, start, end, refText, column):
        sheet = wb.active
        for x in range(start, end):
            if refText in str(sheet[f"{column}{x}"].value):
                sheet[f"{column}{x}"].value = None
        return wb

    def transaction_type_column(wb):
        sheet = wb.active
        Excel.creat_column(wb, header = "Transaction_Type")
        trans_type_column = chr(Excel.find_column_index_by_header(wb, header = "Transaction_Type"))
        withdrawal_column = chr(Excel.find_column_index_by_header(wb, header = "Withdrawal"))
        deposit_column = chr(Excel.find_column_index_by_header(wb, header = "Deposit"))
        for i in range(2, sheet.max_row + 1):
            if sheet[f"{withdrawal_column}{i}"].value is not None:
                sheet[f"{trans_type_column}{i}"].value = "Debit"
            if sheet[f"{deposit_column}{i}"].value is not None:
                sheet[f"{trans_type_column}{i}"].value = "Credit"
        return wb
