import openpyxl


def seperate_debit_credit_column(wb, start, end):
    sheet = wb.active
    for row in range(start, end):
        if sheet[f"F{row}"].value == "DR":
            sheet[f"H{row}"] = sheet[f"E{row}"].value
        if sheet[f"F{row}"].value == "CR":
            sheet[f"I{row}"] = sheet[f"E{row}"].value
    return wb


if __name__ == '__main__':
    path = "C:/Users/Admin/Desktop/New folder/1._Axis_-_9514__28-12-2023-21-45-18.xlsx"
    wb = openpyxl.load_workbook(path)
    sheet = wb.active
    start = 6
    end = 1471
    seperate_debit_credit_column(wb, start, end)
    wb.save("C:/Users/Admin/Desktop/New folder/1._Axis_-_9514__28-12-2023-21-45-18.xlsx")