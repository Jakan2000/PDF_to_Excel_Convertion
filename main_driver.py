import configparser
import os
import os.path
import tkinter
from datetime import datetime
from io import BytesIO
from tkinter import filedialog

import aspose.pdf as ap
import pandas
import psycopg2
import requests
from PyPDF2 import PdfReader, PdfWriter
from openpyxl import Workbook, load_workbook

from AXIS1 import axis1_main
from CANARA1 import canara1_main
from CITY_UNION1 import cityunion1_main
from DBS1 import dbs1_main
from EQUITAS1 import equitas1_main
from FEDERAL1 import federal1_main
from HDFC1 import hdfc1_main
from ICICI1 import icici1_main
from ICICI2 import icici2_main
from ICICI3 import icici3_main
from INDIAN_BANK1 import indian_bank1_main
from INDUSIND1 import indusind1_main
from INDUSIND2 import indusind2_main
from IOB1 import iob1_main
from KOTAK1 import kotak1_main
from KOTAK2 import kotak2_main
from SBI1 import sbi1_main
from TMB1 import tmb1_main
from YES1 import yes1_main


def driver(work_book, bank, type, path):
    # work_book.save("C:/Users/Admin/Desktop/test.xlsx")
    # exit()
    banks = {"axis": {"type1": axis1_main},
             "canara": {"type1": canara1_main},
             "city union": {"type1": cityunion1_main},
             "dbs": {"type1": dbs1_main},
             "equitas": {"type1": equitas1_main},
             "federal": {"type1": federal1_main},
             "hdfc": {"type1": hdfc1_main},
             "icici": {"type1": icici1_main,
                       "type2": icici2_main,
                       "type3": icici3_main},
             "indian bank": {"type1": indian_bank1_main},
             "indusind": {"type1": indusind1_main,
                          "type2": indusind2_main},
             "iob": {"type1": iob1_main},
             "kotak": {"type1": kotak1_main,
                       "type2": kotak2_main},
             "sbi": {"type1": sbi1_main},
             "tmb": {"type1": tmb1_main},
             "yes": {"type1": yes1_main}}

    if bank in banks and type in banks[bank]:
        wb = banks[bank][type](work_book)
        temp = str(os.path.basename(path)).replace(".pdf", ".xlsx")
        file = temp.replace(".PDF", ".xlsx")
        # creatig a temp file in downloads folder
        download_folder = os.path.expanduser("~\\Downloads")
        download_path = os.path.join(download_folder, f"TEMP_{file}")
        wb.save(download_path)

        try:
            df = pandas.read_excel(download_path, na_values=[""])
            # removing the temp file
            os.remove(download_path)
            # Convert "Transaction_Date" and "Value_Date" columns to datetime
            df['Transaction_Date'] = pandas.to_datetime(df['Transaction_Date'], format='%Y-%m-%d')
            df['Value_Date'] = pandas.to_datetime(df['Value_Date'], format='%Y-%m-%d')
            # Extract the date component from "Transaction_Date" and "Value_Date" columns
            df['Transaction_Date'] = df['Transaction_Date'].dt.date
            df['Value_Date'] = df['Value_Date'].dt.date
            column_name1 = "Sl.No."
            column_name2 = "Transaction_Date"
            column_name3 = "Value_Date"
            column_name4 = "ChequeNo_RefNo"
            column_name5 = "Narration"
            column_name6 = "Transaction_Type"
            column_name7 = "Deposit"
            column_name8 = "Withdrawal"
            column_name9 = "Balance"
            column_data1 = df[column_name1]
            column_data2 = df[column_name2]
            column_data3 = df[column_name3]
            column_data4 = df[column_name4]
            column_data5 = df[column_name5]
            column_data6 = df[column_name6]
            column_data7 = df[column_name7]
            column_data8 = df[column_name8]
            column_data9 = df[column_name9]
            new_df = pandas.DataFrame(
                {column_name1: column_data1, column_name2: column_data2, column_name3: column_data3,
                 column_name4: column_data4, column_name5: column_data5, column_name6: column_data6,
                 column_name7: column_data7, column_name8: column_data8, column_name9: column_data9})
            root = tkinter.Tk()
            root.withdraw()
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            new_df.to_excel(file_path, index=False)
            df = new_df.rename(
                columns={"Transaction_Type": "trx_type", "Transaction_Date": "trx_date", "Value_Date": "value_date",
                         "ChequeNo_RefNo": "ref_no_org", "Narration": "description", "Deposit": "deposit",
                         "Withdrawal": "withdrawal", "Balance": "balance"})
            column_names = ["trx_type", "trx_date", "value_date", "ref_no_org", "description", "deposit", "withdrawal",
                            "balance"]
            column_data = df[column_names]
            column_data = column_data.applymap(lambda x: None if pandas.isna(x) else x)
            #  reading data from .env file
            config = configparser.ConfigParser()
            config.read(".env")
            user = config.get("DEFAULT", "USER")
            password = config.get("DEFAULT", "PASSWORD")
            host = config.get("DEFAULT", "HOST")
            port = config.get("DEFAULT", "PORT")
            database = config.get("DEFAULT", "DATABASE")
            postgres_credentials = {
                "user": user,
                "password": password,
                "host": host,
                "port": port,
                "database": database,
            }
            schema = "ksv"
            table_name = "bank_stmt_lines_t"
            connection = psycopg2.connect(**postgres_credentials)
            cursor = connection.cursor()
            insert_query = f"""
                INSERT INTO {schema}.{table_name}
                (trx_type, trx_date, value_date, ref_no_org, description, deposit, withdrawal, balance)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """
            values = [tuple(row) for row in column_data.values]
            cursor.executemany(insert_query, values)
            connection.commit()
            cursor.close()
            connection.close()
            print("Records inserted successfully.")
            return "Records inserted successfully."
        except Exception as e:
            print(f"An error occurred: {e}")
    else:
        raise Exception(f"<Bank or type not found in the dictionary>")


def convert_url_to_bytes(pdf_url):
    bytes_list = []
    response = requests.get(pdf_url)
    response.raise_for_status()  # Check for any request errors

    bytes_stream = BytesIO(response.content)
    reader = PdfReader(bytes_stream)

    for page in reader.pages:
        writer = PdfWriter()
        writer.add_page(page)
        with BytesIO() as bytes_stream:
            writer.write(bytes_stream)
            bytes_stream.seek(0)
            bytes_list.append(bytes_stream.getvalue())

    return bytes_list


def convert_bytes_to_excel(pdf_bytes):
    def create_output_excel(output_xlsx):
        workbook = Workbook()
        sheet = workbook.active
        sheet['A1'] = ''
        workbook.save(output_xlsx)

    def join_data(temp_xlsx, output_xlsx):
        source_workbook_1 = load_workbook(output_xlsx)
        sheet1 = source_workbook_1.active
        source_workbook_2 = load_workbook(temp_xlsx)
        sheet2 = source_workbook_2['Sheet1']
        for row in sheet2.iter_rows(min_row=2, values_only=True):
            sheet1.append(row)
        source_workbook_1.save(output_xlsx)

    def convert_excel(pdf, temp_xlsx, output_xlsx):
        with open(temp_xlsx, 'wb') as f:
            document = ap.Document(BytesIO(pdf))
            save_option = ap.ExcelSaveOptions()
            save_option.minimize_the_number_of_worksheets = True
            document.save(f, options=save_option)
        join_data(temp_xlsx, output_xlsx)

    now = datetime.now()
    t = now.strftime("__%d-%m-%Y-%H-%M-%S")
    output_xlsx = 'output' + t + '.xlsx'
    temp_xlsx = 'temp' + t + '.xlsx'
    create_output_excel(output_xlsx)
    for page_bytes in pdf_bytes:
        convert_excel(page_bytes, temp_xlsx, output_xlsx)
    os.remove(temp_xlsx)

    # Load and return the final workbook
    return load_workbook(output_xlsx), output_xlsx


def pdf_to_excel_main(pdf_url, bank, type):
    pdf_bytes = convert_url_to_bytes(pdf_url)
    output_workbook, output_workbook_xlsx = convert_bytes_to_excel(pdf_bytes)
    sheet = output_workbook.active
    max_column = sheet.max_column
    if max_column < 2:
        raise Exception("Insufficient Data In convert_bytes_to_excel To Process Driver Dictionary")
    else:
        result = driver(output_workbook, bank, type, pdf_url)
    os.remove(output_workbook_xlsx)  # Remove the output workbook XLSX file
    return result


if __name__ == "__main__":
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/1.Axis_-_8874-PW_-_GNAN842166790_unlocked.pdf", "axis", "type1")
    pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/test_Axis.pdf", "axis", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/1.Canara_-_6183.pdf", "canara", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/CITY_UNION_BANK_-_SB-500101012199098.pdf", "city union", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/LVB_-_0145P.W_-_1L1675876_unlocked.pdf", "dbs", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/LVB_-_rama2408_unlocked.pdf", "dbs", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/Equitas_-_6802_unlocked.pdf", "equitas", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/2.%20R%20RAVICHANDRAN%20-%20Federal%20-%202416%20Pass%20-%20RAVI016%20.pdf", "federal", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/HDFC_-_7768.pdf", "hdfc", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/ICICI_-_3281.pdf", "icici", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/ICICI_-_2207PW-088601502207_unlocked.pdf", "icici", "type2")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/ilovepdf_merged_7.pdf", "icici", "type3")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/SRT_-_INDIAN_BANK_-_6096825697_.pdf", "indian bank", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/Senthil_indusind_pdf.io.pdf", "indusind", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/1._Indusind_-_2673.pdf", "indusind", "type2")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/IOB_-_8713.pdf", "iob", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/Kotak1._Apr-22_637102.pdf", "kotak", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/Kotak_-_5887.PDF", "kotak", "type2")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/SBI12._March_-_2023.pdf", "sbi", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/TMB_-_2333.pdf", "tmb", "type1")
    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/2._Jan_to_March_-_TMB_-_363100050305246_Pass-1994_unlocked.pdf", "tmb", "type1") # muthu bro statement

    # pdf_to_excel_main("http://ksvca-server-01:3502/ksv/bank_statements/8._YES_bank_-_8241_Aug-Oct.pdf", "yes", "type1")
