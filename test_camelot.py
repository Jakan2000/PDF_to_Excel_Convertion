import requests
from io import BytesIO
import camelot
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook
import tempfile

class CamelotConversionError(Exception):
    pass


def pandas_df_to_openpyxl(df):
    # Create a new Openpyxl Workbook
    workbook = Workbook()
    # Create a new worksheet
    worksheet = workbook.active
    # Append the DataFrame data to the worksheet
    for row in dataframe_to_rows(df, index=False, header=False):
        worksheet.append(row)

    return workbook


def test_camelot_main(pdf_url):
    # Download the PDF file from the URL
    response = requests.get(pdf_url)
    pdf_data = BytesIO(response.content)

    # Save the BytesIO content to a temporary PDF file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(pdf_data.getvalue())
        temp_pdf_path = temp_pdf.name

    try:
        # Extract tables from PDF using Camelot
        tables = camelot.read_pdf(temp_pdf_path, flavor='stream', pages='1-end')

        # Concatenate DataFrames for each page into a single DataFrame
        df = pd.concat([table.df for table in tables])

        # Convert DataFrame to Openpyxl Workbook
        wb = pandas_df_to_openpyxl(df)

    except ZeroDivisionError as e:
        print(f"Error: {e}")
        raise CamelotConversionError("Error during Camelot conversion")

    finally:
        # Remove the temporary PDF file
        temp_pdf.close()

    return wb


if __name__ == "__main__":
    pdf_url = "http://ksvca-server-01:3502/ksv/%2Funlock_pdf/2._ICICI_-_4642.pdf"
    # pdf_url = "http://ksvca-server-01:3502/ksv/bank_statements/1.Axis_-_8874-PW_-_GNAN842166790_unlocked.pdf"
    # wb = openpyxl.load_workbook(path)
    result = test_camelot_main(pdf_url)
    result.save('C:/Users/Admin/Desktop/ICICI4output.xlsx')