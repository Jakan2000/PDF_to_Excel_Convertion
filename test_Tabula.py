# import tabula
#
# df = tabula.read_pdf("C:/Users/Admin/Downloads/2. ICICI - 4642.pdf", pages='all')
#
# for i in range(len(df)):
#     df[i].to_excel('file' + str(i) + '.xlsx')

# =====================================================================================================================

import tabula
import pandas as pd

# Read PDF and store DataFrames in a list
df_list = tabula.read_pdf("C:/Users/Admin/Desktop/KSV/source_PDF_files/SBI_type1.pdf", pages='all')

# Concatenate all DataFrames into a single DataFrame
combined_df = pd.concat(df_list, ignore_index=True)

# Save the combined DataFrame to a single Excel file
combined_df.to_excel('temp_output.xlsx', index=False)

# ========================================================================================================================


#  not working
# import os
# import os.path
# from datetime import datetime
# from io import BytesIO
#
# import tabula
# import PySimpleGUI as sg
# import pandas as pd
# # import aspose.pdf as ap
# from PyPDF2 import PdfReader, PdfWriter
# from openpyxl import Workbook, load_workbook
#
#
# def create_output_excel(output_xlsx):
#     # Create a new workbook
#     workbook = Workbook()
#
#     # Get the active sheet (default sheet created with the workbook)
#     sheet = workbook.active
#
#     # Add data to the sheet
#     sheet['A1'] = ''
#
#     # Save the workbook
#     workbook.save(output_xlsx)
#
#
# def join_data(temp_xlsx, output_xlsx):
#     source_workbook_1 = load_workbook(output_xlsx)
#     sheet1 = source_workbook_1.active
#
#     source_workbook_2 = load_workbook(temp_xlsx)
#
#     sheet2 = source_workbook_2['Sheet1']
#
#     for row in sheet2.iter_rows(min_row=2, values_only=True):
#         # print (row)
#         sheet1.append(row)
#     source_workbook_1.save(output_xlsx)
#     # print(f"added {page} data")
#
#
# def conver_excel(pdf, temp_xlsx, output_xlsx):
#     df_list = tabula.read_pdf(pdf, pages='all')
#     print("1")
#     x = 2
#     with pd.ExcelWriter(temp_xlsx, engine='xlsxwriter') as writer:
#         print(x)
#         x += 1
#         for i, df in enumerate(df_list):
#             print(f"Processing page {i + 1}")
#             print(x)
#             x += 1
#             df.to_excel(writer, sheet_name=f'Page_{i + 1}', index=False)
#
#     join_data(temp_xlsx, output_xlsx)
#
#
# def split_pdf(input_pdf, temp_xlsx, output_xlsx):
#     with open(input_pdf, "rb") as pdf:
#         bytes_stream = BytesIO(pdf.read())
#
#     reader = PdfReader(bytes_stream)
#     count = 0
#     for page in reader.pages:
#         writer = PdfWriter()
#         writer.add_page(page)
#
#         with BytesIO() as bytes_stream:
#             writer.write(bytes_stream)
#             bytes_stream.seek(0)
#             conver_excel(bytes_stream, temp_xlsx, output_xlsx)
#         count = count + 1
#         print(f"page {count} done")
#     return count
#
#
# def main(input_pdf):
#     now = datetime.now()
#
#     t = now.strftime("__%d-%m-%Y-%H-%M-%S")
#     output_xlsx = input_pdf.replace('.pdf', t) + '.xlsx'
#     temp_xlsx = input_pdf.replace('.pdf', '_temp.xlsx')
#     create_output_excel(output_xlsx)
#     count = split_pdf(input_pdf, temp_xlsx, output_xlsx)
#     os.remove(temp_xlsx)
#     #
#     return 'Process Completed', f"{count} pages converted"
#
#
# file_list_column = [
#     [
#         sg.Text("File Folder"),
#         sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),
#         sg.FolderBrowse(),
#     ],
#     [
#         sg.Listbox(
#             values=[], enable_events=True, size=(40, 20), key="-FILE LIST-"
#         )
#     ],
# ]
#
# # Right side of the GUI
# # Shows the name of the file that was chosen from folder to be converted
# convert_column = [
#     [sg.Text("Choose an .pdf file from list on left:")],
#     [sg.Text(size=(100, 1), key="-TOUT-")],
#     [sg.Button("Convert pdf to excel")],
#     [sg.Text(size=(100, 1), key="-RESP-")],
#     [sg.Text(size=(100, 1), key="-CHAP-")]
# ]
#
# # Combining the left and right layout
# layout = [
#     [
#         sg.Column(file_list_column),
#         sg.VSeperator(),
#         sg.Column(convert_column),
#     ]
# ]
#
# window = sg.Window("KSV PDF converter", layout)  ## App name
#
# # Run the Event Loop
# while True:
#     event, values = window.read()
#     if event == "Exit" or event == sg.WIN_CLOSED:
#         break
#     # Folder name was filled in, make a list of files in the folder
#     if event == "-FOLDER-":
#         folder = values["-FOLDER-"]
#         try:
#             # Get list of files in folder
#             file_list = os.listdir(folder)
#
#         except:
#             file_list = []
#
#         fnames = [
#             f
#             for f in file_list
#             if os.path.isfile(os.path.join(folder, f))
#                and f.lower().endswith((".pdf"))  # Filtering the .xlsx files in the folder
#         ]
#         window["-FILE LIST-"].update(fnames)  # Display the .xlsx files in the list of left column
#
#     elif event == "-FILE LIST-":  # A file was chosen from the listbox
#         try:
#             filename = os.path.join(
#                 values["-FOLDER-"], values["-FILE LIST-"][0]
#             )
#             window["-RESP-"].update(" ")
#             window["-TOUT-"].update(filename)  # Display the full name in the convert column of GUI
#
#         except:
#             pass
#
#     elif event == "Convert pdf to excel":  # When the convert button is pressed
#         try:
#             print("convert process")
#             # chapter_data = chapter_data + "  " + converter(filename)  ##Send the full file name to the converter block
#             return_state = main(filename)
#             window["-RESP-"].update(return_state[0])
#             window["-CHAP-"].update(return_state[1])
#
#         except:
#             pass
#
# window.close()


#=======================================================================================================================

import tabula
import pandas as pd
import os

# Read PDF and store DataFrames in a list
df_list = tabula.read_pdf("C:/Users/Admin/Desktop/KSV/source_PDF_files/SBI_type1.pdf", pages='all')

# Create a temporary directory to store individual Excel files
temp_dir = 'temp_excel_files'
os.makedirs(temp_dir, exist_ok=True)

# Save each DataFrame to a temporary Excel file
temp_excel_files = []
for i, df in enumerate(df_list):
    temp_file_path = os.path.join(temp_dir, f'temp_file_{i}.xlsx')
    df.to_excel(temp_file_path, index=False)
    temp_excel_files.append(temp_file_path)

# Concatenate all Excel files into a single DataFrame
final_df = pd.DataFrame()
for temp_file in temp_excel_files:
    temp_df = pd.read_excel(temp_file)
    final_df = pd.concat([final_df, temp_df], ignore_index=True)

# Save the combined DataFrame to a single Excel file
final_df.to_excel('temp_output.xlsx', index=False)

# Clean up: Remove temporary Excel files and directory
for temp_file in temp_excel_files:
    os.remove(temp_file)
os.rmdir(temp_dir)