# import aspose.pdf as ap
#
#
# def convert_pdf_to_excel(input_pdf_path, output_excel_path):
#     # Open PDF document
#     document = ap.Document(input_pdf_path)
#
#     save_option = ap.ExcelSaveOptions()
#
#     # Save the file into MS Excel format
#     document.save(output_excel_path, save_option)
#
#
# if __name__ == "__main__":
#     input_pdf = "C:/Users/Admin/Downloads/KVB_-_7136_PW_-_9875045_unlocked.pdf"
#     output_excel = "C:/Users/Admin/Desktop/output.xlsx"
#
#     convert_pdf_to_excel(input_pdf, output_excel)

# import aspose.pdf as ap
#
# input_pdf = "C:/Users/Admin/Downloads/KVB_-_7136_PW_-_9875045_unlocked.pdf"
# output_excel = "C:/Users/Admin/Desktop/output.csv"
# # Open PDF document
# document = ap.Document(input_pdf)
#
# save_option = ap.ExcelSaveOptions()
# save_option.format = ap.ExcelSaveOptions.ExcelFormat.CSV
#
# # Save the file
# document.save(output_excel, save_option)


import camelot
import pandas
tables = camelot.read_pdf("C:/Users/Admin/Downloads/KVB_-_7136_PW_-_9875045_unlocked.pdf", flavor='stream',pages='all')
df = pandas.concat([table.df for table in tables])
df.to_csv("C:/Users/Admin/Desktop/output.csv")


