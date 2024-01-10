import configparser
import email
import imaplib
import os

import fitz
from minio import Minio

from main_driver import pdf_to_excel_main

config = configparser.ConfigParser()  # Create a ConfigParser object
config.read(".env")  # Read the configuration data from the ".env" file


def unlock_pdf(input_path, output_path, password):
    """
       Unlock a password-protected PDF file and save the unlocked version.

       Parameters:
       - input_path (str): Path to the input PDF file.
       - output_path (str): Path to save the unlocked PDF file.
       - password (str): Password used to unlock the PDF.

       Returns:
       str: Path to the saved unlocked PDF file.
    """
    pdf_document = fitz.open(input_path)  # Open the PDF document
    if pdf_document.is_encrypted:  # Check if the PDF is encrypted
        pdf_document.authenticate(password)  # Authenticate with the provided password
    pdf_document.save(output_path)  # Save the unlocked PDF to the specified output path
    pdf_document.close()
    return output_path


def minio_upload_pdf(file_path, bucket_name, folder_path):
    """
        Upload a PDF file to a Minio S3 bucket.

        Parameters:
        - file_path (str): The local path of the PDF file to be uploaded.
        - bucket_name (str): The name of the Minio S3 bucket.
        - folder_path (str): The folder path within the bucket to store the PDF file.

        Returns:
        str: The URL of the uploaded PDF file.
    """
    client = Minio('ksvca-server-01:3502', access_key=config.get("DEFAULT", "MINIO_ACCESS_KEY"),
                   secret_key=config.get("DEFAULT", "MINIO_SECRET_KEY"), secure=False)  # connecting with minio
    file_name = file_path.split('/')[-1]  # Extract file name from the file path
    client.fput_object(bucket_name, folder_path + file_name, file_path)  # Upload the PDF file to the specified bucket and folder
    url = client.presigned_get_object(bucket_name, folder_path + file_name, response_headers={'response-content-type': 'application/pdf'})  # Get the presigned URL for the uploaded PDF file
    return url  # returning pdf url


def read_emails():  # reading unreaded email in inbox folder in Gmail
    """
        Read unread emails in the inbox folder of a Gmail account.

        Returns:
        list: A list of dictionaries containing email data, including sender, subject, email content,
              PDF attachment information, URL after unlocking and uploading, and additional processing.
    """
    receiver_email = "cyborgwayne2000@gmail.com"  # receiver email -> KSV email
    mail = imaplib.IMAP4_SSL('imap.gmail.com')  # Connect to Gmail server
    email_password = config.get("DEFAULT", "MAIL_APP_PASSWORD")  # getting password from environment variable file
    mail.login(receiver_email, email_password)  # logging in to gmail account
    mail.select('inbox')  # selecting inbox folder

    _, data = mail.search(None, 'ALL')  # Fetch All emails
    mail_ids = data[0]  # retrieves first element of tuple, which contains the space-separated string of email IDs.

    id_list = mail_ids.split()  # splits string using whitespace as delimiter, resulting in list of individual mail IDs.

    download_directory = os.path.join(os.path.expanduser('~'), 'Downloads')  # download directory
    email_data_list = []  # List to store email data
    if not os.path.exists(download_directory):  # if download folder not exists in os.path
        os.makedirs(download_directory)  # create new download directory
    typ, mailbox_data = mail.list()
    all_mail_folder = None

    for item in mailbox_data:  # iterate through each item in mailbox_data iterable
        if b'"[Gmail]/All Mail"' in item:  # checks if byte string b'"[Gmail]/All Mail"' present in current item. The b prefix indicates that it is byte string.
            all_mail_folder = item.split()[-1].decode('utf-8').strip('\"')  # decodes byte string into Unicode string using UTF-8 encoding.

    if all_mail_folder:  # check if all_mail_folder is not empty
        for num in id_list:  # Iterates over the list of email
            _, data = mail.fetch(num, '(RFC822)')  # Fetches email data for current email ID using IMAP FETCH command with 'RFC822' data item.
            raw_email = data[0][1]  # Extracts the raw email data from the fetched data.
            raw_email_string = raw_email.decode('utf-8')  # Decodes the raw email data into a Unicode string.
            email_message = email.message_from_string(raw_email_string)  # Parses the email message from the string representation.
            subject = email_message["Subject"]  # Extracts email component subject
            from_address = email_message["From"]  # Extracts email component from address
            email_content = ""
            pdf_attachment = None
            for part in email_message.walk():  # Iterate through parts of email (attachments, text content, etc...)
                if part.get_content_type() == "text/plain":  # if current part is of type "text/plain," indicating it contains plain text content.
                    email_content = part.get_payload(decode=True).decode()  # extracts the content of the part. get_payload method is used to retrieve the payload (content) of the part, and decode=True is used to decode it as bytes.
                if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
                    # Skip further processing for multipart or non-dispositioned parts
                    continue
                if part.get_filename():  # if the part has a filename, indicating it is an attachment
                    if part.get_content_type() == 'application/pdf':
                        file_name = part.get_filename()  # Extracts email component attachment file name
                        file_path = os.path.join(download_directory, file_name)  # combines two components into full file path
                        pdf_attachment = file_path
                        with open(file_path, 'wb') as fp:
                            fp.write(part.get_payload(decode=True))  # Save PDF attachment to the specified file path
            # creating json with from address, subject, email content, pdf attachment, URL
            email_data_list.append({
                "From": from_address,
                "Subject": subject,
                "EmailContent": email_content,
                "PDFAttachment": pdf_attachment.replace("\\", "/") if pdf_attachment else None,
                "URL": None,
            })

        for num in id_list:  # Iterates through the list of email IDs
            mail.copy(num, all_mail_folder)  # copies the email with the specified ID (num)
            mail.store(num, '+FLAGS', '\\Deleted')  # Marks the copied email for deletion. The '+FLAGS' argument is used to add flags to the email, and \\Deleted is a flag indicating that the email should be deleted.

        mail.expunge()  # Permanently removes the emails marked for deletion
    mail.close()  # Closes the currently selected mailbox
    mail.logout()  # Logs out of the IMAP server.

    if email_data_list:  # Checks if the email_data_list is not empty.
        for index in range(len(email_data_list)):  # Iterates through indices of the email data list.
            if email_data_list[index]["PDFAttachment"]:  # Checks if the current email data has a PDF attachment.
                temp_op = email_data_list[index]['PDFAttachment'].split(".pdf")
                output_path = temp_op[0] + "_unlocked" + ".pdf"  # Construct new file path for the unlocked PDF by appending "_unlocked" to the original filename.
                op = unlock_pdf(input_path=email_data_list[index]["PDFAttachment"], output_path=output_path, password="srin2005")  # modify password dinamically
                url = minio_upload_pdf(file_path=op, bucket_name='ksv', folder_path='bank_statements/')  # Uploads the unlocked PDF to a Minio S3 bucket.
                if url:  # Checks if the upload was successful and a URL was returned.
                    temp = url.split("?")  # Splits the URL to remove any query parameters.
                    pdf_url = temp[0]  #  Extracts the base URL without query parameters.
                    email_data_list[index]["URL"] = pdf_url  # Updates email data list with Minio URL of uploaded PDF.
                    os.remove(email_data_list[index]["PDFAttachment"])  # Removes the local PDF file.
                    email_data_list[index]["PDFAttachment"] = op  # Updates email data list with path to unlocked PDF.
                    os.remove(op)  # Removes the local unlocked PDF file.
                    # customise the bank and type dynamically
                    pdf_to_excel_main(email_data_list[index]["URL"], "axis", "type1", "email")

    else:
        print("Email data list is empty.")
        raise Exception("Email data list is empty.")
    print(email_data_list)
    return email_data_list


if __name__ == "__main__":
    read_emails()
    # minio_upload_pdf(file_path="C:/Users/Admin/Downloads/test_Axis.pdf")
    # output_path = unlock_pdf(input_path="C:/Users/Admin/Desktop/Statement_2023MTH10_184523781.pdf", output_path="C:/Users/Admin/Desktop/Statement_2023MTH10_184523781_unlocked.pdf", password='srin2005')
    # os.remove("C:/Users/Admin/Desktop/Statement_2023MTH10_184523781.pdf")
    # os.remove(output_path)
