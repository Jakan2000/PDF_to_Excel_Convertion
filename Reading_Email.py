import configparser
import email
import imaplib
import os
from minio import Minio

from FormatingExcelFiles.main_driver import pdf_to_excel_main

config = configparser.ConfigParser()
config.read(".env")


def minio_upload_pdf(file_path):

    client = Minio('ksvca-server-01:3502', access_key=config.get("DEFAULT", "MINIO_ACCESS_KEY"), secret_key=config.get("DEFAULT", "MINIO_SECRET_KEY"), secure=False)
    bucket_name = 'ksv'
    folder_path = 'bank_statements/'
    file_name = file_path.split('/')[-1]
    client.fput_object(bucket_name, folder_path + file_name, file_path)
    url = client.presigned_get_object(bucket_name, folder_path + file_name, response_headers={'response-content-type': 'application/pdf'})
    return url


def read_emails():
    receiver_email = "cyborgwayne2000@gmail.com"
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    email_password = config.get("DEFAULT", "MAIL_APP_PASSWORD")
    mail.login(receiver_email, email_password)
    mail.select('inbox')

    _, data = mail.search(None, 'UNSEEN')  # Fetch only unread emails
    mail_ids = data[0]
    id_list = mail_ids.split()

    download_directory = os.path.join(os.path.expanduser('~'), 'Downloads')
    email_data_list = []  # List to store email data
    if not os.path.exists(download_directory):
        os.makedirs(download_directory)
    typ, mailbox_data = mail.list()
    all_mail_folder = None

    for item in mailbox_data:
        if b'"[Gmail]/All Mail"' in item:
            all_mail_folder = item.split()[-1].decode('utf-8').strip('\"')

    if all_mail_folder:
        for num in id_list:
            _, data = mail.fetch(num, '(RFC822)')
            raw_email = data[0][1]
            raw_email_string = raw_email.decode('utf-8')
            email_message = email.message_from_string(raw_email_string)
            subject = email_message["Subject"]
            from_address = email_message["From"]
            email_content = ""
            pdf_attachment = None
            for part in email_message.walk():
                if part.get_content_type() == "text/plain":
                    email_content = part.get_payload(decode=True).decode()
                if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
                    continue
                if part.get_filename():
                    if part.get_content_type() == 'application/pdf':
                        file_name = part.get_filename()

                        file_path = os.path.join(download_directory, file_name)
                        pdf_attachment = file_path
                        with open(file_path, 'wb') as fp:
                            fp.write(part.get_payload(decode=True))

            email_data_list.append({
                "From": from_address,
                "Subject": subject,
                "EmailContent": email_content,
                "PDFAttachment": pdf_attachment.replace("\\", "/"),
                "URL": None,
            })

        for num in id_list:
            mail.copy(num, all_mail_folder)
            mail.store(num, '+FLAGS', '\\Deleted')

        mail.expunge()
    mail.close()
    mail.logout()

    if email_data_list:
        for index in range(0, len(email_data_list)):
            url = minio_upload_pdf(file_path=email_data_list[index]["PDFAttachment"])
            temp = url.split("?")
            pdf_url = temp[0]
            email_data_list[index]["URL"] = pdf_url
            os.remove(email_data_list[index]["PDFAttachment"])
            pdf_to_excel_main(email_data_list[index]["URL"], "axis", "type1")

    else:
        print("Email data list is empty.")
        raise Exception("Email data list is empty.")

    return email_data_list  # Return the list of email data


if __name__ == "__main__":
    print(read_emails())
    # minio_upload_pdf(file_path="C:/Users/Admin/Downloads/test_Axis.pdf")
