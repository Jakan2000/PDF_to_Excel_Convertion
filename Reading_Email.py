import configparser
import email
import imaplib
import os
from datetime import datetime

config = configparser.ConfigParser()
config.read(".env")

subject = "Hii bro !.."
body = "hehe hehe he..."


def read_email(sender_email, received_date):
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    receiver_email = "cyborgwayne2000@gmail.com"
    email_password = config.get("DEFAULT", "MAIL_APP_PASSWORD")  # Assuming you've imported 'config' module correctly

    mail.login(receiver_email, email_password)
    mail.select('inbox')

    received_date_formatted = received_date.strftime('%d-%b-%Y')

    _, data = mail.search(None, '(FROM "{}" SINCE "{}")'.format(sender_email, received_date_formatted))
    mail_ids = data[0]
    id_list = mail_ids.split()

    download_directory = os.path.join(os.path.expanduser('~'), 'Downloads')

    if not os.path.exists(download_directory):
        os.makedirs(download_directory)

    for num in id_list:
        _, data = mail.fetch(num, '(RFC822)')
        raw_email = data[0][1]
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)

        subject = email_message["Subject"]
        from_address = email_message["From"]

        print(f"From: {from_address}")
        print(f"Subject: {subject}")

        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                email_content = part.get_payload(decode=True).decode()
                print(email_content)
            if part.get_content_maintype() == 'multipart' or part.get('Content-Disposition') is None:
                continue
            if part.get_filename():
                file_name = part.get_filename()
                file_path = os.path.join(download_directory, file_name)
                with open(file_path, 'wb') as fp:
                    fp.write(part.get_payload(decode=True))

    mail.close()
    mail.logout()


if __name__ == "__main__":
    read_email(sender_email="jakanbob2000@gmail.com", received_date=datetime(2023, 11, 2))
