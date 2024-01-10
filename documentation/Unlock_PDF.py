import base64

import fitz  # PyMuPDF


def unlock_pdf(pdf_document, password):
    """
        Unlock a password-protected PDF file using the provided password.

        Parameters:
        - pdf_document (PyMuPDF.fitz.Document): The locked PDF document to be unlocked.
        - password (str): The password to unlock the PDF document.

        Returns:
        - dict: A dictionary containing the unlocked PDF data in base64 format and an optional success/error message.

        Note:
        This function attempts to unlock a password-protected PDF file using the provided password. If the PDF is encrypted,
        and the password is correct, it creates a new PDF writer, inserts the unlocked PDF, and returns the unlocked PDF
        data in base64 format. If the PDF is not encrypted, it returns a message indicating that no unlocking is necessary.
        If the provided password is incorrect or an error occurs during the process, it returns an error message.
    """
    try:
        if pdf_document.isEncrypted:  # Check if the PDF is encrypted
            if pdf_document.authenticate(password):  # Authenticate using the provided password
                pdf_writer = fitz.open()  # Create a new PDF writer
                pdf_writer.insert_pdf(pdf_document)  # Insert the locked PDF into the writer
                unlocked_pdf_bytes = pdf_writer.write()  # Write the unlocked PDF to a byte stream
                pdf_writer.close()
                unlocked_pdf_base64 = base64.b64encode(unlocked_pdf_bytes).decode('utf-8')  # Encode the byte stream to base6
                response_data = {"data": unlocked_pdf_base64,
                                 "msg": "PDF unlocked and Downloaded successfully."}
                return response_data
            else:
                response_data = {"data": None,
                                 "msg": "Incorrect password. PDF could not be unlocked."}
                return response_data  # returning response with error msg
        else:
            response_data = {"data": None,
                             "msg": "PDF is not encrypted. No need to unlock."}
            return response_data  # returning response with error msg

    except Exception as e:
        # Handle any exceptions that may occur during the process
        response_data = {"data": None,
                         "msg": f"An error occurred: {str(e)}"}
        return response_data  # returning response with error msg
