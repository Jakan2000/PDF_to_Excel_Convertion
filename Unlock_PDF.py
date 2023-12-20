import base64

import fitz  # PyMuPDF


def unlock_pdf(pdf_document, password):
    try:
        if pdf_document.isEncrypted:
            if pdf_document.authenticate(password):
                pdf_writer = fitz.open()
                pdf_writer.insert_pdf(pdf_document)
                unlocked_pdf_bytes = pdf_writer.write()
                pdf_writer.close()
                unlocked_pdf_base64 = base64.b64encode(unlocked_pdf_bytes).decode('utf-8')

                response_data = {"data": unlocked_pdf_base64,
                                 "msg": "PDF unlocked and Downloaded successfully."}
                return response_data
            else:
                response_data = {"data": None,
                                 "msg": "Incorrect password. PDF could not be unlocked."}
                return response_data
        else:
            response_data = {"data": None,
                             "msg": "PDF is not encrypted. No need to unlock."}
            return response_data

    except Exception as e:
        # Handle any exceptions that may occur during the process
        response_data = {"data": None,
                         "msg": f"An error occurred: {str(e)}"}
        return response_data
