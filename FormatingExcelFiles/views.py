import io
import time

import fitz  # PyMuPDF
import requests
from django.http import HttpResponse
from django.http import JsonResponse

import Unlock_PDF
from Reading_Email import read_emails
from main_driver import pdf_to_excel_main


def process_pdf(request):
    url = request.GET.get('pdf_url')
    bank = request.GET.get('bank')
    type = request.GET.get('type')
    caller = request.GET.get('caller')
    response = pdf_to_excel_main(url, bank, type, caller)
    return JsonResponse(response)


def read_mail(request):
    while True:
        time.sleep(60)
        result = read_emails()
        return HttpResponse({'message': result})


def unlockPDF(request):
    pdf_url = request.GET.get('pdf')
    password = request.GET.get('password')
    pdf_response = requests.get(pdf_url)
    pdf_data = pdf_response.content
    pdf_stream = io.BytesIO(pdf_data)
    pdf_document = fitz.open(stream=pdf_stream, filetype="pdf")
    result = Unlock_PDF.unlock_pdf(pdf_document, password)
    return JsonResponse(result)



