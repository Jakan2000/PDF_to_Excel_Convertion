from django.urls import re_path

from . import views

urlpatterns = [
    re_path("converter/", views.process_pdf),
    re_path("reader/", views.read_mail),
    re_path("unlock_pdf/", views.unlockPDF),
]
