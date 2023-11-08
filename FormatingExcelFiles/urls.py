from django.urls import re_path

from . import views

urlpatterns = [
    re_path("converter/", views.driver),
    re_path("reader/", views.read_mail),
]
