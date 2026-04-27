from django.urls import path
from .views import *



urlpatterns = [
    path("", extract_pdf_data, name="extract_pdf"),
    path("thosibachecking/", thosiba_checking, name="thosiba_checking"),
]