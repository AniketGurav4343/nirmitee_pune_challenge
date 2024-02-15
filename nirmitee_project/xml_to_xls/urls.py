from django.urls import path

from xml_to_xls.views import *

urlpatterns = [

    path('conver_xml_to_xls/', ConverXMLtoXLSAPIView.as_view()),
]