import os
import xmltodict

from django.db import transaction

import pandas as pd

from nirmitee_project.settings import BASE_DIR
#rest_framework import 
from rest_framework.response import Response
from rest_framework.views import APIView
from rest_framework import status

import xml.etree.ElementTree as ET
from openpyxl import Workbook

class ConverXMLtoXLSAPIView(APIView):

    @transaction.atomic
    def post(self, request, format=None):

        response = {}
        folder_path = os.path.join(BASE_DIR, './static/Input.xml')

        with open(folder_path, 'r') as file:
            xml_data = file.read()
            xml_dict = xmltodict.parse(xml_data)
            # json_data = xmltodict.unparse(xml_dict, pretty=True)
            print(xml_dict)

            df = pd.json_normalize(xml_dict)

            # Save DataFrame to Excel
            output_folder = 'static'
            file_name = 'output.xlsx'
            output_file_path = os.path.join(output_folder, file_name)
            df.to_excel(output_file_path, index=False)

        response["success"] = True
        response['message'] = "Good"
        response['data'] = xml_dict
        return Response(response, status=status.HTTP_200_OK)
