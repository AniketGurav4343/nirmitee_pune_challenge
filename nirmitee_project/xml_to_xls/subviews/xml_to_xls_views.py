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
            # print(xml_dict)

            # flat_data = flatten(xml_dict)

            # # Convert to DataFrame
            # df = pd.DataFrame([flat_data])

            # # df = pd.json_normalize(xml_dict)

            # # Save DataFrame to Excel
            # output_folder = 'static'
            # file_name = 'output.xlsx'
            # output_file_path = os.path.join(output_folder, file_name)
            # df.to_excel(output_file_path, index=False)

        # tree = ET.parse('./static/Input.xml')
        # root = tree.getroot()

        # # Create a new workbook
        # wb = Workbook()
        # ws = wb.active
        # for row_index, child in enumerate(root.iter()):
        #     if child.text:  # Check if the element contains text
        #         for col_index, item in enumerate(child):
        #             print(row_index+1, col_index+1, item.text)
        #             ws.cell(row=row_index+1, column=col_index+1, value=item.text)

        # file_name = os.path.basename('./static/Input.xml').split('.')[0] + '.xlsx'
        # output_file_path = os.path.join('./static/output/', file_name)
        # wb.save(output_file_path)




        workbook = Workbook()
        sheet = workbook.active

        # sheet.append(list(xml_dict['ENVELOPE']["BODY"]["IMPORTDATA"]["REQUESTDESC"]["STATICVARIABLES"]["SVCURRENTCOMPANY"]["VOUCHER"].keys()))
        sheet.append(list(xml_dict['ENVELOPE']["BODY"]["IMPORTDATA"]["REQUESTDESC"]["STATICVARIABLES"].keys()))
        sheet.append([xml_dict['ENVELOPE']["BODY"]["IMPORTDATA"]["REQUESTDESC"]["STATICVARIABLES"]["SVCURRENTCOMPANY"]])
        # print("xml_dict_voucher", xml_dict['ENVELOPE']["BODY"]["IMPORTDATA"]["REQUESTDATA"]["TALLYMESSAGE"])
        try:
            for item in xml_dict['ENVELOPE']["BODY"]["IMPORTDATA"]["REQUESTDATA"]["TALLYMESSAGE"]:
                # print(item["VOUCHER"]["DATE"], item["VOUCHER"]["REFERENCEDATE"], item["VOUCHER"]["VOUCHERTYPENAME"])
                # sheet.append([item["VOUCHER"]["DATE"], item["VOUCHER"]["REFERENCEDATE"], item["VOUCHER"]["VOUCHERTYPENAME"]])
                sheet.append(["@xmlns:UDF", item["@xmlns:UDF"]])
                sheet.append(['VOUCHER'])
                sheet.append(['', '@REMOTEID', item["VOUCHER"].get("@REMOTEID",'')])
                sheet.append(['', '@VCHKEY', item["VOUCHER"].get("@VCHKEY",'')])
                sheet.append(['', '@VCHTYPE', item["VOUCHER"].get("@VCHTYPE",'')])
                sheet.append(['', '@ACTION', item["VOUCHER"].get("@ACTION",'')])
                sheet.append(['', '@OBJVIEW', item["VOUCHER"].get("@OBJVIEW",'')])
        except:
            pass

        file_name = os.path.basename('./static/Input.xml').split('.')[0] + '.xlsx'
        output_file_path = os.path.join('./static/output/', file_name)
        workbook.save(output_file_path)

        response["success"] = True
        response['message'] = "Good"
        response['data'] = xml_dict
        return Response(response, status=status.HTTP_200_OK)
