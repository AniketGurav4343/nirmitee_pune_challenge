import os
import xmltodict

from django.db import transaction

import pandas as pd

from datetime import datetime

from nirmitee_project.settings import BASE_DIR
#rest_framework import 
from rest_framework.response import Response
from rest_framework.views import APIView
from rest_framework import status

import xml.etree.ElementTree as ET
from openpyxl import Workbook

def date_format_converter(date_string):
    date_object = datetime.strptime(date_string, "%Y%m%d").date()

    formatted_date = date_object.strftime("%d/%m/%Y")
    return formatted_date

def check_object_is_dict_or_list(item):
    if isinstance(item, (str, int, float)):
        return "T"
    elif isinstance(item, type(None)):
        return "N"
    elif isinstance(item, dict):
        return "D"
    elif isinstance(item, list):
        return "L"
    else:
        return "False"

    

class ConverXMLtoXLSAPIView(APIView):

    @transaction.atomic
    def post(self, request, format=None):

        response = {}
        folder_path = os.path.join(BASE_DIR, './static/Input.xml')

        with open(folder_path, 'r') as file:
            xml_data = file.read()
            xml_dict = xmltodict.parse(xml_data)

        workbook = Workbook()
        sheet = workbook.active

        # sheet.append(list(xml_dict['ENVELOPE']["BODY"]["IMPORTDATA"]["REQUESTDESC"]["STATICVARIABLES"]["SVCURRENTCOMPANY"]["VOUCHER"].keys()))
        sheet.append(list(xml_dict['ENVELOPE']["BODY"]["IMPORTDATA"]["REQUESTDESC"]["STATICVARIABLES"].keys()))
        sheet.append([xml_dict['ENVELOPE']["BODY"]["IMPORTDATA"]["REQUESTDESC"]["STATICVARIABLES"]["SVCURRENTCOMPANY"]])
        sheet.append(["TALLYMESSAGE"])
        # print("xml_dict_voucher", xml_dict['ENVELOPE']["BODY"]["IMPORTDATA"]["REQUESTDATA"]["TALLYMESSAGE"])
        # try:
        for item_level_1 in xml_dict['ENVELOPE']["BODY"]["IMPORTDATA"]["REQUESTDATA"]["TALLYMESSAGE"]:
            response_item_level_1 = check_object_is_dict_or_list(item_level_1)
            if response_item_level_1 == "D":
                if len(item_level_1) > 1:
                    for item_level_2 in item_level_1:

                        response_item_level_2 = check_object_is_dict_or_list(item_level_1[item_level_2])
                        if response_item_level_2 == "D":
                            if len(item_level_2) > 1:


                                for item_level_3 in item_level_1[item_level_2]:
                                    response_item_level_3 = check_object_is_dict_or_list(item_level_1[item_level_2][item_level_3])
                                    if response_item_level_3 in ["D", "T"] and len(item_level_3) > 1:
                                        for item_level_4 in item_level_3:
                                            continue
                                            # response_item_level_4 = check_object_is_dict_or_list(item_level_4[item_level_3])
                                            # if response_item_level_4 == "D":
                                            #     if len(item_level_4) > 1:


                                            #         for item_level_5 in item_level_4:
                                            #             response_item_level_5 = check_object_is_dict_or_list(item_level_5[item_level_4])
                                            #             if response_item_level_5 == "D":
                                            #                 if len(item_level_5) > 1:


                                            #                     for item_level_6 in item_level_5:
                                            #                         continue


                                            #             elif response_item_level_5 == "T":
                                            #                 sheet.append(['','','','',item_level_5, item_level_4[item_level_5]])
                                            #             else:
                                            #                 print(response_item_level_5)


                                            # elif response_item_level_4 == "T":
                                            #     sheet.append(['','','', item_level_4, item_level_3[item_level_4]])
                                            # else:
                                            #     print(response_item_level_4)


                                    elif response_item_level_3 in ["T", "N"]:
                                        sheet.append(['','',item_level_1[item_level_2], item_level_1[item_level_2][item_level_3]])
                                    else:
                                        print(response_item_level_3)


                        elif response_item_level_2 in ["T", "N"]:
                            sheet.append(['',item_level_2, item_level_1[item_level_2]])
                        else:
                            print(response_item_level_2)
                else:
                    continue
            else:
                print("bad")
        # except:
        #     pass

        file_name = os.path.basename('./static/Input.xml').split('.')[0] + '.xlsx'
        output_file_path = os.path.join('./static/output/', file_name)
        workbook.save(output_file_path)

        response["success"] = True
        response['message'] = "Good"
        response['data'] = xml_dict
        return Response(response, status=status.HTTP_200_OK)
