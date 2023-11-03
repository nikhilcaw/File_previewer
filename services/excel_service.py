import win32com.client
import pythoncom
import os
from flask import Response
import uuid
import excel2img
import io


def excel_to_image_service(excel_file_path):
    try:
        excel_app = win32com.client.Dispatch(
            "Excel.Application", pythoncom.CoInitialize()
        )
        workbook = excel_app.Workbooks.Open(excel_file_path)
        guid = uuid.uuid4()
        excel_image_folder = os.path.abspath("excel_images")
        image_folder_name = str(guid) + "_excel_image"
        image_folder_path = os.path.join(excel_image_folder, image_folder_name)
        if not os.path.exists(image_folder_path):
            os.makedirs(image_folder_path)
        worksheets = workbook.Worksheets
        sheet_number = 1
        for worksheet in worksheets:
            image_name = "image_" + str(sheet_number) + ".png"
            image_path = os.path.join(image_folder_path, image_name)
            excel2img.export_img(excel_file_path, image_path, worksheet.Name, None)
            sheet_number += 1

        workbook.Close()
        excel_app.Quit()
        excel_app = None
    except Exception as e:
        print("Exception occured in excel_to_image service: ", e)
        workbook.Close()
        excel_app.Quit()
        excel_app = None
        return None
    return str(guid)


def send_excel_image_service(guid, sheet_number):
    try:
        excel_image_path = os.path.abspath("excel_images")
        image_folder_name = str(guid) + "_excel_image"
        image_folder_path = os.path.join(excel_image_path, image_folder_name)
        image_name = "image_" + str(sheet_number) + ".png"
        with open(os.path.join(image_folder_path, image_name), "rb") as image_file:
            image_data = io.BytesIO(image_file.read())
            response = Response(image_data, mimetype="image/png")
            response.headers["Content-Type"] = "image/png"
            return response
    except Exception as e:
        print("Exception occured in send_excel_image_service: ", e)
    return None


def save_excel_service(file):
    try:
        temp_excel_folder_path = os.path.abspath("temp_excel_folder")
        file_name = file.filename
        excel_file_path = os.path.join(temp_excel_folder_path, file_name)
        file.save(excel_file_path)
        guid = excel_to_image_service(excel_file_path=excel_file_path)
        os.remove(excel_file_path)
    except Exception as e:
        print("Exception occured in save_excel_service: ", e)
        return None
    return str(guid)
