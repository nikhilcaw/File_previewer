import os
import win32com.client
import pythoncom
import uuid
from flask import Response
import io


def ppt_to_image(ppt_file_path):
    try:
        pptapp = win32com.client.Dispatch(
            "PowerPoint.Application", pythoncom.CoInitialize()
        )
        # breakpoint()
        presentation = pptapp.Presentations.Open(ppt_file_path)
        guid = uuid.uuid4()
        ppt_image_folder = os.path.abspath("ppt_images")
        image_folder_name = str(guid) + "ppt_images"
        image_folder_path = os.path.join(ppt_image_folder, image_folder_name)
        if not os.path.exists(image_folder_path):
            os.makedirs(image_folder_path)

        image_number = 1
        for slide in presentation.Slides:
            image_name = "image_" + str(image_number) + ".jpg"
            image_path = os.path.join(image_folder_path, image_name)
            slide.Export(image_path, "JPG", 800, 600)
            image_number += 1

        presentation.Close()
        pptapp.Quit()
        pptapp = None
    except Exception as e:
        print("Error occured in ppt_to_image_service : ", e)
        return None
    return str(guid)


def send_ppt_image_service(guid, slide_no):
    try:
        ppt_image_path = os.path.abspath("ppt_images")
        image_folder_name = str(guid) + "ppt_images"
        image_folder_path = os.path.join(ppt_image_path, image_folder_name)
        image_name = "image_" + str(slide_no) + ".jpg"
        with open(os.path.join(image_folder_path, image_name), "rb") as image_file:
            image_data = io.BytesIO(image_file.read())
            response = Response(image_data, mimetype="image/jpeg")
            response.headers["Content-Type"] = "image/jpeg"
            return response
    except Exception as e:
        print("Exception occured in send_ppt_image_service: ", e)
    return None


def save_ppt_service(file):
    try:
        temp_ppt_folder_path = os.path.abspath("temp_ppt_folder")
        file_name = file.filename
        ppt_file_path = os.path.join(temp_ppt_folder_path, file_name)
        file.save(ppt_file_path)
        guid = ppt_to_image(ppt_file_path)
        os.remove(ppt_file_path)
    except Exception as e:
        print("Error occured in save_ppt_service : ", e)
        return None
    return str(guid)
