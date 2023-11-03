import os
import fitz
import win32com.client
import pythoncom
import uuid
from flask import Response
import io


def word_to_image_service(file_name, word_file_path):
    try:
        word_app = win32com.client.Dispatch(
            "Word.Application", pythoncom.CoInitialize()
        )
        document = word_app.Documents.Open(word_file_path)
        guid = uuid.uuid4()
        image_folder_name = str(guid) + "_word_images"
        word_to_pdf_folder_path = os.path.abspath("temp_word_to_pdf_folder")
        pdf_file_name = str(file_name) + str(guid) + ".pdf"
        pdf_file_path = os.path.join(word_to_pdf_folder_path, pdf_file_name)
        document.ExportAsFixedFormat(pdf_file_path, 17)
        document.Close()
        word_app.Quit()
        word_app = None
        word_image_folder = os.path.abspath("word_images")
        image_folder_path = os.path.join(word_image_folder, image_folder_name)
        if not os.path.exists(image_folder_path):
            os.makedirs(image_folder_path)
        doc = fitz.open(pdf_file_path)
        zoom = 1
        mat = fitz.Matrix(zoom, zoom)
        pages = len(doc)
        for page_number in range(pages):
            image_name = "image_" + str(page_number + 1) + ".png"
            image_path = os.path.join(image_folder_path, image_name)
            page = doc.load_page(page_number)
            image = page.get_pixmap(matrix=mat)
            image.save(image_path)
        doc.close()
        os.remove(pdf_file_path)
        return str(guid)
    except Exception as e:
        print("Exception occured in word_to_image_service : ", e)
        document.Close()
        word_app.Quit()
        word_app = None

    return None


def save_word_file_service(file):
    try:
        # breakpoint()
        temp_word_folder_path = os.path.abspath("temp_word_folder")
        file_name = file.filename
        word_file_path = os.path.join(temp_word_folder_path, file_name)
        file.save(word_file_path)
        guid = word_to_image_service(file_name=file_name, word_file_path=word_file_path)
        os.remove(word_file_path)
    except Exception as e:
        print("Error occured in save_word_file_service : ", e)
        return None
    return str(guid)


def send_word_image_service(guid, page_number):
    try:
        word_image_path = os.path.abspath("word_images")
        image_folder_name = str(guid) + "_word_images"
        image_folder_path = os.path.join(word_image_path, image_folder_name)
        image_name = "image_" + str(page_number) + ".png"
        with open(os.path.join(image_folder_path, image_name), "rb") as image_file:
            image_data = io.BytesIO(image_file.read())
            response = Response(image_data, mimetype="image/png")
            response.headers["Content-Type"] = "image/png"
        return response
    except Exception as e:
        print("Exception occured in send_word_image_service : ", e)
    return None
