from flask import jsonify
from services import presentation_service, word_service, excel_service


def generate_images_controller(request):
    try:
        file = request.files.get("file")
        file_name = file.filename.lower()
        file_extension = file_name.split(".")[-1]
        if file_extension == "ppt" or file_extension == "pptx":
            guid = presentation_service.save_ppt_service(file)

        elif file_extension == "doc" or file_extension == "docx":
            guid = word_service.save_word_file_service(file)

        elif file_extension == "xls" or file_extension == "xlsx":
            guid = excel_service.save_excel_service(file)

        else:
            raise Exception("The file extension is not supported")

        data = {"guid": str(guid)}
        json_data = jsonify(data)
        json_data.status_code = 200
        if not guid:
            raise Exception("guid not generated")
    except Exception as e:
        print("Exception occured in excel_to_image_controller: ", e)
        data = {"error": "could not generate the images"}
        json_data = jsonify(data)
        json_data.status_code = 500
    return json_data


def get_image_controller(request):
    try:
        file_extension = request.json["file_type"]
        guid = request.json["guid"]
        image_number = request.json["image_number"]
        if file_extension == "ppt" or file_extension == "pptx":
            response = presentation_service.send_ppt_image_service(
                guid=guid, slide_no=image_number
            )

        elif file_extension == "doc" or file_extension == "docx":
            response = word_service.send_word_image_service(
                guid=guid, page_number=image_number
            )

        elif file_extension == "xls" or file_extension == "xlsx":
            response = excel_service.send_excel_image_service(
                guid=guid, sheet_number=image_number
            )

        else:
            raise Exception("The file extension is not supported")
        return response
    except Exception as e:
        print("Exception occured in get_image_controller: ", e)
    return None
