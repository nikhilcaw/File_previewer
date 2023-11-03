from flask import Flask, request
from controllers import controller

app = Flask(__name__)


@app.route("/upload", methods=["POST"])
def file_upload():
    return controller.generate_images_controller(request)


@app.route("/get_image", methods=["GET"])
def get_images():
    return controller.get_image_controller(request=request)


if __name__ == "__main__":
    app.run(debug=True, port=3000)
