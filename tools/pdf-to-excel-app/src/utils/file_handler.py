import os
from flask import request, send_from_directory

def save_file(upload_folder):
    if 'file' not in request.files:
        return None
    file = request.files['file']
    if file.filename == '':
        return None
    file_path = os.path.join(upload_folder, file.filename)
    file.save(file_path)
    return file_path

def get_converted_file_path(filename, converted_folder):
    return os.path.join(converted_folder, filename)