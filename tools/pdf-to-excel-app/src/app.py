import os
import sys

# Allow imports from src/
sys.path.insert(0, os.path.dirname(__file__))

from flask import Flask, send_from_directory
from routes.index import set_routes

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..', '..'))
LANDING_DIR = PROJECT_ROOT
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
CONVERTED_FOLDER = os.path.join(BASE_DIR, 'converted')

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

app = Flask(
    __name__,
    template_folder=os.path.join(BASE_DIR, 'templates'),
    static_folder=os.path.join(BASE_DIR, 'static'),
)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB limit

@app.route('/')
def landing_page():
    return send_from_directory(LANDING_DIR, 'index.html')

@app.route('/styles.css')
def landing_styles():
    return send_from_directory(LANDING_DIR, 'styles.css')

set_routes(app, url_prefix='/tools/pdf-to-excel-app')

if __name__ == '__main__':
    app.run(debug=True, port=5000)