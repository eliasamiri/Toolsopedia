from flask import Blueprint, render_template, current_app, jsonify
from controllers.convert import ConvertController, get_job, cancel_job

bp = Blueprint('main', __name__)

@bp.route('/', methods=['GET'])
def index():
    controller = ConvertController(current_app.config)
    return controller.upload_pdf()

@bp.route('/convert', methods=['POST'])
def convert():
    controller = ConvertController(current_app.config)
    return controller.start_job()

@bp.route('/progress/<job_id>', methods=['GET'])
def progress(job_id):
    job = get_job(job_id)
    if not job:
        return jsonify({'error': 'Job not found'}), 404
    return jsonify(job)

@bp.route('/cancel/<job_id>', methods=['POST'])
def cancel(job_id):
    if cancel_job(job_id):
        return jsonify({'status': 'cancelled'})
    return jsonify({'error': 'Job not found'}), 404

@bp.route('/download/<filename>', methods=['GET'])
def download(filename):
    controller = ConvertController(current_app.config)
    return controller.download_file(filename)

@bp.route('/result/<job_id>', methods=['GET'])
def result(job_id):
    controller = ConvertController(current_app.config)
    return controller.render_result(job_id)

def set_routes(app, url_prefix=None):
    app.register_blueprint(bp, url_prefix=url_prefix)