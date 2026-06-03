import os
import time
import threading
import uuid
from concurrent.futures import ProcessPoolExecutor, as_completed
from flask import request, render_template, send_from_directory, url_for, jsonify
from werkzeug.utils import secure_filename
from services.converter import ConverterService, convert_range_worker, merge_workbooks

ALLOWED_EXTENSIONS = {'pdf'}

class ConvertController:
    def __init__(self, config):
        self.upload_folder = config['UPLOAD_FOLDER']
        self.converted_folder = config['CONVERTED_FOLDER']
        self.converter = ConverterService()

    def _allowed(self, filename):
        return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

    def start_job(self):
        if 'file' not in request.files:
            return jsonify({'error': 'No file part in the request.'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected.'}), 400
        if not self._allowed(file.filename):
            return jsonify({'error': 'Only PDF files are allowed.'}), 400

        filename = secure_filename(file.filename)
        pdf_path = os.path.join(self.upload_folder, filename)
        file.save(pdf_path)

        excel_name = os.path.splitext(filename)[0] + '.xlsx'
        excel_path = os.path.join(self.converted_folder, excel_name)

        sheet_mode = request.form.get('sheet_mode', 'multi')
        single_sheet = sheet_mode == 'single'

        analysis = self.converter.analyze(pdf_path)
        total_pages = analysis['page_count']

        job_id = str(uuid.uuid4())
        _set_job(job_id, {
            'status': 'processing',
            'percent': 0,
            'current_page': 0,
            'total_pages': total_pages,
            'filename': excel_name,
            'download_url': url_for('main.download', filename=excel_name),
            'share_url': request.host_url.rstrip('/') + url_for('main.download', filename=excel_name),
            'elapsed_seconds': None,
            'error': None,
            'cancelled': False,
            'single_sheet': single_sheet,
        })

        thread = threading.Thread(
            target=_run_job,
            args=(job_id, self.converter, pdf_path, excel_path, single_sheet),
            daemon=True,
        )
        thread.start()

        return jsonify({'job_id': job_id}), 202

    def upload_pdf(self):
        return render_template('index.html')

    def download_file(self, filename):
        return send_from_directory(self.converted_folder, filename, as_attachment=True)

    def render_result(self, job_id):
        job = get_job(job_id)
        if not job or job.get('status') != 'done':
            return render_template('index.html', error='Conversion not completed yet.')

        return render_template(
            'download.html',
            filename=job['filename'],
            download_url=job['download_url'],
            share_url=job['share_url'],
            elapsed_seconds=job['elapsed_seconds'],
            page_count=job['total_pages'],
        )


JOBS = {}
JOBS_LOCK = threading.Lock()


def _set_job(job_id, data):
    with JOBS_LOCK:
        JOBS[job_id] = data


def _update_job(job_id, updates):
    with JOBS_LOCK:
        JOBS[job_id].update(updates)


def get_job(job_id):
    with JOBS_LOCK:
        return JOBS.get(job_id)


def _run_job(job_id, converter, pdf_path, excel_path, single_sheet):
    start_time = time.perf_counter()

    def progress_cb(current_page, total_pages):
        if get_job(job_id).get('cancelled'):
            raise RuntimeError('Conversion cancelled')
        percent = int((current_page / max(1, total_pages)) * 100)
        _update_job(job_id, {
            'current_page': current_page,
            'total_pages': total_pages,
            'percent': percent,
        })

    try:
        analysis = converter.analyze(pdf_path)
        total_pages = analysis['page_count']
        chunk_size = 100
        max_workers = max(2, min(4, os.cpu_count() or 2))

        if total_pages <= chunk_size:
            converter.convert(pdf_path, excel_path, progress_cb=progress_cb, single_sheet=single_sheet)
        else:
            chunks = []
            start = 1
            while start <= total_pages:
                end = min(start + chunk_size - 1, total_pages)
                chunks.append((start, end))
                start = end + 1

            part_paths = []
            args_list = []
            for index, (start_page, end_page) in enumerate(chunks, start=1):
                part_path = f"{excel_path}.part_{index}.xlsx"
                part_paths.append(part_path)
                args_list.append((pdf_path, part_path, start_page, end_page, single_sheet))

            completed_pages = 0
            with ProcessPoolExecutor(max_workers=max_workers) as executor:
                future_map = {executor.submit(convert_range_worker, args): args for args in args_list}
                for future in as_completed(future_map):
                    if get_job(job_id).get('cancelled'):
                        executor.shutdown(cancel_futures=True)
                        _update_job(job_id, {
                            'status': 'cancelled',
                        })
                        return
                    _ = future.result()
                    start_page, end_page = future_map[future][2], future_map[future][3]
                    completed_pages += (end_page - start_page + 1)
                    progress_cb(completed_pages, total_pages)

            if get_job(job_id).get('cancelled'):
                _update_job(job_id, {
                    'status': 'cancelled',
                })
                return

            merge_workbooks(part_paths, excel_path, single_sheet=single_sheet)
            for path in part_paths:
                if os.path.exists(path):
                    os.remove(path)

        elapsed_seconds = round(time.perf_counter() - start_time, 2)
        _update_job(job_id, {
            'status': 'done',
            'percent': 100,
            'elapsed_seconds': elapsed_seconds,
        })
    except Exception as e:
        if str(e) == 'Conversion cancelled':
            _update_job(job_id, {
                'status': 'cancelled',
            })
            return
        _update_job(job_id, {
            'status': 'error',
            'error': str(e),
        })


def cancel_job(job_id):
    if job_id in JOBS:
        _update_job(job_id, {
            'cancelled': True,
            'status': 'cancelled',
        })
        return True
    return False