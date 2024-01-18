from flask import Flask, request, jsonify
from flask_cors import CORS  # Import CORS
from werkzeug.utils import secure_filename
import os
import json
from refactor import exec_compare 

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'twbx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/', methods=['GET'])
def home_page():
    return {'status_code': 200, 'status': 'success', 'message': 'Go to the /upload endpoint to compare the file'}


@app.route('/upload', methods=['POST'])
def upload_file():
      # Check if the 'assignmentFile' and 'actualFile' are present in the request files
    if 'assignmentFile' not in request.files or 'actualFile' not in request.files:
        return jsonify({'error': 'Both assignmentFile and actualFile are required'})


    assignment_file = request.files['assignmentFile']
    actual_file = request.files['actualFile']

    if assignment_file.filename == '' or actual_file.filename == '':
        return jsonify({'error': 'One or both of the selected files have no filename'})

    if allowed_file(assignment_file.filename) and allowed_file(actual_file.filename):
        # Save file names
        assignment_filename = secure_filename(assignment_file.filename)
        actual_filename = secure_filename(actual_file.filename)

        assignment_file_path = os.path.join(app.config['UPLOAD_FOLDER'], assignment_filename)
        actual_file_path = os.path.join(app.config['UPLOAD_FOLDER'], actual_filename)

        assignment_file.save(assignment_file_path)
        actual_file.save(actual_file_path)

        # Call your Python script with the file_paths
        result = process_files(assignment_file_path, actual_file_path)

        # Remove the uploaded files after processing if needed
        os.remove(assignment_file_path)
        os.remove(actual_file_path)

        return jsonify(result)

    return jsonify({'error': 'Invalid file format'})



def process_files(assignment_file_path, actual_file_path):
    op = exec_compare(actual_file_path, assignment_file_path)
    return op


if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    app.run(debug=True)
