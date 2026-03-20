import json

from flask import send_file
import os
from mainframe import generate_test_cases, upload_test_cases_ado
from flask import Flask
from flask_cors import CORS


app = Flask(__name__)
os.makedirs("output", exist_ok=True)
CORS(app)


@app.route('/generate', methods=['POST'])
def generate():
    try:
        if 'file' not in request.files or request.files['file'].filename == '':
            return jsonify({'status': 'error', 'message': 'No file provided'}), 400
        file = request.files['file']
        if not file.filename.endswith('.xlsx'):
            return jsonify({'status': 'error', 'message': 'Only .xlsx allowed'}), 400
        file.save("temp_input.xlsx")
        image_path = None
        if 'image' in request.files and request.files['image'].filename != '':
            image_file = request.files['image']
            file_ext = os.path.splitext(secure_filename(image_file.filename))[1]
            image_path = f"temp_image{file_ext}"
            image_file.save(image_path)
        output_file = generate_test_cases("temp_input.xlsx", image_path=image_path)
        count = len(pd.read_excel(output_file))
        if image_path and os.path.exists(image_path):
            os.remove(image_path)
        os.remove("temp_input.xlsx")
        return jsonify({
            'status': 'success',
            'count': count,
            'filename': os.path.basename(output_file)
        }), 200
    except Exception as e:
        if 'image_path' in locals() and image_path and os.path.exists(image_path):
            os.remove(image_path)
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/download', methods=['GET'])
def download():
    try:
        filename = request.args.get('filename')
        if not filename: return jsonify({'status': 'error', 'message': 'Filename required'}), 400
        file_path = os.path.join("output", secure_filename(filename))
        if not os.path.exists(file_path): return jsonify({'status': 'error', 'message': 'File not found'}), 404
        return send_file(file_path, as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/upload', methods=['POST'])
def upload():
    try:
        if 'file' not in request.files or request.files['file'].filename == '':
            return jsonify({'status': 'error', 'message': 'No file provided'}), 400
        file = request.files['file']
        if not file.filename.endswith('.xlsx'):
            return jsonify({'status': 'error', 'message': 'Only .xlsx allowed'}), 400

        org = request.form.get('org')
        proj = request.form.get('project')
        pat = request.form.get('pat')
        plan_name = request.form.get('plan_name')
        suite_name = request.form.get('suite_name', 'LOGIN')

        if not all([org, proj, pat, plan_name]):
            return jsonify(
                {'status': 'error', 'message': 'Missing ADO config: org, project, pat, plan_name required'}), 400

        file.save("temp_output.xlsx")
        upload_count, error_count = upload_test_cases_ado("temp_output.xlsx", org, proj, pat, plan_name, suite_name)
        os.remove("temp_output.xlsx")
        return jsonify({'status': 'success' if upload_count > 0 else 'fail', 'uploaded_count': upload_count,
                        'failed_count': error_count}), 200
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@app.route('/download-template', methods=['GET'])
def download_template():
    try:
        if not os.path.exists('template.xlsx'):
            return jsonify({'status': 'error', 'message': 'Template not found'}), 404
        return send_file('template.xlsx', as_attachment=True, download_name='template.xlsx')
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

from collections import OrderedDict

from collections import OrderedDict
import pandas as pd, os
from flask import jsonify, request
from werkzeug.utils import secure_filename

@app.route('/get-test-cases', methods=['GET'])
def get_test_cases():
    try:
        filename = request.args.get('filename')
        if not filename:
            return jsonify({'status': 'error', 'message': 'Filename required'}), 400
        file_path = os.path.join("output", secure_filename(filename))
        if not os.path.exists(file_path):
            return jsonify({'status': 'error', 'message': 'File not found'}), 404
        df = pd.read_excel(file_path)
        preferred_order = [
            'S.No.',
            'User Story',
            'Acceptance Criteria',
            'Title',
            'Steps',
            'Priority',
            'Test Type'
        ]
        existing_cols = [c for c in preferred_order if c in df.columns]
        df = df.where(pd.notnull(df), None)
        test_cases = []
        for _, row in df.iterrows():
            ordered_row = OrderedDict()
            for col in existing_cols:
                ordered_row[col] = row[col]
            test_cases.append(ordered_row)
        response = {'status': 'success', 'test_cases': test_cases}
        return app.response_class(
            response=json.dumps(response, ensure_ascii=False),
            mimetype='application/json'
        )
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)