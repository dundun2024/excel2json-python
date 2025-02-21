import os

from flask import Flask, request, jsonify, render_template
import json
import pandas as pd
import webbrowser

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

def excel_to_json(file_path):
    df = pd.read_excel(file_path)
    return df.to_json(orient='records')

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    try:
        json_data = excel_to_json(file_path)
        return json.dumps(json.loads(json_data), ensure_ascii=False)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    webbrowser.open('http://localhost:5000')
    app.run(threaded=True)

