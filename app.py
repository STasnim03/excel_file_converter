from flask import Flask, request, redirect, url_for, send_file, render_template, flash
import pandas as pd
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
app.secret_key = 'supersecretkey'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx'}

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def validate_excel(file_path, required_columns):
    df = pd.read_excel(file_path)
    if not all(column in df.columns for column in required_columns):
        return False, df
    return True, df

@app.route('/')
def index():
    return render_template('index.html', result_file=None)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file1' not in request.files or 'file2' not in request.files:
        flash('No file part')
        return redirect(request.url)

    file1 = request.files['file1']
    file2 = request.files['file2']

    if file1.filename == '' or file2.filename == '':
        flash('No selected file')
        return redirect(request.url)

    if not allowed_file(file1.filename) or not allowed_file(file2.filename):
        flash('File type not allowed')
        return redirect(request.url)

    filename1 = secure_filename(file1.filename)
    file1_path = os.path.join(app.config['UPLOAD_FOLDER'], filename1)
    file1.save(file1_path)

    filename2 = secure_filename(file2.filename)
    file2_path = os.path.join(app.config['UPLOAD_FOLDER'], filename2)
    file2.save(file2_path)

    required_columns_file1 = ['Item Code', 'Item Name']
    required_columns_file2 = ['Item Code']

    valid_file1, df1 = validate_excel(file1_path, required_columns_file1)
    valid_file2, df2 = validate_excel(file2_path, required_columns_file2)

    if not valid_file1 or not valid_file2:
        flash(f'File is missing required columns')
        return redirect(url_for('index'))

    # Process the files
    code_to_item = pd.Series(df1['Item Name'].values, index=df1['Item Code']).to_dict()
    df2['Item Name'] = df2['Item Code'].map(code_to_item)

    output_file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Final_File.xlsx')
    df2.to_excel(output_file_path, index=False)

    return render_template('index.html', result_file='Final_File.xlsx')

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
