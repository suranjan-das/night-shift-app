from flask import Flask, render_template, request, send_file, redirect, url_for
import os
from calculation import prepare_apc_sheet
from calculation_edm import prepare_edm_file, prepare_txt_file, prepare_daily_data_entry
from datetime import datetime

app = Flask(__name__)

# Function to generate files
def generate_files(date):
    # Your Python program logic to generate files goes here
    if not os.path.exists('generated'):
        os.makedirs('generated')
    prepare_apc_sheet(date)
    prepare_edm_file(date)
    prepare_txt_file(date)
    prepare_daily_data_entry(date)

# Route for home page
@app.route('/')
def home():
    files_in_generated = os.listdir('generated')
    files_in_uploads = os.listdir('uploads')
    return render_template('index.html', files=files_in_generated, uploads_count=len(files_in_uploads))

# Route for file upload
@app.route('/upload', methods=['POST'])
def upload():
    uploaded_files = request.files.getlist('file[]')
    # Check if the uploaded_files list is empty
    if not uploaded_files or uploaded_files[0].filename == '':
        return 'No files were uploaded!', 400  # Returning a 400 Bad Request status
    for file in uploaded_files:
        file.save(os.path.join('uploads', file.filename))
    return redirect(url_for('home'))

# Route for file download
@app.route('/download/<path:filename>')
def download(filename):
    return send_file(filename, as_attachment=True)

# Route for generating files
@app.route('/generate', methods=['POST'])
def generate():
    selected_date = request.form['datePicker']
    # Example with the standard date and time format
    date_format = "%Y-%m-%d"
    date_obj = datetime.strptime(selected_date, date_format)
    print(date_obj)
    generate_files(date_obj)
    return redirect(url_for('home'))


# Route for deleting all files
@app.route('/delete_all')
def delete_all():
    files = os.listdir('generated')
    for file in files:
        os.remove(os.path.join('generated', file))
    uploaded_files = os.listdir('uploads')
    for file in uploaded_files:
        os.remove(os.path.join('uploads', file))
    return redirect(url_for('home'))

if __name__ == '__main__':
    # Ensure 'uploads' directory exists
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(debug=True)
