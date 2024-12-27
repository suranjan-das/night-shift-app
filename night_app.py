from flask import Flask, render_template, request, send_file, redirect, url_for, session, Response
import os
from calculation import prepare_apc_sheet, prepare_sap_helper
from calculation_edm import prepare_edm_file, prepare_txt_file, prepare_daily_data_entry
from datetime import datetime
from dotenv import load_dotenv
import threading
import time

load_dotenv()  # Load environment variables from .env file

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY")

USERNAME = os.getenv("USERID")
PASSWORD = os.getenv("PASSWORD")  # Required for session management

generation_status = None  # Global variable to store the generation status

# Function to generate files
def generate_files(date):
    # Your Python program logic to generate files goes here
    if not os.path.exists('generated'):
        os.makedirs('generated')
    # do the calculatoin and make reports
    prepare_apc_sheet(date) # genrates todays APC sheet
    prepare_edm_file(date) # genrates EDM file
    prepare_txt_file(date) # generates profile txt files
    prepare_daily_data_entry(date) # prepare daily entry file
    prepare_sap_helper(date.strftime('%d.%m.%Y')) # prepare sap helper

# Function to run file generation in a thread
def generate_files_thread(date):
    global generation_status
    try:
        generate_files(date)
        generation_status = 'completed'
    except Exception as e:
        generation_status = f'error: {e}'

# Route for login page
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        # get submitted userid and password
        userid = request.form['userid']
        password = request.form['password']
        # authenticate
        if userid == USERNAME and password == PASSWORD:
            session['user'] = userid
            return redirect(url_for('home'))
        else:
            return render_template('login.html', error="Invalid credentials")
    return render_template('login.html')

# Route for logout
@app.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('login'))

# Route for home page
@app.route('/')
def home():
    if 'user' not in session:
        return redirect(url_for('login'))
    error = request.args.get('error')  # Get error message from query parameters
    files_in_generated = os.listdir('generated')
    files_in_uploads = os.listdir('uploads')
    if error:
        return render_template('index.html', files=files_in_generated, uploads_count=len(files_in_uploads), error=error)
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

    # Start the file generation in a new thread
    threading.Thread(target=generate_files_thread, args=(date_obj,)).start()

    return redirect(url_for('home'))

# # Route for generating files
# @app.route('/generate', methods=['POST'])
# def generate():
#     selected_date = request.form['datePicker']
#     generate_option = request.form.get('generateOption')

#     if not generate_option:
#         return redirect(url_for('home', error="No generation option selected"))
#     # Example with the standard date and time format
#     date_format = "%Y-%m-%d"
#     date_obj = datetime.strptime(selected_date, date_format)
#     try:
#         generate_files(date_obj, generate_option)
#     except FileNotFoundError as e:
#         # Handle the error by returning a message or logging it
#         print(f"Error: {e}")
#         return redirect(url_for('home', error=f"Error: {e}"))
#     except Exception as e:
#         return redirect(url_for('home', error=f"Error: {e}"))

#     return redirect(url_for('home'))


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

# Server-Sent Events (SSE) to notify client when generation is complete
@app.route('/status')
def status():
    def generate():
        while True:
            global generation_status
            if generation_status:
                yield f'data: {generation_status}\n\n'
                generation_status = None
                break
            time.sleep(1)

    return Response(generate(), mimetype='text/event-stream')

if __name__ == '__main__':
    # Ensure 'uploads' directory exists
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(debug=True)
