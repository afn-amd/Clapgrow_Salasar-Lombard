# app.py
from flask import Flask, request, send_file, send_from_directory
import os
from processing import process_excel

app = Flask(__name__, static_folder='static')

@app.route('/', methods = ['GET'])
def mainPage():
    #return url_for('static', filename='index.html')
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    print(request.files)
    if 'file1' not in request.files or 'file2' not in request.files:
        return 'No files uploaded', 400

    file1 = request.files['file1']
    file2 = request.files['file2']

    # Save files temporarily
    file1_path = os.path.join('temp', file1.filename)
    file2_path = os.path.join('temp', file2.filename)
    print(file1_path, file2_path)
    os.makedirs('temp', exist_ok=True)  # Create temp directory if it doesn't exist
    file1.save(file1_path)
    file2.save(file2_path)

    # Process files
    output_file = process_excel(file1_path, file2_path)

    # Send the output file back
    tempp = send_file(output_file, as_attachment=True)
    #os.remove(output_file)
    return tempp

if __name__ == '__main__':
    app.run(debug=True)