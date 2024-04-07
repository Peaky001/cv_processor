from flask import Flask, request, render_template, send_file
import pdfplumber
import pandas as pd
import os

app = Flask(__name__)

def extract_information_from_pdf(pdf_file):
    text = ''
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text += page.extract_text()
    return text

def process_cvs(cvs):
    extracted_text = []
    for cv in cvs:
        if cv.filename.endswith('.pdf'):
            text = extract_information_from_pdf(cv)
            extracted_text.append(text)
        # Add handling for other file types if necessary
    return extracted_text

@app.route('/')
def upload_cv():
    return render_template('upload_cv.html')

def export_to_excel(data):
    df = pd.DataFrame(data, columns=['Text'])
    df.to_excel('cv_information.xlsx', index=False)
    return send_file('cv_information.xlsx', as_attachment=True)

@app.route('/upload', methods=['POST'])
def extract_information():
    cvs = request.files.getlist('cv_files')
    extracted_text = process_cvs(cvs)
    return export_to_excel(extracted_text)

if __name__ == '__main__':
    app.run(debug=True)
