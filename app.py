from flask import Flask, request, render_template, send_file, jsonify
import pdfplumber
import pandas as pd
import os
from docx import Document

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
        elif cv.filename.endswith('.docx'):
            text = extract_information_from_docx(cv)
            extracted_text.append(text)    
       
    return extracted_text

@app.route('/')
def upload_cv():
    return render_template('upload_cv.html')

def export_to_excel(data):
    df = pd.DataFrame(data, columns=['Text'])
    # Set the column width for 'Text' column
    df['Text'] = df['Text'].apply(lambda x: ' '.join(str(x).splitlines()))  # Remove line breaks
    excel_filename = 'cv_information.xlsx'
    df.to_excel(excel_filename, index=False)
    # Adjust the column width using openpyxl
    from openpyxl import load_workbook
    wb = load_workbook(excel_filename)
    ws = wb.active
    ws.column_dimensions['A'].width = 100  # Set the width of the 'Text' column
    wb.save(excel_filename)
    return excel_filename

@app.route('/upload', methods=['POST'])
def extract_information():
    cvs = request.files.getlist('cv_files')
    extracted_text = process_cvs(cvs)
    excel_filename = export_to_excel(extracted_text)
    return jsonify({'excel_filename': excel_filename})

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(filename, as_attachment=True)

def extract_information_from_docx(docx_file):
    text = ''
    doc = Document(docx_file)
    for paragraph in doc.paragraphs:
        text += paragraph.text
    return text

if __name__ == '__main__':
    app.run(debug=True)
