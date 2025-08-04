from flask import Flask, request, render_template, send_file
from PIL import Image
import os
import tempfile
from PyPDF2 import PdfMerger
from pdf2docx import Converter
import pandas as pd
import pdfplumber

app = Flask(__name__)

# A4 dimensions in pixels (at 72 DPI)
A4_WIDTH_PX = 595
A4_HEIGHT_PX = 842

# ----------- UI ROUTES -----------

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/jpg-to-pdf')
def jpg_to_pdf_page():
    return render_template('jpg_to_pdf.html')

@app.route('/merge-jpg')
def merge_jpg_page():
    return render_template('merge_jpg.html')

@app.route('/merge-pdf')
def merge_pdf_page():
    return render_template('merge_pdf.html')

@app.route('/pdf-to-word')
def pdf_to_word_page():
    return render_template('pdf_to_word.html')

@app.route('/pdf-to-excel')
def pdf_to_excel_page():
    return render_template('pdf_to_excel.html')

# ----------- PROCESSING ROUTES -----------

@app.route('/convert-jpg-to-pdf', methods=['POST'])
def convert_jpg_to_pdf():
    if 'file' not in request.files:
        return "No file uploaded", 400

    file = request.files['file']
    if file.filename == '':
        return "No file selected", 400

    if file and (file.filename.endswith('.jpg') or file.filename.endswith('.jpeg')):
        img = Image.open(file)

        # Resize and center on A4
        img.thumbnail((A4_WIDTH_PX, A4_HEIGHT_PX))
        a4_img = Image.new('RGB', (A4_WIDTH_PX, A4_HEIGHT_PX), (255, 255, 255))
        img_position = ((A4_WIDTH_PX - img.width) // 2, (A4_HEIGHT_PX - img.height) // 2)
        a4_img.paste(img, img_position)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
            pdf_path = temp_file.name
            a4_img.save(pdf_path, "PDF", resolution=100.0)

        return send_file(pdf_path, as_attachment=True, download_name=f"{os.path.splitext(file.filename)[0]}.pdf")
    else:
        return "Invalid file format. Please upload a JPG image.", 400


@app.route('/merge-jpg-to-pdf', methods=['POST'])
def merge_jpg_to_pdf():
    files = request.files.getlist('files')
    images = []

    for file in files:
        if file.filename.endswith(('.jpg', '.jpeg')):
            img = Image.open(file)
            img.thumbnail((A4_WIDTH_PX, A4_HEIGHT_PX))
            a4_img = Image.new('RGB', (A4_WIDTH_PX, A4_HEIGHT_PX), (255, 255, 255))
            img_position = ((A4_WIDTH_PX - img.width) // 2, (A4_HEIGHT_PX - img.height) // 2)
            a4_img.paste(img, img_position)
            images.append(a4_img)

    if images:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
            pdf_path = temp_file.name
            images[0].save(pdf_path, save_all=True, append_images=images[1:], resolution=100.0)

        return send_file(pdf_path, as_attachment=True, download_name="merged_images.pdf")
    else:
        return "Invalid file format. Please upload JPG images.", 400


@app.route('/merge-pdf', methods=['POST'])
def merge_pdf():
    files = request.files.getlist('files')
    merger = PdfMerger()

    for file in files:
        if file.filename.endswith('.pdf'):
            merger.append(file)

    if merger.pages:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
            pdf_path = temp_file.name
            merger.write(pdf_path)
            merger.close()

        return send_file(pdf_path, as_attachment=True, download_name="merged_document.pdf")
    else:
        return "Invalid file format. Please upload PDF documents.", 400


@app.route('/convert-pdf-to-word', methods=['POST'])
def convert_pdf_to_word():
    if 'file' not in request.files:
        return "No file uploaded", 400

    file = request.files['file']
    if file.filename == '':
        return "No file selected", 400

    if file and file.filename.endswith('.pdf'):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf_file:
            pdf_path = temp_pdf_file.name
            file.save(pdf_path)

        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_word_file:
                docx_path = temp_word_file.name
                cv = Converter(pdf_path)
                cv.convert(docx_path)
                cv.close()

            return send_file(docx_path, as_attachment=True, download_name=f"{os.path.splitext(file.filename)[0]}.docx")
        finally:
            os.remove(pdf_path)
    else:
        return "Invalid file format. Please upload a PDF document.", 400


@app.route('/convert-pdf-to-excel', methods=['POST'])
def convert_pdf_to_excel():
    if 'file' not in request.files:
        return "No file uploaded", 400

    file = request.files['file']
    if file.filename == '':
        return "No file selected", 400

    if file and file.filename.endswith('.pdf'):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf_file:
            pdf_path = temp_pdf_file.name
            file.save(pdf_path)

        try:
            with pdfplumber.open(pdf_path) as pdf:
                all_tables = []
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        df = pd.DataFrame(table[1:], columns=table[0])
                        all_tables.append(df)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_excel_file:
                excel_path = temp_excel_file.name
                with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                    for i, df in enumerate(all_tables):
                        df.to_excel(writer, index=False, sheet_name=f'Sheet{i+1}')

            os.remove(pdf_path)
            return send_file(excel_path, as_attachment=True, download_name=f"{os.path.splitext(file.filename)[0]}.xlsx")

        except Exception as e:
            return str(e), 500

    else:
        return "Invalid file format. Please upload a PDF document.", 400


if __name__ == '__main__':
    app.run(debug=True)
