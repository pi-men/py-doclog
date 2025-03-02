import os
import pdfplumber
import pytesseract
import pandas as pd
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from PIL import Image

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Set path Tesseract jika diperlukan (Windows)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_text_from_pdf(pdf_path):
    """Fungsi untuk ekstraksi teks dari PDF dengan OCR"""
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Ambil teks langsung dari PDF jika tersedia
            text += page.extract_text() or ""

            # OCR untuk gambar dalam PDF
            for img in page.images:
                image = Image.open(img['stream'])
                text += pytesseract.image_to_string(image)

    return text.strip()

def extract_tables_from_pdf(pdf_path):
    """Fungsi untuk ekstraksi tabel dari PDF ke DataFrame"""
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            extracted_tables = page.extract_tables()
            for table in extracted_tables:
                df = pd.DataFrame(table)
                tables.append(df)
    return tables

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "file" not in request.files:
            return "No file uploaded", 400

        file = request.files["file"]
        if file.filename == "":
            return "No selected file", 400

        if file:
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)

            # Ekstraksi teks dan tabel
            extracted_text = extract_text_from_pdf(filepath)
            tables = extract_tables_from_pdf(filepath)

            # Simpan tabel ke file Excel jika ada tabel
            excel_path = None
            if tables:
                excel_path = os.path.join(app.config['UPLOAD_FOLDER'], "extracted_data.xlsx")
                with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                    for i, df in enumerate(tables):
                        df.to_excel(writer, sheet_name=f"Sheet{i+1}", index=False)

            return render_template("index.html", text=extracted_text, excel_file=excel_path)

    return render_template("index.html", text=None, excel_file=None)

@app.route("/download")
def download():
    """Fungsi untuk mengunduh file Excel hasil ekstraksi tabel"""
    excel_path = os.path.join(app.config['UPLOAD_FOLDER'], "extracted_data.xlsx")
    if os.path.exists(excel_path):
        return send_file(excel_path, as_attachment=True)
    return "No file available", 404

if __name__ == "__main__":
    app.run(debug=True)
