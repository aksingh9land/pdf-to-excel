import os
import pdfplumber
import pytesseract
import pdf2image
import pandas as pd
from flask import Flask, request, jsonify, send_file
from werkzeug.utils import secure_filename

# Set up Flask app
app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
RESULTS_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULTS_FOLDER, exist_ok=True)

# Tesseract Path (change this if needed)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# PDF to Text (OCR)
def ocr_pdf(pdf_path):
    images = pdf2image.convert_from_path(pdf_path)
    extracted_text = ""
    for img in images:
        extracted_text += pytesseract.image_to_string(img, lang="eng") + "\n"
    return extracted_text.strip()

# PDF to Excel
def pdf_to_excel(pdf_path, output_excel):
    with pdfplumber.open(pdf_path) as pdf:
        all_tables = []
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                df = pd.DataFrame(table)
                all_tables.append(df)
        
        if all_tables:
            final_df = pd.concat(all_tables, ignore_index=True)
            final_df.to_excel(output_excel, index=False)
            return True
    return False

# Upload API
@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No selected file"}), 400

    filename = secure_filename(file.filename)
    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(pdf_path)

    # OCR & Excel Processing
    text_output = os.path.join(RESULTS_FOLDER, f"{filename}_text.txt")
    excel_output = os.path.join(RESULTS_FOLDER, f"{filename}.xlsx")

    extracted_text = ocr_pdf(pdf_path)
    with open(text_output, "w", encoding="utf-8") as f:
        f.write(extracted_text)

    excel_success = pdf_to_excel(pdf_path, excel_output)

    response_data = {
        "message": "File processed successfully",
        "text_download_url": f"/download/text/{filename}_text.txt",
        "excel_download_url": f"/download/excel/{filename}.xlsx" if excel_success else None
    }
    return jsonify(response_data)

# Download OCR Text
@app.route("/download/text/<filename>", methods=["GET"])
def download_text(filename):
    return send_file(os.path.join(RESULTS_FOLDER, filename), as_attachment=True)

# Download Excel File
@app.route("/download/excel/<filename>", methods=["GET"])
def download_excel(filename):
    return send_file(os.path.join(RESULTS_FOLDER, filename), as_attachment=True)

# Run App
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
