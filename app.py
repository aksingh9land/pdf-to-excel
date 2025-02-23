import os
import pdfplumber
import pytesseract
import pdf2image
import pandas as pd
from flask import Flask, request, send_file, jsonify

app = Flask(__name__)

# Define folders
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/upload", methods=["POST"])
def upload_pdf():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    filename = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(filename)

    # Extract text from PDF
    text_file = os.path.join(OUTPUT_FOLDER, "output.txt")
    extracted_text = extract_text_from_pdf(filename)
    with open(text_file, "w", encoding="utf-8") as f:
        f.write(extracted_text)

    # Convert PDF tables to Excel
    excel_file = os.path.join(OUTPUT_FOLDER, "output.xlsx")
    pdf_to_excel(filename, excel_file)

    return jsonify({
        "text_download_url": f"https://pdf-to-excel-xpfz.onrender.com/download/text",
        "excel_download_url": f"https://pdf-to-excel-xpfz.onrender.com/download/excel"
    })

@app.route("/download/text", methods=["GET"])
def download_text():
    text_file = os.path.join(OUTPUT_FOLDER, "output.txt")
    if os.path.exists(text_file):
        return send_file(text_file, as_attachment=True)
    else:
        return jsonify({"error": "File not found"}), 404

@app.route("/download/excel", methods=["GET"])
def download_excel():
    excel_file = os.path.join(OUTPUT_FOLDER, "output.xlsx")
    if os.path.exists(excel_file):
        return send_file(excel_file, as_attachment=True)
    else:
        return jsonify({"error": "File not found"}), 404

def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    return text

def pdf_to_excel(pdf_path, output_excel):
    data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    data.append(row)

    df = pd.DataFrame(data)
    df.to_excel(output_excel, index=False)

if __name__ == "__main__":
    app.run(debug=True)
