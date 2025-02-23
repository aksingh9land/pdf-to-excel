import os
import pdfplumber
import pytesseract
import pdf2image
import pandas as pd
from flask import Flask, request, send_file, jsonify
from PIL import Image

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

    # Check file type
    if file.filename.lower().endswith(".pdf"):
        extracted_text = extract_text_from_pdf(filename)
    elif file.filename.lower().endswith((".png", ".jpg", ".jpeg")):
        extracted_text = extract_text_from_image(filename)
    else:
        return jsonify({"error": "Unsupported file format"}), 400

    # Save extracted text as a file
    text_file = os.path.join(OUTPUT_FOLDER, "output.txt")
    with open(text_file, "w", encoding="utf-8") as f:
        f.write(extracted_text)

    # Convert extracted text into Excel
    excel_file = os.path.join(OUTPUT_FOLDER, "output.xlsx")
    convert_text_to_excel(extracted_text, excel_file)

    return jsonify({
        "text_download_url": "https://pdf-to-excel-xpfz.onrender.com/download/text",
        "excel_download_url": "https://pdf-to-excel-xpfz.onrender.com/download/excel"
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
            page_text = page.extract_text()
            if page_text:  # If text is found, use it
                text += page_text + "\n"
            else:  # If no text is found, use OCR
                text += extract_text_from_image(pdf_path) + "\n"
    return text.strip()

def extract_text_from_image(image_path):
    # If it's a PDF, convert it to images
    if image_path.lower().endswith(".pdf"):
        images = pdf2image.convert_from_path(image_path)
    else:
        images = [Image.open(image_path)]

    extracted_text = ""
    for img in images:
        extracted_text += pytesseract.image_to_string(img) + "\n"
    return extracted_text.strip()

def convert_text_to_excel(text, output_excel):
    lines = text.split("\n")
    df = pd.DataFrame({"Extracted Text": lines})
    df.to_excel(output_excel, index=False)

if __name__ == "__main__":
    app.run(debug=True)
