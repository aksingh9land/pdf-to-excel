import os
import pdfplumber
import pandas as pd
import threading
import uuid  # Import UUID to generate unique file names
from flask import Flask, request, send_file, jsonify, session
from flask_session import Session  # For user session management

app = Flask(__name__)

# Configure session to use filesystem
app.config["SESSION_TYPE"] = "filesystem"
app.config["SESSION_PERMANENT"] = False
app.config["SESSION_FILE_DIR"] = "./sessions"
Session(app)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/upload", methods=["POST"])
def upload_pdf():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]

    # Generate a unique ID for each file
    unique_id = str(uuid.uuid4())  # Example: "f47ac10b-58cc-4372-a567-0e02b2c3d479"
    session["file_id"] = unique_id  # Store the file ID in the user's session

    filename = os.path.join(UPLOAD_FOLDER, f"{unique_id}.pdf")
    file.save(filename)

    # Convert PDF to Excel with a unique output file
    excel_file = os.path.join(OUTPUT_FOLDER, f"{unique_id}.xlsx")
    pdf_to_excel(filename, excel_file)

    return jsonify({
        "excel_download_url": f"https://pdf-to-excel-xpfz.onrender.com/download/excel"
    })

@app.route("/download/excel", methods=["GET"])
def download_excel():
    file_id = request.cookies.get("file_id")  # Get file_id from cookie

    if not file_id:
        print("⚠️ No file_id found in cookies!")
        return jsonify({"error": "No file found"}), 404

    excel_file = os.path.join(OUTPUT_FOLDER, f"{file_id}.xlsx")

    if os.path.exists(excel_file):
        print("✅ File found:", excel_file)
        
        # Create response with file
        response = send_file(excel_file, as_attachment=True)

        # Schedule file deletion after response is sent
        def delete_file():
            print(f"🗑️ Deleting file: {excel_file}")
            os.remove(excel_file)

        # Run delete file function after 5 seconds (to allow browser time to download)
        threading.Timer(5, delete_file).start()

        return response
    else:
        print("❌ File not found:", excel_file)
        return jsonify({"error": "File not found"}), 404

def pdf_to_excel(pdf_path, output_excel):
    data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                for row in table:
                    data.append(row)

    if data:
        df = pd.DataFrame(data)
        df.to_excel(output_excel, index=False)
    else:
        print("No tables found in the PDF")

if __name__ == "__main__":
    app.run(debug=True)
