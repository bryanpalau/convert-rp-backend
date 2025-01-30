from flask import Flask, request, send_file
from flask_cors import CORS
import os
from werkzeug.utils import secure_filename
from docx import Document

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=True, allow_headers=["Content-Type"])

# ✅ Use Vercel's /tmp/ directory
UPLOAD_FOLDER = "/tmp/uploads"
OUTPUT_FOLDER = "/tmp/outputs"

# ✅ Ensure folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return {"error": "No file part"}, 400
    
    file = request.files["file"]
    if file.filename == "":
        return {"error": "No selected file"}, 400
    
    filename = secure_filename(file.filename)
    filepath = os.path.join(UPLOAD_FOLDER, filename)
    file.save(filepath)
    
    processed_filepath = process_report_card(filepath)
    return send_file(processed_filepath, as_attachment=True)

def process_report_card(filepath):
    """ Process the uploaded .docx file and return the modified file path. """
    doc = Document(filepath)

    output_path = os.path.join(OUTPUT_FOLDER, "processed_" + os.path.basename(filepath))
    doc.save(output_path)
    return output_path

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
