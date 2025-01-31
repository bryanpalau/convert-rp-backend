from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import re
import sys
import logging
import tempfile
from datetime import datetime
from pathlib import Path
from werkzeug.utils import secure_filename
from docx import Document
from docx.shared import Pt
import pdfkit
from pdf2docx import Converter

# Configure detailed logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization"
    return response

TEMP_DIR = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'docx', 'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_pdf_to_docx(pdf_path):
    docx_path = pdf_path.replace(".pdf", ".docx")
    cv = Converter(pdf_path)
    cv.convert(docx_path, start=0, end=None)
    cv.close()
    return docx_path

def clean_course_title(title):
    if not title or 'Study Hall' in title:
        return ''
    
    has_plus = title.startswith('+')
    title = title[1:].strip() if has_plus else title.strip()
    
    patterns = [
        r'\bG\d+(-\d+)?(?=\s|$|\()',  
        r'\b(?:Grade\s*)?\d{1,2}(?:th)?\s*(?:Grade\s*)?(?:-\d+)?(?=\s|$)',  
        r'^(?:Senior|Junior)?\s*Electives\s*\d*-?',
        r'^(?:Math|Science|Career Planning|Visual Performing Arts|Foreign Language|Military Training|Social Studies)-',
        r'\s*Group\s*\d+\s*-',
        r'-\d+(\s|$)',
        r'\s*\([^)]*\)',
        r'\s*-\s*(?=\S)',
        r'\s+-\s*$',
    ]
    
    for pattern in patterns:
        title = re.sub(pattern, '', title, flags=re.IGNORECASE)
    
    title = re.sub(r'\s+', ' ', title).strip('- ')
    
    return f"+ {title}" if has_plus else title

def process_table(table):
    seen_courses = set()
    for row in table.rows[1:]:
        cells = row.cells
        if len(cells) >= 3:
            cleaned_title = clean_course_title(cells[0].text.strip())
            grade = cells[1].text.strip()
            gpa = cells[2].text.strip()
            if not cleaned_title:
                continue
            
            course_key = (cleaned_title, grade, gpa)
            if course_key in seen_courses:
                table._element.getparent().remove(row._element)
            else:
                seen_courses.add(course_key)
                cells[0].text = cleaned_title

def process_report_card(filepath, output_format='docx'):
    try:
        if filepath.endswith(".pdf"):
            filepath = convert_pdf_to_docx(filepath)
        
        doc = Document(filepath)
        for table in doc.tables:
            process_table(table)
        
        output_path = os.path.join(TEMP_DIR, f"processed_{os.path.basename(filepath)}")
        doc.save(output_path)
        
        if output_format == 'pdf':
            pdf_output_path = output_path.replace('.docx', '.pdf')
            pdfkit_config = pdfkit.configuration(wkhtmltopdf='/usr/local/bin/wkhtmltopdf' if os.path.exists('/usr/local/bin/wkhtmltopdf') else '/opt/homebrew/bin/wkhtmltopdf')
            pdfkit.from_file(output_path, pdf_output_path, configuration=pdfkit_config)
            return pdf_output_path
        
        return output_path
    except Exception as e:
        logger.error(f"Error processing document: {str(e)}")
        return None

@app.route("/")
def home():
    return "Flask backend is running!", 200

@app.route("/favicon.ico")
def favicon():
    return "", 204

@app.route("/upload", methods=["POST"])
def upload_file():
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file part"}), 400
        
        file = request.files["file"]
        if file.filename == "" or not allowed_file(file.filename):
            return jsonify({"error": "Invalid file type"}), 400
        
        input_file = os.path.join(TEMP_DIR, secure_filename(file.filename))
        file.save(input_file)
        
        output_format = request.args.get("format", "docx")
        processed_filepath = process_report_card(input_file, output_format)
        
        if not processed_filepath:
            return jsonify({"error": "Failed to process file"}), 500
        
        return send_file(
            processed_filepath,
            as_attachment=True,
            download_name=f"converted.{output_format}",
            mimetype="application/pdf" if output_format == "pdf" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        logger.error(f"Error in upload_file: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5001)
