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

# Configure detailed logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger(__name__)

app = Flask(__name__, static_url_path="/api")
CORS(app, resources={r"/*": {"origins": "*"}}, supports_credentials=True)

TEMP_DIR = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def clean_course_title(title):
    if not title or 'Study Hall' in title:
        return ''
    
    has_plus = title.startswith('+')
    title = title[1:].strip() if has_plus else title.strip()
    
    patterns = [
        r'\bG\d+(-\d+)?(?=\s|$|\()',  # Removes "G10", "G10-2"
        r'\b(?:Grade\s*)?\d{1,2}(?:th)?\s*(?:Grade\s*)?(?:-\d+)?(?=\s|$)',  # "Grade 10", "10"
        r'^(?:Senior|Junior)?\s*Electives\s*\d*-?',  # "Senior Electives-", "Electives 1 (G11)-"
        r'\s*Group\s*\d+\s*-',
        r'-\d+(\s|$)',  # Removes semester indicators "-1", "-2"
        r'\s*\([^)]*\)',  # Removes any text inside parentheses
        r'\s*-\s*(?=\S)',
        r'\s+-\s*$',
    ]
    
    for pattern in patterns:
        title = re.sub(pattern, '', title, flags=re.IGNORECASE)
    
    title = re.sub(r'\s+', ' ', title).strip('- ')
    
    logger.debug(f"Course title cleaned: {title}")
    return f"+ {title}" if has_plus else title

def process_table(table):
    seen_courses = set()
    rows_to_remove = []
    
    for row in table.rows[1:]:
        cells = row.cells
        if len(cells) >= 3:
            original_title = cells[0].text.strip()
            cleaned_title = clean_course_title(original_title)
            grade = cells[1].text.strip()
            gpa = cells[2].text.strip()
            
            if not cleaned_title:
                rows_to_remove.append(row)
                continue
            
            course_key = (cleaned_title.lower(), grade, gpa)  # Normalize casing
            if course_key in seen_courses:
                rows_to_remove.append(row)
            else:
                seen_courses.add(course_key)
                cells[0].text = cleaned_title
    
    for row in rows_to_remove:
        tbl = row._element.getparent()
        tbl.remove(row._element)
        logger.debug("Row removed from table")

def process_report_card(filepath):
    try:
        doc = Document(filepath)
        logger.debug("Document loaded successfully")
        
        for table in doc.tables:
            logger.debug("Processing table...")
            process_table(table)
        
        output_path = os.path.join(TEMP_DIR, f"processed_{os.path.basename(filepath)}")
        doc.save(output_path)
        
        if not os.path.exists(output_path):
            logger.error("Failed to save processed document")
            return None
        
        return output_path
    except Exception as e:
        logger.error(f"Error processing document: {str(e)}")
        return None

@app.route("/")
def home():
    return "Flask backend is running!", 200

@app.route("/api/upload", methods=["POST", "OPTIONS"])
def upload_file():
    if request.method == "OPTIONS":
        response = jsonify({"message": "CORS preflight passed"})
        response.headers.add("Access-Control-Allow-Origin", "*")
        response.headers.add("Access-Control-Allow-Methods", "POST, OPTIONS")
        response.headers.add("Access-Control-Allow-Headers", "Content-Type")
        return response, 200
    
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file uploaded"}), 400
        
        file = request.files["file"]
        if file.filename == "" or not allowed_file(file.filename):
            return jsonify({"error": "Invalid file type"}), 400
        
        input_file = os.path.join(TEMP_DIR, secure_filename(file.filename))
        file.save(input_file)
        logger.debug(f"File uploaded: {input_file}")
        
        processed_filepath = process_report_card(input_file)
        if not processed_filepath:
            return jsonify({"error": "File processing failed"}), 500
        
        return send_file(
            processed_filepath,
            as_attachment=True,
            download_name=f"converted_{secure_filename(file.filename)}",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    
    except Exception as e:
        logger.error(f"Error processing file: {str(e)}")
        return jsonify({"error": "Server encountered an error"}), 500

if __name__ == "__main__":
    from waitress import serve
    serve(app, host="0.0.0.0", port=5001)
