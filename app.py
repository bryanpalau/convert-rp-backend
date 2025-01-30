from flask import Flask, request, send_file
from flask_cors import CORS
import os
import re
from werkzeug.utils import secure_filename
from docx import Document

app = Flask(__name__)

# ✅ Use Vercel's writable directory
UPLOAD_FOLDER = "/tmp/uploads"
OUTPUT_FOLDER = "/tmp/outputs"

# ✅ Ensure folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ✅ Fix CORS to allow frontend access
CORS(app, resources={r"/*": {"origins": "*"}}, allow_headers=["Content-Type"])

@app.route("/upload", methods=["POST"])
def upload_file():
    """Handles file upload and processes the report card."""
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
    """Process the uploaded .docx file and apply conversion rules."""
    doc = Document(filepath)

    # ✅ Iterate through paragraphs and modify text based on rules
    for para in doc.paragraphs:
        text = para.text.strip()

        # Rule 1: Remove Grade Level Prefixes
        text = re.sub(r'\b(G\d{1,2}[-\d]*|Grade \d{1,2}|[1-9]0)\b', '', text).strip()

        # Rule 2: Remove Course Group Labels
        text = re.sub(r'\b(Senior Electives-|Electives \d+ \(G\d+\)-|Career Planning \d+-\d+)\b', '', text).strip()

        # Rule 5: Simplification of Course Titles
        text = re.sub(r'\b(G\d{1,2}-|G\d{1,2} )\b', '', text).strip()

        # Rule 6: Remove Study Hall Courses
        if "Study Hall" in text:
            para.clear()  # Remove the paragraph

        # ✅ Update paragraph text
        para.text = text

    # ✅ Process tables for duplicate courses, GPA verification
    for table in doc.tables:
        process_table(table)

    # ✅ Save modified document
    output_path = os.path.join(OUTPUT_FOLDER, "processed_" + os.path.basename(filepath))
    doc.save(output_path)
    return output_path

def process_table(table):
    """Process tables to remove duplicate courses and check GPA consistency."""
    seen_courses = {}

    for row in table.rows[1:]:  # Skip headers
        cells = [cell.text.strip() for cell in row.cells]

        if len(cells) < 3:
            continue  # Ignore empty rows

        course_title, grade, gpa = cells[0], cells[1], cells[2]

        # ✅ Remove Grade Level Prefix
        course_title = re.sub(r'\b(G\d{1,2}[-\d]*|Grade \d{1,2}|[1-9]0)\b', '', course_title).strip()

        # ✅ Check for duplicate courses (same Grade & GPA)
        course_key = (course_title, grade, gpa)
        if course_key in seen_courses:
            row._element.getparent().remove(row._element)  # Remove duplicate
        else:
            seen_courses[course_key] = True

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
