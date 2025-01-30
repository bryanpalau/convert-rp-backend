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

    # ✅ Modify paragraphs directly
    for para in doc.paragraphs:
        new_text = apply_conversion_rules(para.text)
        if new_text.strip():
            para.clear()
            para.add_run(new_text)

    # ✅ Process tables for courses, duplicate removal, and GPA adjustments
    for table in doc.tables:
        process_table(table)

    # ✅ Save modified document
    output_path = os.path.join(OUTPUT_FOLDER, "processed_" + os.path.basename(filepath))
    doc.save(output_path)
    return output_path

def apply_conversion_rules(text):
    """Applies all conversion rules to a given text."""

    # ✅ Remove Grade Level Prefixes
    text = re.sub(r'\b(G\d{1,2}[-\d]*|Grade \d{1,2}|[1-9]0)\b', '', text).strip()

    # ✅ Remove Course Group Labels (Electives, Senior Electives, Career Planning)
    text = re.sub(r'\b(Senior Electives-|Electives \d+ \(G\d+\)-|Career Planning \d+-\d+|Junior Electives-)\b', '', text).strip()

    # ✅ Simplify Course Titles
    text = re.sub(r'\b(G\d{1,2}-|G\d{1,2} )\b', '', text).strip()

    # ✅ Remove "Study Hall" courses
    if "Study Hall" in text:
        return ""  # Remove entire course

    return text

def process_table(table):
    """Process tables to remove duplicate courses and check GPA consistency."""
    seen_courses = {}

    for row in table.rows[1:]:  # Skip headers
        cells = [cell.text.strip() for cell in row.cells]

        if len(cells) < 3:
            continue  # Ignore empty rows

        course_title, grade, gpa = cells[0], cells[1], cells[2]

        # ✅ Apply conversion rules to course title
        clean_course_title = apply_conversion_rules(course_title)

        # ✅ Check for duplicate courses (same Grade & GPA)
        course_key = (clean_course_title, grade, gpa)
        if course_key in seen_courses:
            row._element.getparent().remove(row._element)  # Remove duplicate
        else:
            seen_courses[course_key] = True

        # ✅ Update row content
        row.cells[0].text = clean_course_title

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
