from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import re
import logging
import tempfile
from typing import Dict, List, Tuple
from werkzeug.utils import secure_filename
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Use temporary directory for file operations
TEMP_DIR = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'docx'}

# Configure CORS
CORS(app, resources={
    r"/*": {
        "origins": "*",
        "methods": ["OPTIONS", "POST"],
        "allow_headers": ["Content-Type"]
    }
})

def allowed_file(filename: str) -> bool:
    """Check if the file extension is allowed."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def clean_course_title(title: str) -> str:
    """Enhanced course title cleaning with more precise rules."""
    if not title or 'Study Hall' in title:
        return ''
        
    title = title.strip()
    
    # Keep the '+' prefix for AP/Honors courses
    has_plus = title.startswith('+')
    if has_plus:
        title = title[1:].strip()
    
    # Remove grade level and group prefixes more aggressively
    patterns = [
        r'^Math \d+[A-Z]?(?:-\d+)?-',  # Math prefixes
        r'^Science \d+[A-Z]?(?:-\d+)?-?',  # Science prefixes
        r'(?:G|Grade )\d+(?:-\d+)?-',  # Grade indicators
        r'^\d{1,2}(?:th)?\s*Grade\s*-?',  # Grade numbers
        r'(?:Junior|Senior)\s+Electives-',  # Elective prefixes
        r'Electives \d+ \([^)]+\)-',  # Elective group labels
        r'Career Planning \d+(?:-\d+)?',  # Career Planning prefixes
        r'Foreign Language-',  # Foreign Language prefix
        r'Individual Society Environment(?:\s*G\d+(?:-\d+)?)?-?',  # ISE prefix
        r'Military Training(?:\s*G\d+(?:-\d+)?)?-?',  # Military Training prefix
        r'Visual Performing Arts(?:\s*G\d+(?:-\d+)?)?-?',  # VPA prefix
        r'Group \d+-',  # Group numbers
        r'\d+[A-Z](?:-\d+)?-?'  # Grade section indicators (e.g., 7A-2)
    ]
    
    for pattern in patterns:
        title = re.sub(pattern, '', title, flags=re.IGNORECASE)
    
    # Clean up multiple spaces and hyphens
    title = re.sub(r'\s+', ' ', title)
    title = re.sub(r'-+', '-', title)
    title = title.strip(' -')
    
    # Add back the '+' prefix if it existed
    if has_plus:
        title = '+ ' + title
        
    return title

def process_table(table) -> None:
    """Process table with improved duplicate handling."""
    courses_by_semester = {}
    current_semester = None
    
    # First pass: collect and clean data
    for row in table.rows:
        cells = [cell.text.strip() for cell in row.cells]
        
        # Skip rows that don't have enough cells or are headers
        if len(cells) < 3 or any(header in cells[0].lower() for header in ['course title', 'average']):
            continue
            
        # Detect semester headers
        if any(sem in ' '.join(cells).upper() for sem in ['1ST SEMESTER', '2ND SEMESTER']):
            current_semester = '1st' if '1ST SEMESTER' in ' '.join(cells).upper() else '2nd'
            continue
            
        course_title, grade, gpa = cells[0], cells[1], cells[2]
        
        # Skip empty rows or non-numeric grades
        if not course_title or not grade.replace('.', '').isdigit():
            continue
            
        # Apply conversion rules
        clean_title = clean_course_title(course_title)
        if not clean_title:  # Skip if course was removed (e.g., Study Hall)
            continue
            
        # Store processed course with original grade and GPA
        if current_semester not in courses_by_semester:
            courses_by_semester[current_semester] = []
            
        courses_by_semester[current_semester].append({
            'title': clean_title,
            'grade': grade,
            'gpa': gpa
        })
    
    # Process courses, keeping those with different grades or GPAs
    for semester in courses_by_semester:
        unique_courses = []
        seen_exact_matches = set()
        
        for course in courses_by_semester[semester]:
            # Create key with all three values to check for exact duplicates
            course_key = (course['title'], course['grade'], course['gpa'])
            
            if course_key not in seen_exact_matches:
                unique_courses.append(course)
                seen_exact_matches.add(course_key)
        
        courses_by_semester[semester] = unique_courses
    
    # Clear existing table content while preserving formatting
    while len(table.rows) > 1:  # Keep header
        table._element.remove(table.rows[-1]._element)
    
    # Rebuild table with processed courses
    for semester, courses in courses_by_semester.items():
        # Add semester header
        semester_row = table.add_row()
        semester_cell = semester_row.cells[0]
        semester_cell.text = f"{semester} Semester"
        semester_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Add processed courses
        for course in courses:
            new_row = table.add_row()
            new_row.cells[0].text = course['title']
            new_row.cells[1].text = course['grade']
            new_row.cells[2].text = course['gpa']
            
            # Center align grade and GPA cells
            for cell in new_row.cells[1:]:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

def process_report_card(filepath: str) -> str:
    """Process the report card document while preserving formatting."""
    try:
        doc = Document(filepath)
        
        # Process each table in the document
        for table in doc.tables:
            process_table(table)
        
        # Save processed document to temporary file
        output_path = os.path.join(TEMP_DIR, f"processed_{os.path.basename(filepath)}")
        doc.save(output_path)
        return output_path
        
    except Exception as e:
        logger.error(f"Error processing report card: {str(e)}")
        raise

@app.route("/upload", methods=["POST"])
def upload_file():
    """Handle file upload and process the report card."""
    try:
        if "file" not in request.files:
            return jsonify({"error": "No file part"}), 400
        
        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "No selected file"}), 400
        
        if not allowed_file(file.filename):
            return jsonify({"error": "Invalid file type. Please upload a .docx file"}), 400
        
        filename = secure_filename(file.filename)
        filepath = os.path.join(TEMP_DIR, filename)
        
        try:
            file.save(filepath)
            processed_filepath = process_report_card(filepath)
            
            return send_file(
                processed_filepath,
                as_attachment=True,
                download_name=f"processed_{filename}",
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        finally:
            # Cleanup temporary files
            for path in [filepath, processed_filepath]:
                if os.path.exists(path):
                    try:
                        os.remove(path)
                    except Exception as e:
                        logger.error(f"Error removing temporary file {path}: {str(e)}")
                        
    except Exception as e:
        logger.error(f"Error in upload_file: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)