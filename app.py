import os
import re
from collections import defaultdict
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from datetime import datetime
from flask import Flask, request, send_file, render_template

app = Flask(__name__)

# Use /tmp for Vercel (serverless writable directory)
UPLOAD_FOLDER = "/tmp/uploads"
OUTPUT_FILE = "/tmp/FINAL_DYNAMIC_TABLES.docx"
PDF_OUTPUT = "/tmp/PDF_TO_DOCX_OUTPUT.docx"
GRADE_OUTPUT = "/tmp/GRADE_PROCESSED.docx"

# Create upload folder if it doesn't exist
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


# -------------------------------------------------------
# Add borders
# -------------------------------------------------------
def set_borders(table):
    tbl = table._element
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        elem = OxmlElement(f"w:{edge}")
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:sz"), "10")
        elem.set(qn("w:color"), "000000")
        borders.append(elem)
    tbl.tblPr.append(borders)


# -------------------------------------------------------
# New Header Format - Match Image Style with Campus Name
# -------------------------------------------------------
def add_new_header(doc, grade_info):
    """Add header with Learners' Survey on left, Academic Session on right, and Campus+Grade centered"""
    
    # Header table with 3 columns
    header_table = doc.add_table(rows=1, cols=3)
    header_table.autofit = False
    header_table.allow_autofit = False
    
    # Remove borders
    for row in header_table.rows:
        for cell in row.cells:
            cell._element.get_or_add_tcPr().append(OxmlElement('w:tcBorders'))
    
    # Left: Learners' Survey
    left_cell = header_table.rows[0].cells[0]
    left_para = left_cell.paragraphs[0]
    left_run = left_para.add_run("Learners' Survey")
    left_run.font.size = Pt(11)
    left_run.font.color.rgb = RGBColor(150, 150, 150)
    left_run.italic = True
    left_para.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    
    # Right: Academic Session
    right_cell = header_table.rows[0].cells[2]
    right_para = right_cell.paragraphs[0]
    right_run = right_para.add_run("Academic Session 2026-2027")
    right_run.font.size = Pt(11)
    right_run.font.color.rgb = RGBColor(150, 150, 150)
    right_run.italic = True
    right_para.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
    
    doc.add_paragraph()
    
    # Campus Name and Grade Title - Centered and Bold
    campus_name = grade_info.get("campus_name", "Primary Campus")
    grade_section = grade_info.get("grade_section", "Grade-3")
    
    title_para = doc.add_paragraph()
    title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Campus name on first line
    campus_run = title_para.add_run(campus_name + "\n")
    campus_run.font.size = Pt(20)
    campus_run.font.bold = True
    campus_run.font.color.rgb = RGBColor(31, 56, 100)  # Dark blue
    
    # Grade on second line
    grade_run = title_para.add_run(grade_section)
    grade_run.font.size = Pt(20)
    grade_run.font.bold = True
    grade_run.font.color.rgb = RGBColor(31, 56, 100)  # Dark blue
    
    # Add horizontal line
    doc.add_paragraph("_" * 100)


# -------------------------------------------------------
# UNIVERSAL TABLE HANDLER (for IG campuses) - Extract subject info
# -------------------------------------------------------
def extract_table(tb, qnum, campus, data_dict):
    # Auto-assign default campus if missing
    if not campus:
        campus = "Percentage"

    header = [c.text.strip().lower() for c in tb.rows[0].cells]

    name_col = None
    subject_col = None
    percent_col = None
    ranking_col = None
    responses_col = None

    for idx, h in enumerate(header):
        if "name" in h or "teacher" in h:
            name_col = idx
        if "subject" in h:
            subject_col = idx
        if "percentage" in h or "%" in h:
            percent_col = idx
        if "ranking" in h:
            ranking_col = idx
        if "response" in h:
            responses_col = idx

    is_ranking = ranking_col is not None
    actual_col = ranking_col if is_ranking else percent_col

    if name_col is None or actual_col is None:
        return False

    for row in tb.rows[1:]:
        name = row.cells[name_col].text.strip()

        if not name:
            continue

        # Extract subject if available from Subject column
        subject = ""
        if subject_col is not None and subject_col < len(row.cells):
            subject = row.cells[subject_col].text.strip()

        # If no separate subject column, try to extract from teacher name
        # Format: "Mr. Adnan Zafar - Computer" or "Ms. Saira Hanif - Math"
        if not subject and " - " in name:
            parts = name.split(" - ", 1)  # Split only on first dash
            name = parts[0].strip()
            subject = parts[1].strip()

        # Normalize names
        name_normalized = name.strip().lower().title()

        # Extract responses count if available
        responses = ""
        if responses_col is not None and responses_col < len(row.cells):
            responses = row.cells[responses_col].text.strip()

        raw = row.cells[actual_col].text.strip().replace("%", "")
        value = raw if is_ranking else (raw + "%" if raw else "")

        if value:
            data_dict[name_normalized][qnum][campus] = {
                "value": value,
                "subject": subject,
                "responses": responses
            }
            
            # Debug print
            print(f"Extracted - Teacher: {name_normalized}, Subject: {subject}, Responses: {responses}, Value: {value}")

    return True



# -------------------------------------------------------
# Detect Campus Name (IG format)
# -------------------------------------------------------
def detect_campus(text):
    clean = " ".join(text.split())
    dash = r"[-–—]"

    patterns = [
        rf"(IG-[I1]+)\s*{dash}\s*(.+)$",
        rf"(IG-[I1]+)\s+(.+)$",
        rf"Grade(?:\s*\d+)?\s*{dash}\s*(.+)$",
        rf"Grade(?:\s*\d+)?\s+(.+)$",
    ]

    for pat in patterns:
        m = re.search(pat, clean, re.IGNORECASE)
        if m:
            if m.lastindex == 2:
                campus = m.group(2).strip()
            else:
                campus = m.group(1).strip()
                campus = re.sub(r"IG-[I1]+\s*", "", campus, flags=re.IGNORECASE).strip()
            
            return campus

    return None


# -------------------------------------------------------
# Add Grade Section Info Table - Match Image Style
# -------------------------------------------------------
def add_grade_section_table(doc, grade_info):
    """Add a table displaying grade section information with proper styling"""
    if not any(grade_info.values()):
        return
    
    table = doc.add_table(rows=4, cols=2)
    set_borders(table)
    
    # Set column widths
    table.columns[0].width = Inches(2.0)
    table.columns[1].width = Inches(4.5)
    
    labels = ["Grade-Section", "Survey Date", "Total Students", "Total Responses"]
    keys = ["grade_section", "survey_date", "total_students", "total_responses"]
    
    for idx, (label, key) in enumerate(zip(labels, keys)):
        row = table.rows[idx]
        
        # Label cell - left side, bold
        label_cell = row.cells[0]
        label_para = label_cell.paragraphs[0]
        label_para.text = ""
        label_run = label_para.add_run(label)
        label_run.font.bold = True
        label_run.font.size = Pt(11)
        
        # Light gray background for label
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), 'E7E6E6')
        label_cell._element.get_or_add_tcPr().append(shading)
        
        # Value cell - right side
        value_cell = row.cells[1]
        value_para = value_cell.paragraphs[0]
        
        # Format survey date if needed
        value = grade_info.get(key, "")
        if key == "survey_date" and value:
            try:
                date_obj = datetime.strptime(value, "%Y-%m-%d")
                value = date_obj.strftime("%B %d, %Y")
            except:
                pass
        
        value_para.text = ""
        value_run = value_para.add_run(value)
        value_run.font.size = Pt(11)
        
        if key == "grade_section":
            value_run.font.bold = True
            value_run.font.color.rgb = RGBColor(31, 56, 100)  # Dark blue
    
    doc.add_paragraph()


# -------------------------------------------------------
# Create styled question table matching the image
# -------------------------------------------------------
def create_question_table(doc, question_text, teacher_data, total_responses):
    """Create a table with Teacher, Subject, Responses, % columns with proper styling"""
    
    # Question heading with underline
    q_para = doc.add_paragraph()
    q_run = q_para.add_run(question_text)
    q_run.font.size = Pt(11)
    q_run.font.bold = True
    q_run.font.color.rgb = RGBColor(31, 56, 100)  # Dark blue
    
    # Add underline
    doc.add_paragraph("_" * 100)
    
    # Create table with 4 columns: Teacher, Subject, Responses, %
    table = doc.add_table(rows=1, cols=4)
    set_borders(table)
    
    # Set column widths
    table.columns[0].width = Inches(2.5)
    table.columns[1].width = Inches(2.0)
    table.columns[2].width = Inches(1.2)
    table.columns[3].width = Inches(0.8)
    
    # Header row with dark blue background
    hdr_cells = table.rows[0].cells
    headers = ["Teacher", "Subject", "Responses", "%"]
    
    for idx, header_text in enumerate(headers):
        cell = hdr_cells[idx]
        
        # Dark blue background
        shading = OxmlElement('w:shd')
        shading.set(qn('w:fill'), '1F3864')  # Dark blue
        cell._element.get_or_add_tcPr().append(shading)
        
        # White bold text
        para = cell.paragraphs[0]
        para.text = ""
        run = para.add_run(header_text)
        run.font.bold = True
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(255, 255, 255)  # White
        para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Data rows
    for teacher, data in sorted(teacher_data.items()):
        row = table.add_row()
        
        # Teacher name
        teacher_cell = row.cells[0]
        teacher_para = teacher_cell.paragraphs[0]
        teacher_para.text = ""
        teacher_run = teacher_para.add_run(teacher)
        teacher_run.font.size = Pt(10)
        
        # Subject
        subject_cell = row.cells[1]
        subject_para = subject_cell.paragraphs[0]
        subject_para.text = ""
        subject_run = subject_para.add_run(data.get("subject", ""))
        subject_run.font.size = Pt(10)
        
        # Responses count
        responses_cell = row.cells[2]
        responses_para = responses_cell.paragraphs[0]
        responses_para.text = ""
        responses_run = responses_para.add_run(str(data.get("responses", "")))
        responses_run.font.size = Pt(10)
        responses_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
        # Percentage in orange/gold color
        pct_cell = row.cells[3]
        pct_para = pct_cell.paragraphs[0]
        pct_para.text = ""
        
        # Calculate percentage if we have responses and total
        pct_value = data.get("value", "")
        if not pct_value and data.get("responses") and total_responses:
            try:
                resp_count = int(data.get("responses", 0))
                total_resp = int(total_responses)
                if total_resp > 0:
                    pct_calc = round((resp_count / total_resp) * 100)
                    pct_value = f"{pct_calc}%"
            except:
                pass
        
        pct_run = pct_para.add_run(pct_value)
        pct_run.font.size = Pt(10)
        pct_run.font.bold = True
        pct_run.font.color.rgb = RGBColor(191, 143, 0)  # Gold/orange
        pct_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    doc.add_paragraph()


# -------------------------------------------------------
# Process DOCX (IG campuses - updated for new format)
# -------------------------------------------------------
def process_docx(path, grade_info=None):
    doc = Document(path)
    paragraphs = doc.paragraphs
    tables = doc.tables

    teacher_data = defaultdict(lambda: defaultdict(dict))
    question_text = {}
    table_index = 0
    current_campus = None
    found_campuses = set()

    header_lines = []

    valid_keywords = [
        "learners",
        "academic year",
        "igcse boys",
        "igcse girls",
        "college campus",
        "igcse-i",
        "igcse ii",
        "igcse iii"
    ]

    # Extract header
    for p in paragraphs:
        t = p.text.strip()
        low = t.lower()

        if any(k in low for k in valid_keywords):
            header_lines.append(t)

        # Stop reading header once questions start
        if low.startswith("q#1") or low.startswith("q-1"):
            break

    survey_header = "\n".join(header_lines)

    # PARSE DOCUMENT
    for p in paragraphs:
        t = p.text.strip()

        # Detect campus (IG format)
        camp = detect_campus(t)
        if camp:
            current_campus = camp
            found_campuses.add(camp)

        # Detect questions
        m = re.match(r"(Q[#\-]\d+)", t, re.IGNORECASE)
        if m:
            qnum = m.group(1).upper().replace("-", "#")
            question_text[qnum] = t

            # ★ AUTO-ASSIGN CAMPUS USING Q#1 if none found ★
            if current_campus is None:
                current_campus = "Percentage"
                found_campuses.add("Percentage")

            # Extract next table
            while table_index < len(tables):
                tb = tables[table_index]
                table_index += 1
                if extract_table(tb, qnum, current_campus, teacher_data):
                    break

    # BUILD OUTPUT
    out = Document()
    
    # Add new header format
    add_new_header(out, grade_info if grade_info else {})
    
    # Add grade section info table if provided
    if grade_info:
        add_grade_section_table(out, grade_info)
    
    # Get total responses from grade_info
    total_responses = grade_info.get("total_responses", "") if grade_info else ""
    
    # Group data by question instead of by teacher
    questions_data = defaultdict(lambda: {})
    
    for teacher, qs in teacher_data.items():
        for qnum, campuses in qs.items():
            for campus, data in campuses.items():
                if isinstance(data, dict):
                    questions_data[qnum][teacher] = data
                else:
                    # Handle old format
                    questions_data[qnum][teacher] = {
                        "value": data,
                        "subject": "",
                        "responses": ""
                    }
    
    # Create table for each question
    sorted_questions = sorted(questions_data.keys(), key=lambda x: int(re.findall(r"\d+", x)[0]))
    
    for qnum in sorted_questions:
        q_text = question_text.get(qnum, qnum)
        teacher_data_for_q = questions_data[qnum]
        create_question_table(out, q_text, teacher_data_for_q, total_responses)

    out.save(OUTPUT_FILE)
    return OUTPUT_FILE



# -------------------------------------------------------
# Process Grade-based DOCX (NEW FUNCTION)
# -------------------------------------------------------
# def process_grade_docx(path):
#     doc = Document(path)
#     paragraphs = doc.paragraphs
#     tables = doc.tables

#     teacher_data = defaultdict(lambda: defaultdict(dict))
#     question_text = {}
#     current_campus = None
#     found_campuses = []
#     table_index = 0

#     header_lines = []
#     valid_keywords = ["learners", "academic year", "pre-school", "preschool", "campus"]

#     # Extract header
#     for p in paragraphs:
#         t = p.text.strip()
#         low = t.lower()
        
#         if any(k in low for k in valid_keywords):
#             header_lines.append(t)
        
#         if low.startswith("q") or low.startswith("dated"):
#             break

#     survey_header = "\n".join(header_lines) if header_lines else "Learner's Survey\nAcademic Year 2025-2026"

#     # Parse document
#     for p in paragraphs:
#         t = p.text.strip()
#         low = t.lower()

#         # Detect campus (Grade 1 - Mars, Grade 1 - Venus, etc.)
#         grade_match = re.search(r"Grade\s+\d+\s*[-–—]\s*(.+)", t, re.IGNORECASE)
#         if grade_match:
#             current_campus = grade_match.group(1).strip()
#             found_campuses.append(current_campus)
#             continue

#         # Detect questions
#         q_match = re.match(r"(Q[-#]?\d+)[\:\.]?\s*(.+)", t, re.IGNORECASE)
#         if q_match and current_campus:
#             qnum = q_match.group(1).upper().replace("-", "#")
#             q_text = q_match.group(2).strip()
#             question_text[qnum] = f"{qnum}: {q_text}"

#             # Extract table data
#             if table_index < len(tables):
#                 tb = tables[table_index]
#                 table_index += 1

#                 # Parse table rows
#                 for row in tb.rows[1:]:  # Skip header
#                     try:
#                         cells = row.cells
#                         if len(cells) >= 3:
#                             teacher_name = cells[0].text.strip()
#                             percentage = cells[2].text.strip()
                            
#                             if teacher_name and percentage:
#                                 teacher_data[teacher_name][qnum][current_campus] = percentage
#                     except:
#                         continue

#     # Build output document
#     out = Document()
#     unique_campuses = []
#     for campus in found_campuses:
#         if campus not in unique_campuses:
#             unique_campuses.append(campus)

#     for teacher, qs in sorted(teacher_data.items()):
#         add_survey_header(out, survey_header)
#         out.add_heading(f"Teacher: {teacher}", level=2)

#         # Filter campuses where this teacher has data
#         active_camps = [c for c in unique_campuses if any(qs.get(q, {}).get(c) for q in qs)]

#         if not active_camps:
#             continue

#         # Create table
#         table = out.add_table(rows=1, cols=1 + len(active_camps))
#         set_borders(table)

#         # Headers
#         hdr = table.rows[0].cells
#         hdr[0].text = "Question"
#         hdr[0].paragraphs[0].runs[0].bold = True

#         for i, campus in enumerate(active_camps):
#             hdr[i + 1].text = campus
#             hdr[i + 1].paragraphs[0].runs[0].bold = True

#         # Questions
#         sorted_qs = sorted(qs.keys(), key=lambda x: int(re.findall(r"\d+", x)[0]) if re.findall(r"\d+", x) else 0)

#         for q in sorted_qs:
#             row = table.add_row().cells
#             row[0].text = question_text.get(q, q)
#             for i, campus in enumerate(active_camps):
#                 row[i + 1].text = qs[q].get(campus, "")

#         # ADD "Monthly Grading" ROW
#         monthly_row = table.add_row().cells
#         monthly_row[0].text = "Monthly Grading"
#         monthly_row[0].paragraphs[0].runs[0].bold = True
#         for i in range(len(active_camps)):
#             monthly_row[i + 1].text = ""

#         out.add_page_break()

#     out.save(GRADE_OUTPUT)
#     return GRADE_OUTPUT


# -------------------------------------------------------
# PDF → DOCX CONVERTER - Updated with new styling
# -------------------------------------------------------
def convert_pdf_to_docx(pdf_path, output_path="PDF_CONVERTED.docx", grade_info=None):
    import pdfplumber
    import re
    from collections import defaultdict
    from docx import Document
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    def add_borders(table):
        tbl = table._element
        borders = OxmlElement('w:tblBorders')
        for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            elem = OxmlElement(f'w:{edge}')
            elem.set(qn('w:val'), 'single')
            elem.set(qn('w:sz'), '10')
            elem.set(qn('w:color'), '000000')
            borders.append(elem)
        tbl.tblPr.append(borders)

    doc = Document()

    # Add new header format
    add_new_header(doc, grade_info if grade_info else {})
    
    # Add grade section info table if provided
    if grade_info:
        add_grade_section_table(doc, grade_info)
    
    # Get total responses from grade_info
    total_responses = grade_info.get("total_responses", "") if grade_info else ""

    q_pattern = re.compile(r".*?(Q[#\s]*\d+)\s*[:\.\-]*\s*(.*)", re.IGNORECASE)
    teacher_pattern = re.compile(r"(.+?)\s+(\d+)$")
    ranking_pattern = re.compile(r"^\s*(\d+)\s+(.*)$")

    questions = {}
    current_q = None

    pdf = pdfplumber.open(pdf_path)

    for page in pdf.pages:
        text = page.extract_text() or ""

        for line in text.split("\n"):
            clean = line.strip()
            if not clean:
                continue

            mq = q_pattern.match(clean)
            if mq:
                q_id = mq.group(1).replace(" ", "").upper()
                q_text = mq.group(1) + " " + mq.group(2)

                current_q = q_id
                questions.setdefault(q_id, {"text": q_text, "entries": []})
                continue

            if not current_q:
                continue

            if current_q == "Q#8":
                mr = ranking_pattern.match(clean)
                if mr:
                    rank = mr.group(1)
                    teacher = mr.group(2).strip()
                    questions[current_q]["entries"].append((teacher, rank))
                continue

            mt = teacher_pattern.search(clean)
            if mt:
                teacher = mt.group(1).strip()
                count = int(mt.group(2))
                questions[current_q]["entries"].append((teacher, count))

    pdf.close()

    any_data = False

    for q_id, block in questions.items():
        entries = block["entries"]
        if not entries:
            continue

        any_data = True
        
        # Question heading with underline
        q_para = doc.add_paragraph()
        q_run = q_para.add_run(block["text"])
        q_run.font.size = Pt(11)
        q_run.font.bold = True
        q_run.font.color.rgb = RGBColor(31, 56, 100)  # Dark blue
        
        # Add underline
        doc.add_paragraph("_" * 100)

        if q_id == "Q#8":
            # Ranking table - 2 columns
            table = doc.add_table(rows=1, cols=2)
            add_borders(table)
            
            # Set column widths
            table.columns[0].width = Inches(4.5)
            table.columns[1].width = Inches(1.5)

            # Header row with dark blue background
            hdr_cells = table.rows[0].cells
            headers = ["Teacher", "Ranking"]
            
            for idx, header_text in enumerate(headers):
                cell = hdr_cells[idx]
                
                # Dark blue background
                shading = OxmlElement('w:shd')
                shading.set(qn('w:fill'), '1F3864')
                cell._element.get_or_add_tcPr().append(shading)
                
                # White bold text
                para = cell.paragraphs[0]
                para.text = ""
                run = para.add_run(header_text)
                run.font.bold = True
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor(255, 255, 255)
                para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            for teacher_full, rank in entries:
                # Split teacher name and subject if combined
                teacher_name = teacher_full
                if " - " in teacher_full:
                    parts = teacher_full.split(" - ", 1)
                    teacher_name = parts[0].strip()
                
                row = table.add_row().cells
                
                # Teacher (without subject for ranking)
                row[0].paragraphs[0].text = ""
                t_run = row[0].paragraphs[0].add_run(teacher_name)
                t_run.font.size = Pt(10)
                
                # Ranking
                row[1].paragraphs[0].text = ""
                r_run = row[1].paragraphs[0].add_run(str(rank))
                r_run.font.size = Pt(10)
                row[1].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            doc.add_paragraph()
            continue

        # Regular questions - 4 columns: Teacher, Subject, Responses, %
        # First, split teacher names and subjects, then group
        teacher_subject_map = {}  # Store subject for each teacher
        grouped = defaultdict(int)
        total = 0
        
        for teacher_full, count in entries:
            # Split teacher name and subject if combined
            teacher_name = teacher_full
            subject = ""
            
            if " - " in teacher_full:
                parts = teacher_full.split(" - ", 1)
                teacher_name = parts[0].strip()
                subject = parts[1].strip()
            
            teacher_subject_map[teacher_name] = subject
            grouped[teacher_name] += count
            total += count

        table = doc.add_table(rows=1, cols=4)
        add_borders(table)
        
        # Set column widths
        table.columns[0].width = Inches(2.5)
        table.columns[1].width = Inches(2.0)
        table.columns[2].width = Inches(1.2)
        table.columns[3].width = Inches(0.8)

        # Header row with dark blue background
        hdr_cells = table.rows[0].cells
        headers = ["Teacher", "Subject", "Responses", "%"]
        
        for idx, header_text in enumerate(headers):
            cell = hdr_cells[idx]
            
            # Dark blue background
            shading = OxmlElement('w:shd')
            shading.set(qn('w:fill'), '1F3864')
            cell._element.get_or_add_tcPr().append(shading)
            
            # White bold text
            para = cell.paragraphs[0]
            para.text = ""
            run = para.add_run(header_text)
            run.font.bold = True
            run.font.size = Pt(11)
            run.font.color.rgb = RGBColor(255, 255, 255)
            para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        sorted_teachers = sorted(
            grouped.items(),
            key=lambda x: (x[0].lower().startswith("none of the above"), x[0].lower())
        )

        for teacher, count in sorted_teachers:
            row = table.add_row()
            
            # Teacher name (without subject)
            row.cells[0].paragraphs[0].text = ""
            t_run = row.cells[0].paragraphs[0].add_run(teacher)
            t_run.font.size = Pt(10)
            
            # Subject (extracted from teacher name)
            row.cells[1].paragraphs[0].text = ""
            subject = teacher_subject_map.get(teacher, "")
            s_run = row.cells[1].paragraphs[0].add_run(subject if subject else "-")
            s_run.font.size = Pt(10)
            
            # Responses count
            row.cells[2].paragraphs[0].text = ""
            resp_run = row.cells[2].paragraphs[0].add_run(str(count))
            resp_run.font.size = Pt(10)
            row.cells[2].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            
            # Percentage in gold/orange - calculate from total_responses provided by user
            pct = 0
            if total_responses:
                try:
                    total_resp = int(total_responses)
                    if total_resp > 0:
                        pct = round((count / total_resp) * 100)
                except:
                    pct = 0
            
            row.cells[3].paragraphs[0].text = ""
            pct_run = row.cells[3].paragraphs[0].add_run(f"{pct}%")
            pct_run.font.size = Pt(10)
            pct_run.font.bold = True
            pct_run.font.color.rgb = RGBColor(191, 143, 0)  # Gold/orange
            row.cells[3].paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        doc.add_paragraph()

    if not any_data:
        raise Exception("Nothing detected in PDF.")

    doc.save(output_path)
    return output_path


# -------------------------------------------------------
# ROUTES
# -------------------------------------------------------
@app.route("/")
def index():
    return render_template("upload.html")


@app.route("/upload", methods=["POST"])
def upload():
    f = request.files["file"]
    file_path = os.path.join(UPLOAD_FOLDER, f.filename)
    f.save(file_path)
    
    # Get grade section information
    campus_name = request.form.get("campus_name", "").strip()
    grade_section = request.form.get("grade_section", "").strip()
    survey_date = request.form.get("survey_date", "").strip()
    total_students = request.form.get("total_students", "").strip()
    total_responses = request.form.get("total_responses", "").strip()
    
    grade_info = {
        "campus_name": campus_name,
        "grade_section": grade_section,
        "survey_date": survey_date,
        "total_students": total_students,
        "total_responses": total_responses
    }
    
    output = process_docx(file_path, grade_info)
    return send_file(output, as_attachment=True)


@app.route("/upload_grade", methods=["POST"])
def upload_grade():
    f = request.files["file"]
    file_path = os.path.join(UPLOAD_FOLDER, f.filename)
    f.save(file_path)
    output = process_grade_docx(file_path)
    return send_file(output, as_attachment=True)


@app.route("/convert_pdf", methods=["POST"])
def convert_pdf():
    f = request.files.get("pdf_file")
    
    # Get grade section information
    campus_name = request.form.get("campus_name", "").strip()
    grade_section = request.form.get("grade_section", "").strip()
    survey_date = request.form.get("survey_date", "").strip()
    total_students = request.form.get("total_students", "").strip()
    total_responses = request.form.get("total_responses", "").strip()
    
    grade_info = {
        "campus_name": campus_name,
        "grade_section": grade_section,
        "survey_date": survey_date,
        "total_students": total_students,
        "total_responses": total_responses
    }

    if not f:
        return "No file selected", 400

    pdf_path = os.path.join(UPLOAD_FOLDER, f.filename)
    f.save(pdf_path)

    try:
        output = convert_pdf_to_docx(pdf_path, grade_info=grade_info)
        return send_file(output, as_attachment=True)
    except Exception as e:
        return f"Error: {str(e)}", 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)