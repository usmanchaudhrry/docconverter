from flask import Flask, request, send_file, render_template
import os
import re
from collections import defaultdict
from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FILE = "FINAL_DYNAMIC_TABLES.docx"

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


# -------------------------------------------------------
# Add table borders
# -------------------------------------------------------
def set_borders(table):
    tbl = table._element
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        elem = OxmlElement(f"w:%s" % edge)
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:sz"), "10")
        elem.set(qn("w:color"), "000000")
        borders.append(elem)
    tbl.tblPr.append(borders)


# -------------------------------------------------------
# Header with logo + blue box
# -------------------------------------------------------
def add_survey_header(doc, header_text):

    # Add logo
    try:
        p_logo = doc.add_paragraph()
        p_logo.alignment = 1
        run = p_logo.add_run()
        run.add_picture("static/logo.jpg", width=Inches(3))
    except:
        print("Logo missing (static/logo.jpg).")

    # Blue header table
    table = doc.add_table(rows=1, cols=1)

    shade = OxmlElement("w:shd")
    shade.set(qn("w:fill"), "B7D3F2")

    tcPr = table.rows[0].cells[0]._tc.get_or_add_tcPr()
    tcPr.append(shade)

    set_borders(table)

    cell = table.rows[0].cells[0]
    p = cell.paragraphs[0]
    p.alignment = 1

    for line in header_text.split("\n"):
        r = p.add_run(line)
        r.bold = True
        p.add_run("\n")

    doc.add_paragraph()


# -------------------------------------------------------
# Extract teacher rows — IGNORE "Responses" column
# -------------------------------------------------------
def extract_teacher_data(tb, question, campus, data):

    col_count = len(tb.rows[0].cells)

    # Skip feedback tables
    header = [c.text.strip().lower() for c in tb.rows[0].cells]
    if header == ["id", "responses"]:
        return False

    # ---------------------- Q1–Q7 (3-column tables) ----------------------
    if col_count == 3:
        name_col = 0
        percentage_col = 2  # responses column ignored
        is_ranking = False

    # ---------------------- Q8 (Ranking) ----------------------
    elif col_count == 2:
        name_col = 0
        percentage_col = 1
        is_ranking = True

    else:
        return False

    # Extract rows
    for row in tb.rows[1:]:
        name = row.cells[name_col].text.strip()
        if not name or "none" in name.lower():
            continue

        raw = row.cells[percentage_col].text.strip().replace("%", "")

        # Ranking (no %)
        if question == "Q#8" or is_ranking:
            value = raw
        else:
            value = raw + "%" if raw else ""

        if value:
            data[name][question][campus] = value

    return True


# -------------------------------------------------------
# Detect campus
# -------------------------------------------------------
def detect_campus(text):
    m = re.search(r"IG-III\s+([A-Za-z]+)\s*-\s*Boys", text, re.IGNORECASE)
    if m:
        return m.group(1).strip()

    m = re.search(r"IG-II\s+([A-Za-z]+)\s*-\s*Boys", text, re.IGNORECASE)
    return m.group(1).strip() if m else None


# -------------------------------------------------------
# MAIN PROCESSOR
# -------------------------------------------------------
def process_docx(path):

    doc = Document(path)

    data = defaultdict(lambda: defaultdict(dict))
    question_text = {}
    found_campuses = set()
    tables = doc.tables
    paragraphs = doc.paragraphs
    table_index = 0
    current_campus = None

    # ---------------------------------------------------
    # Extract clean survey header
    # ---------------------------------------------------
    valid_header_keywords = [
        "learners",
        "academic year",
        "igcse boys",
        "college campus",
        "igcse-ii",
        "igcse iii",
        "igcse-iii"
    ]

    survey_header_lines = []

    for p in paragraphs:
        txt = p.text.strip()
        low = txt.lower()

        if any(k in low for k in valid_header_keywords):
            survey_header_lines.append(txt)
            continue

        if low.startswith("q#1"):
            break

    survey_header = "\n".join(survey_header_lines)

    # ---------------------------------------------------
    # Extract all questions + tables
    # ---------------------------------------------------
    for p in paragraphs:
        txt = p.text.strip()

        # Detect campus
        campus = detect_campus(txt)
        if campus:
            current_campus = campus
            found_campuses.add(campus)

        # Detect Q#
        m = re.match(r"(Q#\d+)\s*[:\-\.]?\s*(.*)", txt)
        if m:
            qnum = m.group(1)
            question_text[qnum] = txt

            # Read next table
            while table_index < len(tables):
                tb = tables[table_index]
                table_index += 1
                if extract_teacher_data(tb, qnum, current_campus, data):
                    break

    # ---------------------------------------------------
    # BUILD OUTPUT DOCX
    # ---------------------------------------------------
    out = Document()
    campuses = sorted(found_campuses)

    for teacher, questions in data.items():

        add_survey_header(out, survey_header)

        out.add_heading(f"Teacher: {teacher}", level=2)

        # Filter campuses with data
        active_camps = [
            c for c in campuses if any(questions[q].get(c, "") for q in questions)
        ]

        table = out.add_table(rows=1, cols=len(active_camps) + 1)
        hdr = table.rows[0].cells
        hdr[0].text = "Question"
        hdr[0].paragraphs[0].runs[0].bold = True

        for i, c in enumerate(active_camps):
            hdr[i + 1].text = c
            hdr[i + 1].paragraphs[0].runs[0].bold = True

        # Question order
        sorted_q = sorted(
            questions.keys(),
            key=lambda x: int(re.findall(r"\d+", x)[0])
        )

        for q in sorted_q:
            row = table.add_row().cells
            row[0].text = question_text.get(q, q)
            for i, c in enumerate(active_camps):
                row[i + 1].text = questions[q].get(c, "")

        set_borders(table)
        out.add_page_break()

    out.save(OUTPUT_FILE)
    return OUTPUT_FILE


# -------------------------------------------------------
# ROUTES
# -------------------------------------------------------
@app.route("/")
def index():
    return render_template("upload.html")


@app.route("/upload", methods=["POST"])
def upload():
    file = request.files["file"]
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)
    output = process_docx(file_path)
    return send_file(output, as_attachment=True)


# -------------------------------------------------------
# RUN APP
# -------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
