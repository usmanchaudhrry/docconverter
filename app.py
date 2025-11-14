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
        elem = OxmlElement(f"w:{edge}")
        elem.set(qn("w:val"), "single")
        elem.set(qn("w:sz"), "10")
        elem.set(qn("w:color"), "000000")
        borders.append(elem)
    tbl.tblPr.append(borders)


# -------------------------------------------------------
# Add page header (logo + blue box)
# -------------------------------------------------------
def add_survey_header(doc, header_text):

    # Logo
    try:
        p_logo = doc.add_paragraph()
        p_logo.alignment = 1
        r = p_logo.add_run()
        r.add_picture("static/logo.jpg", width=Inches(2.8))
    except:
        print("Logo missing (static/logo.jpg)")

    # Blue box
    t = doc.add_table(rows=1, cols=1)
    cell = t.rows[0].cells[0]

    shade = OxmlElement("w:shd")
    shade.set(qn("w:fill"), "B7D3F2")
    cell._tc.get_or_add_tcPr().append(shade)
    set_borders(t)

    p = cell.paragraphs[0]
    p.alignment = 1

    for line in header_text.split("\n"):
        run = p.add_run(line)
        run.bold = True
        p.add_run("\n")

    doc.add_paragraph()


# -------------------------------------------------------
# Universal Table Handler
# -------------------------------------------------------
def extract_table(tb, qnum, campus, data_dict):

    header = [c.text.strip().lower() for c in tb.rows[0].cells]

    name_col = None
    percent_col = None
    ranking_col = None

    for idx, h in enumerate(header):
        if "name" in h:
            name_col = idx
        if "percentage" in h:
            percent_col = idx
        if "ranking" in h:
            ranking_col = idx

    # Ranking mode
    is_ranking = ranking_col is not None
    actual_col = ranking_col if is_ranking else percent_col

    if name_col is None or actual_col is None:
        return False

    for row in tb.rows[1:]:
        name = row.cells[name_col].text.strip()
        if not name or "none of the above" in name.lower():
            continue

        raw = row.cells[actual_col].text.strip().replace("%", "")

        value = raw if is_ranking else (raw + "%" if raw else "")

        if value:
            data_dict[name][qnum][campus] = value

    return True


# -------------------------------------------------------
# Detect Campus Name
# -------------------------------------------------------
def detect_campus(text):
    clean = " ".join(text.split())  # normalize spaces

    patterns = [
        # IG level:
        r"(IG-[I1]+\s+[A-Za-z ]+?)\s*-\s*(Boys|Girls|Campus|Boys Campus|Girls Campus)",
        r"(IG-[I1]+\s+[A-Za-z ]+?)\s*$",

        # Grade level:
        r"(Grade\s*\d+\s+[A-Za-z ]+?)\s*-\s*(Boys|Girls|Campus|Boys Campus|Girls Campus)",
        r"(Grade\s*\d+\s+[A-Za-z ]+?)\s*$"
    ]

    for pat in patterns:
        m = re.search(pat, clean, re.IGNORECASE)
        if m:
            campus = m.group(1)

            # Remove IG level prefix
            campus = re.sub(r"IG-[I1]+\s*", "", campus, flags=re.IGNORECASE)

            # Remove Grade prefix
            campus = re.sub(r"Grade\s*\d+\s*", "", campus, flags=re.IGNORECASE)

            return campus.strip()

    return None



# -------------------------------------------------------
# Process DOCX
# -------------------------------------------------------
def process_docx(path):
    doc = Document(path)

    paragraphs = doc.paragraphs
    tables = doc.tables

    teacher_data = defaultdict(lambda: defaultdict(dict))
    question_text = {}

    table_index = 0
    current_campus = None
    found_campuses = set()

    # ---------------- Extract Clean 4-Line Header ----------------
    header_lines = []

    valid_keywords = [
        "learners",
        "academic year",
        "igcse boys",
        "college campus",
        "igcse-i",
        "igcse ii",
        "igcse iii"
    ]

    for p in paragraphs:
        t = p.text.strip()
        low = t.lower()

        if any(k in low for k in valid_keywords):
            header_lines.append(t)

        if low.startswith("q#1"):
            break

    survey_header = "\n".join(header_lines)

    # ---------------- Extract Questions + Tables ----------------
    for p in paragraphs:
        t = p.text.strip()

        # campus detect
        camp = detect_campus(t)
        if camp:
            current_campus = camp
            found_campuses.add(camp)

        # question detect
        m = re.match(r"(Q#\d+)", t)
        if m:
            qnum = m.group(1)
            question_text[qnum] = t

            while table_index < len(tables):
                tb = tables[table_index]
                table_index += 1
                if extract_table(tb, qnum, current_campus, teacher_data):
                    break

    # ---------------- Build Output DOCX ----------------
    out = Document()
    campuses_sorted = sorted(found_campuses)

    for teacher, qs in teacher_data.items():

        add_survey_header(out, survey_header)
        out.add_heading(f"Teacher: {teacher}", level=2)

        # Only show campuses where teacher has data
        active_camps = [c for c in campuses_sorted if any(qs[q].get(c) for q in qs)]

        table = out.add_table(rows=1, cols=1 + len(active_camps))
        set_borders(table)

        hdr = table.rows[0].cells
        hdr[0].text = "Question"
        hdr[0].paragraphs[0].runs[0].bold = True

        for i, c in enumerate(active_camps):
            hdr[i + 1].text = c
            hdr[i + 1].paragraphs[0].runs[0].bold = True

        sorted_qs = sorted(qs.keys(), key=lambda x: int(re.findall(r"\d+", x)[0]))

        for q in sorted_qs:
            row = table.add_row().cells
            row[0].text = question_text[q]
            for i, c in enumerate(active_camps):
                row[i + 1].text = qs[q].get(c, "")

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
    f = request.files["file"]
    file_path = os.path.join(UPLOAD_FOLDER, f.filename)
    f.save(file_path)
    output = process_docx(file_path)
    return send_file(output, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
