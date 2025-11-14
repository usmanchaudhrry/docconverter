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
# Normalize Teacher Name (case-insensitive + clean)
# -------------------------------------------------------
def normalize_teacher_name(name):
    # Remove long dashes
    name = name.replace("–", "-")

    # Remove repeated spaces
    name = re.sub(r"\s+", " ", name)

    # Strip
    name = name.strip()

    # Key for grouping (lowercase)
    key = name.lower()

    return key, name


# -------------------------------------------------------
# Add borders to table
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
# Add Header: Logo + Blue Box (ONLY 4 LINES)
# -------------------------------------------------------
def add_survey_header(doc, header_text):
    # Logo
    try:
        p = doc.add_paragraph()
        p.alignment = 1
        run = p.add_run()
        run.add_picture("static/logo.jpg", width=Inches(2.8))
    except:
        print("Logo missing: static/logo.jpg")

    # Blue Box Header
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]

    shade = OxmlElement("w:shd")
    shade.set(qn("w:fill"), "B7D3F2")
    cell._tc.get_or_add_tcPr().append(shade)

    set_borders(table)

    p = cell.paragraphs[0]
    p.alignment = 1

    for line in header_text.split("\n"):
        r = p.add_run(line)
        r.bold = True
        p.add_run("\n")

    doc.add_paragraph()


# -------------------------------------------------------
# UNIVERSAL TABLE HANDLER
# Handles: Name/Percentage, Name/Responses/Percentage,
# merged columns, ranking tables.
# -------------------------------------------------------
def extract_table(tb, qnum, campus, teacher_dict):

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

    # Determine mode
    is_ranking = ranking_col is not None
    value_col = ranking_col if is_ranking else percent_col

    if name_col is None or value_col is None:
        return False  # skip table

    # Extract rows
    for row in tb.rows[1:]:
        raw_name = row.cells[name_col].text.strip()
        if not raw_name:
            continue

        if "none of the above" in raw_name.lower():
            continue

        # Normalize name
        name_key, clean_name = normalize_teacher_name(raw_name)

        raw_val = row.cells[value_col].text.strip().replace("%", "")

        value = raw_val if is_ranking else (raw_val + "%" if raw_val else "")

        if value:
            teacher_dict[name_key][qnum][campus] = value
            # Also store pretty name
            teacher_dict[name_key]["_pretty_name"] = clean_name

    return True


# -------------------------------------------------------
# Detect Campus (IG-I / II / III)
# -------------------------------------------------------
def detect_campus(text):
    """
    Detects campus name from headings like:
        IG-I Mars - Boys
        IG-II Earth Computer Science - Girls
        IG-III Jupiter Science Wing
        IG-I South Block
        IG-III Mars Campus - Co-ed
        IG-II Venus
    
    Returns cleaned campus string or None.
    """

    # Pattern for:
    #   IG-I / IG-II / IG-III
    #   followed by multi-word campus
    #   optional hyphen
    #   optional gender (Boys/Girls/Co-ed)
    patterns = [
        r"(IG-I+)\s+([A-Za-z ]+?)(?:\s*-\s*(Boys|Girls|Co[- ]?ed))?$",
        r"(IG-II+)\s+([A-Za-z ]+?)(?:\s*-\s*(Boys|Girls|Co[- ]?ed))?$",
        r"(IG-III+)\s+([A-Za-z ]+?)(?:\s*-\s*(Boys|Girls|Co[- ]?ed))?$",
    ]

    t = text.strip()

    for pat in patterns:
        m = re.search(pat, t, re.IGNORECASE)
        if m:
            raw = m.group(2)
            if raw:
                # Remove trailing common words if present
                raw = raw.replace("Campus", "").strip()
                raw = re.sub(r"\s+", " ", raw)

                # Ensure capitalization consistency
                raw = " ".join(w.capitalize() for w in raw.split(" "))

                return raw

    return None




# -------------------------------------------------------
# PROCESS DOCX
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

    # ---------- Extract ONLY 4-Line Header ----------
    header_lines = []
    valid_keywords = [
        "learners",
        "academic year",
        "igcse boys",
        "college campus",
        "igcse-i",
        "igcse ii",
        "igcse iii",
    ]

    for p in paragraphs:
        text = p.text.strip()
        low = text.lower()

        if any(k in low for k in valid_keywords):
            header_lines.append(text)

        if low.startswith("q#1"):
            break

    survey_header = "\n".join(header_lines)

    # ---------- Extract Q# + Tables ----------
    for p in paragraphs:
        t = p.text.strip()

        # Campus
        camp = detect_campus(t)
        if camp:
            current_campus = camp
            found_campuses.add(camp)

        # Question
        m = re.match(r"(Q#\d+)", t)
        if m:
            qnum = m.group(1)
            question_text[qnum] = t

            # Next table belongs to this Q
            while table_index < len(tables):
                tb = tables[table_index]
                table_index += 1
                if extract_table(tb, qnum, current_campus, teacher_data):
                    break

    # ---------- Build output ----------
    out = Document()
    sorted_campuses = sorted(found_campuses)

    for teacher_key, qs in teacher_data.items():

        teacher_pretty = qs.get("_pretty_name", teacher_key.title())

        add_survey_header(out, survey_header)
        out.add_heading(f"Teacher: {teacher_pretty}", level=2)

        # Which campuses have data?
        active_campuses = [
            c for c in sorted_campuses if any(qs[q].get(c) for q in qs if not q.startswith("_"))
        ]

        # Create table
        table = out.add_table(rows=1, cols=1 + len(active_campuses))
        set_borders(table)

        hdr = table.rows[0].cells
        hdr[0].text = "Question"
        hdr[0].paragraphs[0].runs[0].bold = True

        for i, c in enumerate(active_campuses):
            hdr[i + 1].text = c
            hdr[i + 1].paragraphs[0].runs[0].bold = True

        # Sort Q1–Q8
        sorted_qs = sorted(
            [q for q in qs.keys() if not q.startswith("_")],
            key=lambda x: int(re.findall(r"\d+", x)[0])
        )

        for q in sorted_qs:
            row = table.add_row().cells
            row[0].text = question_text[q]
            for i, c in enumerate(active_campuses):
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
    path = os.path.join(UPLOAD_FOLDER, f.filename)
    f.save(path)
    output = process_docx(path)
    return send_file(output, as_attachment=True)


# -------------------------------------------------------
# RUN
# -------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
