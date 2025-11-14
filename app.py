from flask import Flask, request, send_file, render_template
from docx import Document
import re
from collections import defaultdict
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FILE = "FINAL_DYNAMIC_TABLES.docx"

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

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

def process_docx(path):
    src = Document(path)

    data = defaultdict(lambda: defaultdict(dict))
    found_campuses = set()

    paragraphs = list(src.paragraphs)
    tables = src.tables
    table_index = 0

    current_campus = None
    current_question = None

    for p in paragraphs:
        text = p.text.strip()
        if re.search(r"IG-I.*Boys", text, re.IGNORECASE):
            parts = text.split("IG-I")
            if len(parts) > 1:
                campus_raw = parts[1].strip()
                campus = campus_raw.replace("–", "-").split("-")[0].strip()
                current_campus = campus
                found_campuses.add(campus)

        if re.match(r"Q#\d+", text):
            current_question = text
            if table_index < len(tables):
                tb = tables[table_index]
                table_index += 1

                for row in tb.rows[1:]:
                    teacher = row.cells[0].text.strip()
                    percent = row.cells[1].text.strip().replace("%", "")
                    if percent:
                        percent = f"{percent}%"
                    if teacher:
                        data[teacher][current_question][current_campus] = percent

    out = Document()
    campuses = sorted(list(found_campuses))

    for teacher, questions in data.items():
        h = out.add_heading(level=2)
        run = h.add_run(f"Teacher: {teacher}")
        run.bold = True

        table = out.add_table(rows=1, cols=len(campuses) + 1)
        hdr = table.rows[0].cells

        hdr[0].text = "Question"
        for i, c in enumerate(campuses):
            hdr[i + 1].text = c

        for cell in hdr:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.bold = True

        sorted_q = sorted(
            questions.keys(),
            key=lambda x: int(re.findall(r"\d+", x)[0])
        )

        for q in sorted_q:
            row = table.add_row().cells
            row[0].text = q
            for i, c in enumerate(campuses):
                row[i + 1].text = questions[q].get(c, "")

        set_borders(table)
        out.add_page_break()

    out.save(OUTPUT_FILE)
    return OUTPUT_FILE

@app.route("/")
def index():
    return render_template("upload.html")

@app.route("/upload", methods=["POST"])
def upload():
    file = request.files['file']
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    output_path = process_docx(file_path)
    return send_file(output_path, as_attachment=True)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)

