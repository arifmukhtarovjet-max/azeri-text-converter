import io
import os

from docx import Document
from flask import Flask, render_template, request, send_file

app = Flask("main")

# Your mapping (Cyrillic → Latin)

source = "АБЪЧДЕЯФЭЫЕХИЖКГЛМНОЮПРСШТУЦВЙЗабҹчдеяфэыехижкглмноюпрсштуцвйз"
target = "ABCÇDEƏFGĞHXIİJKQLMNOÖPRSŞTUÜVYZabcçdeəfgğhxıijkqlmnoöprsştuüvyz"

mapping = {source[i]: target[i] for i in range(len(source))}


def convert_text(text):
    return "".join(mapping.get(c, c) for c in text)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    file = request.files["file"]
    doc = Document(file)

    # Convert paragraphs
    for para in doc.paragraphs:
        para.text = convert_text(para.text)

    # Convert tables (important!)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell.text = convert_text(cell.text)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name="converted.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
