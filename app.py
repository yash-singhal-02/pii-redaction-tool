from flask import Flask, render_template, request, send_file
from pathlib import Path
from werkzeug.utils import secure_filename
import tempfile
import os
from pii_redactor import redact_docx

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024

@app.route("/", methods=["GET", "POST"])
def index():
    error = None
    if request.method == "POST":
        uploaded = request.files.get("file")
        if not uploaded or not uploaded.filename:
            error = "Please select a DOCX file."
            return render_template("index.html", error=error)
        if not uploaded.filename.lower().endswith(".docx"):
            error = "Only DOCX files are supported."
            return render_template("index.html", error=error)

        work = Path(tempfile.mkdtemp(prefix="pii_redaction_"))
        input_path = work / secure_filename(uploaded.filename)
        output_path = work / "redacted_output.docx"
        uploaded.save(input_path)

        try:
            redact_docx(input_path, output_path)
            return send_file(
                output_path,
                as_attachment=True,
                download_name="redacted_output.docx",
                mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as exc:
            error = f"Processing failed: {exc}"
            return render_template("index.html", error=error)

    return render_template("index.html", error=error)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
