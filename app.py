"""Mondiale FCL rate automation: upload the quotation PDF, download the workbook."""

from __future__ import annotations

import logging
import os

from flask import Flask, jsonify, render_template, request, send_file

from excel_writer import build_workbook, output_filename
from mondiale_parser import parse_pdf

MAX_UPLOAD_BYTES = 40 * 1024 * 1024

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES

logging.basicConfig(level=logging.INFO)


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/api/process")
def process():
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify(error="No file was uploaded."), 400
    if not upload.filename.lower().endswith(".pdf"):
        return jsonify(error="Please upload the Mondiale FCL quotation as a `.pdf`."), 400

    try:
        regions = parse_pdf(upload.read())
    except Exception:
        app.logger.exception("Failed to parse %s", upload.filename)
        return jsonify(error="That PDF could not be read. Is it a Mondiale FCL quotation?"), 400

    rate_count = sum(len(region.rows) for region in regions)
    if not rate_count:
        return jsonify(
            error="No FCL rate rows were found. Check that this is the Import Seafreight FCL Quotation."
        ), 422

    app.logger.info("Parsed %s rate rows from %s", rate_count, upload.filename)
    workbook = build_workbook(regions)
    return send_file(
        workbook,
        as_attachment=True,
        download_name=output_filename(),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.errorhandler(413)
def too_large(_error):
    return jsonify(error="That file is larger than the 40 MB limit."), 413


if __name__ == "__main__":
    # Local development only; Render runs this through gunicorn (see Procfile).
    # Debug is opt-in because Werkzeug's debugger is a remote shell if exposed.
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 5000)),
        debug=os.environ.get("FLASK_DEBUG") == "1",
    )
