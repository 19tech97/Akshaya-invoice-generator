"""
Flask web application for invoice generation.
Provides a UI for uploading files, configuring parameters, and downloading invoices.
"""

import os
import uuid
import shutil
import tempfile
import zipfile
import json
from flask import Flask, render_template, request, jsonify, send_file, session
from werkzeug.utils import secure_filename

from utils.invoice_engine import generate_invoices
from utils.excel_helpers import load_excel_files

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "invoice-gen-secret-key-change-me")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB max upload

# Temp storage for sessions
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "invoice_gen_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


def get_session_dir():
    """Get or create a unique temp directory for this session."""
    if "session_id" not in session:
        session["session_id"] = str(uuid.uuid4())
    d = os.path.join(UPLOAD_DIR, session["session_id"])
    os.makedirs(d, exist_ok=True)
    return d


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload-conso", methods=["POST"])
def upload_conso():
    """Upload the consolidated data file and return its column names."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Empty filename"}), 400

    sdir = get_session_dir()
    fname = secure_filename(f.filename)
    path = os.path.join(sdir, "conso_" + fname)
    f.save(path)

    try:
        df = load_excel_files(path)
        columns = df.columns.tolist()
        session["conso_path"] = path
        return jsonify({"columns": columns, "rows": len(df), "filename": fname})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route("/upload-template", methods=["POST"])
def upload_template():
    """Upload the invoice template file and return its sheet names."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Empty filename"}), 400

    sdir = get_session_dir()
    fname = secure_filename(f.filename)
    path = os.path.join(sdir, "template_" + fname)
    f.save(path)

    try:
        from openpyxl import load_workbook
        wb = load_workbook(path, read_only=True)
        sheets = wb.sheetnames
        # Also get table names per sheet
        wb2 = load_workbook(path)
        tables = {}
        for sn in wb2.sheetnames:
            ws = wb2[sn]
            tables[sn] = list(ws.tables.keys()) if ws.tables else []
        wb.close()
        wb2.close()
        session["template_path"] = path
        return jsonify({"sheets": sheets, "tables": tables, "filename": fname})
    except Exception as e:
        return jsonify({"error": str(e)}), 400


@app.route("/generate", methods=["POST"])
def generate():
    """Run the invoice generation with the provided configuration."""
    try:
        data = request.get_json()
        sdir = get_session_dir()
        output_dir = os.path.join(sdir, "output")
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        os.makedirs(output_dir)

        # Parse PM info from the JSON (it comes as a list of rule objects)
        pm_info = {}
        for rule in data.get("pm_rules", []):
            if rule.get("type_name"):
                pm_info[rule["type_name"]] = {
                    "page_limit": float(rule.get("page_limit", 0)),
                    "ad_hoc_till_page_limit": float(rule.get("base_fee", 0)),
                    "after_page_limit_basis": rule.get("overage_basis", "rate"),
                    "rate_after_limit": float(rule.get("overage_rate", 0)),
                }

        config = {
            "conso_path": session.get("conso_path", ""),
            "template_path": session.get("template_path", ""),
            "editor_col": data["editor_col"],
            "level_1_header": data["level_1_header"],
            "level_2_header": data["level_2_header"],
            "sheet_name": data["sheet_name"],
            "inv_sheet_name": data["inv_sheet_name"],
            "required_columns": data["required_columns"],
            "starting_cell": data.get("starting_cell", "A3"),
            "table_name": data.get("table_name", "invoice"),
            "inv_no_cell": data.get("inv_no_cell", "B3"),
            "inv_dt_cell": data.get("inv_dt_cell", "B5"),
            "inv_dt": data.get("inv_dt", ""),
            "editor_name_cell": data.get("editor_name_cell", "A37"),
            "inv_total_amt_cell": data.get("inv_total_amt_cell", "A40"),
            "inv_col_name": data.get("inv_col_name", ""),
            "pm_info": pm_info,
            "currency_prefix": data.get("currency_prefix", "SGD"),
            "output_dir": output_dir,
        }

        if not config["conso_path"] or not os.path.exists(config["conso_path"]):
            return jsonify({"error": "Consolidated data file not found. Please re-upload."}), 400
        if not config["template_path"] or not os.path.exists(config["template_path"]):
            return jsonify({"error": "Template file not found. Please re-upload."}), 400

        generated = generate_invoices(config)

        files = [os.path.basename(p) for p in generated]
        session["output_dir"] = output_dir
        return jsonify({"success": True, "files": files, "count": len(files)})

    except Exception as e:
        import traceback
        return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500


@app.route("/download/<filename>")
def download_file(filename):
    """Download a single generated invoice."""
    output_dir = session.get("output_dir", "")
    path = os.path.join(output_dir, secure_filename(filename))
    if not os.path.exists(path):
        return "File not found", 404
    return send_file(path, as_attachment=True)


@app.route("/download-all")
def download_all():
    """Download all generated invoices as a ZIP."""
    output_dir = session.get("output_dir", "")
    if not output_dir or not os.path.exists(output_dir):
        return "No files generated yet", 404

    zip_path = os.path.join(get_session_dir(), "invoices.zip")
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname in os.listdir(output_dir):
            fpath = os.path.join(output_dir, fname)
            if os.path.isfile(fpath):
                zf.write(fpath, fname)

    return send_file(zip_path, as_attachment=True, download_name="invoices.zip")


if __name__ == "__main__":
    app.run(debug=True, port=5000)
