import json
import os
from collections import OrderedDict

import pandas as pd
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
from werkzeug.utils import secure_filename

from mainframe import agentic_flow, generate_from_ado, generate_from_excel, upload_test_cases_ado

app = Flask(__name__)
CORS(app)
os.makedirs("output", exist_ok=True)


# ── Helpers ──────────────────────────────────────────────────────────────────

def _err(msg, code=400):
    return jsonify({"status": "error", "message": msg}), code

def _save_image(file_obj) -> str:
    import os
    from werkzeug.utils import secure_filename
    ext = os.path.splitext(secure_filename(file_obj.filename))[1]
    path = f"temp_image{ext}"
    file_obj.save(path)
    return path

def _clean(path):
    if path and os.path.exists(path):
        os.remove(path)


# ── Routes ───────────────────────────────────────────────────────────────────

@app.route("/generate_excel", methods=["POST"])
def route_generate_excel():
    if "file" not in request.files or not request.files["file"].filename:
        return _err("No file provided")
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return _err("Only .xlsx allowed")

    f.save("temp_input.xlsx")
    image_path = None
    try:
        if "image" in request.files and request.files["image"].filename:
            image_path = _save_image(request.files["image"])
        out = generate_from_excel("temp_input.xlsx", image_path=image_path)
        return jsonify({"status": "success", "count": len(pd.read_excel(out)), "filename": os.path.basename(out)})
    except Exception as e:
        return _err(str(e), 500)
    finally:
        _clean("temp_input.xlsx")
        _clean(image_path)


@app.route("/generate_ado", methods=["POST"])
def route_generate_ado():
    d = request.get_json() or {}
    story_id = str(d.get("story_id", "")).strip()
    org, project, pat = d.get("org", "").strip(), d.get("project", "").strip(), d.get("pat", "").strip()
    if not all([story_id, org, project, pat]):
        return _err("story_id, org, project, and pat are required")
    try:
        out = generate_from_ado(story_id, org, project, pat)
        return jsonify({"status": "success", "count": len(pd.read_excel(out)), "filename": os.path.basename(out)})
    except Exception as e:
        return _err(str(e), 500)


@app.route("/agentic_flow", methods=["POST"])
def route_agentic_flow():
    d = request.get_json() or {}
    story_id = str(d.get("story_id", "")).strip()
    org, project, pat = d.get("org", "").strip(), d.get("project", "").strip(), d.get("pat", "").strip()
    if not all([story_id, org, project, pat]):
        return _err("story_id, org, project, and pat are required")
    try:
        summary = agentic_flow(story_id, org, project, pat)
        return jsonify({"status": "success", **summary})
    except Exception as e:
        return _err(str(e), 500)


@app.route("/upload", methods=["POST"])
def route_upload():
    if "file" not in request.files or not request.files["file"].filename:
        return _err("No file provided")
    f = request.files["file"]
    if not f.filename.endswith(".xlsx"):
        return _err("Only .xlsx allowed")

    org = request.form.get("org")
    proj = request.form.get("project")
    pat = request.form.get("pat")
    plan = request.form.get("plan_name")
    suite = request.form.get("suite_name", "LOGIN")

    if not all([org, proj, pat, plan]):
        return _err("org, project, pat, plan_name are required")

    f.save("temp_output.xlsx")
    try:
        uploaded, failed = upload_test_cases_ado("temp_output.xlsx", org, proj, pat, plan, suite)
        status = "success" if uploaded > 0 else "fail"
        return jsonify({"status": status, "uploaded_count": uploaded, "failed_count": failed})
    except Exception as e:
        return _err(str(e), 500)
    finally:
        _clean("temp_output.xlsx")


@app.route("/download", methods=["GET"])
def route_download():
    filename = request.args.get("filename")
    if not filename:
        return _err("filename required")
    path = os.path.join("output", secure_filename(filename))
    if not os.path.exists(path):
        return _err("File not found", 404)
    return send_file(path, as_attachment=True, download_name=filename)


@app.route("/download-template", methods=["GET"])
def route_download_template():
    if not os.path.exists("template.xlsx"):
        return _err("Template not found", 404)
    return send_file("template.xlsx", as_attachment=True, download_name="template.xlsx")


@app.route("/get-test-cases", methods=["GET"])
def route_get_test_cases():
    filename = request.args.get("filename")
    if not filename:
        return _err("filename required")
    path = os.path.join("output", secure_filename(filename))
    if not os.path.exists(path):
        return _err("File not found", 404)

    df = pd.read_excel(path)
    col_order = ["S.No.", "User Story", "Acceptance Criteria", "Title", "Steps", "Priority", "Test Type"]
    cols = [c for c in col_order if c in df.columns]
    df = df.where(pd.notnull(df), None)

    test_cases = [
        OrderedDict((c, row[c]) for c in cols)
        for _, row in df.iterrows()
    ]
    return app.response_class(
        response=json.dumps({"status": "success", "test_cases": test_cases}, ensure_ascii=False),
        mimetype="application/json",
    )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)