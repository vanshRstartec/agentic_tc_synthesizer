import json
import os
import queue
import threading
import uuid
from collections import OrderedDict

import pandas as pd
from flask import Flask, Response, jsonify, request, send_file, stream_with_context
from flask_cors import CORS
from werkzeug.utils import secure_filename

import mainframe as mf
from mainframe import agentic_flow, agentic_reference, agentic_clone

app = Flask(__name__)
CORS(app)
os.makedirs("output", exist_ok=True)

_jobs: dict = {}


def _err(msg, code=400):
    return jsonify({"status": "error", "message": msg}), code


def _start_job(fn, **kwargs) -> str:
    job_id = str(uuid.uuid4())
    q      = queue.Queue()
    _jobs[job_id] = {"queue": q, "result": None, "error": None}

    def _run():
        mf.set_log_queue(q)
        try:
            _jobs[job_id]["result"] = fn(**kwargs)
        except Exception as e:
            _jobs[job_id]["error"] = str(e)
        finally:
            mf.set_log_queue(None)
            q.put(None)

    threading.Thread(target=_run, daemon=True).start()
    return job_id


def _base_params(d: dict) -> dict:
    """Extract and validate common params shared by all three endpoints."""
    return {
        "story_id":            str(d.get("story_id", "")).strip(),
        "org":                 d.get("org", "").strip(),
        "project":             d.get("project", "").strip(),
        "pat":                 d.get("pat", "").strip(),
        "plan_name_override":  d.get("plan_name_override",  "").strip() or None,
        "suite_name_override": d.get("suite_name_override", "").strip() or None,
    }


def _suite_params(d: dict) -> dict:
    return {
        "ref_plan_id":  str(d.get("ref_plan_id",  "")).strip(),
        "ref_suite_id": str(d.get("ref_suite_id", "")).strip(),
    }


@app.route("/agentic_flow", methods=["POST"])
def route_agentic_flow():
    d      = request.get_json() or {}
    params = _base_params(d)
    if not all([params["story_id"], params["org"], params["project"], params["pat"]]):
        return _err("story_id, org, project, and pat are required")
    return jsonify({"status": "started", "job_id": _start_job(agentic_flow, **params)})


@app.route("/agentic_reference", methods=["POST"])
def route_agentic_reference():
    d      = request.get_json() or {}
    params = {**_base_params(d), **_suite_params(d)}
    if not all([params["story_id"], params["org"], params["project"], params["pat"],
                params["ref_plan_id"], params["ref_suite_id"]]):
        return _err("story_id, org, project, pat, ref_plan_id, and ref_suite_id are required")
    return jsonify({"status": "started", "job_id": _start_job(agentic_reference, **params)})


@app.route("/agentic_clone", methods=["POST"])
def route_agentic_clone():
    d      = request.get_json() or {}
    params = {**_base_params(d), **_suite_params(d)}
    if not all([params["story_id"], params["org"], params["project"], params["pat"],
                params["ref_plan_id"], params["ref_suite_id"]]):
        return _err("story_id, org, project, pat, ref_plan_id, and ref_suite_id are required")
    return jsonify({"status": "started", "job_id": _start_job(agentic_clone, **params)})


@app.route("/agentic_flow_logs/<job_id>", methods=["GET"])
def route_agentic_flow_logs(job_id):
    job = _jobs.get(job_id)
    if not job:
        return _err("Job not found", 404)

    def _generate():
        q = job["queue"]
        while True:
            try:
                msg = q.get(timeout=120)
            except queue.Empty:
                yield "data: ⏱ Timeout waiting for next log line.\n\n"
                break
            if msg is None:
                result = job.get("result")
                error  = job.get("error")
                yield f"data: __RESULT__{json.dumps(result)}\n\n" if result \
                    else f"data: __ERROR__{error or 'Unknown error'}\n\n"
                _jobs.pop(job_id, None)
                break
            yield f"data: {str(msg).replace(chr(10), ' ').replace(chr(13), '')}\n\n"

    return Response(
        stream_with_context(_generate()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no",
                 "Access-Control-Allow-Origin": "*"},
    )


@app.route("/download", methods=["GET"])
def route_download():
    filename = request.args.get("filename")
    if not filename:
        return _err("filename required")
    path = os.path.join("output", secure_filename(filename))
    if not os.path.exists(path):
        return _err("File not found", 404)
    return send_file(path, as_attachment=True, download_name=filename)


@app.route("/get-test-cases", methods=["GET"])
def route_get_test_cases():
    filename = request.args.get("filename")
    if not filename:
        return _err("filename required")
    path = os.path.join("output", secure_filename(filename))
    if not os.path.exists(path):
        return _err("File not found", 404)

    df        = pd.read_excel(path)
    col_order = ["S.No.", "User Story", "Title", "Steps", "Priority", "Test Type"]
    cols      = [c for c in col_order if c in df.columns]
    df        = df.where(pd.notnull(df), None)
    test_cases = [OrderedDict((c, row[c]) for c in cols) for _, row in df.iterrows()]

    return app.response_class(
        response=json.dumps({"status": "success", "test_cases": test_cases}, ensure_ascii=False),
        mimetype="application/json",
    )


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)