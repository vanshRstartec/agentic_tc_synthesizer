"""
Flask API for the Agentic Test Case Synthesizer.

Endpoints
─────────
POST /agentic_flow            standard generation + review + upload
POST /agentic_reference       reference-suite-guided generation + review + upload
POST /agentic_clone           clone-from-suite + upload (no review)
GET  /agentic_flow_logs/<id>  Server-Sent Events stream of live job logs
GET  /download                download generated .xlsx
GET  /get-test-cases          read generated .xlsx as JSON
"""
from __future__ import annotations

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
from mainframe import agentic_clone, agentic_flow, agentic_reference

app = Flask(__name__)
CORS(app)

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── Job registry ──────────────────────────────────────────────────────────────
_jobs: dict[str, dict] = {}
_LOG_TIMEOUT_S = 1200


def _err(msg: str, code: int = 400):
    return jsonify({"status": "error", "message": msg}), code


def _start_job(fn, **kwargs) -> str:
    """Run an agentic pipeline on a background thread; capture logs into a queue."""
    job_id = str(uuid.uuid4())
    q      = queue.Queue()
    _jobs[job_id] = {"queue": q, "result": None, "error": None}

    def _runner():
        mf.set_log_queue(q)
        try:
            _jobs[job_id]["result"] = fn(**kwargs)
        except Exception as e:
            _jobs[job_id]["error"] = str(e)
            mf._log(f"❌ {e}")
        finally:
            mf.set_log_queue(None)
            q.put(None)

    threading.Thread(target=_runner, daemon=True).start()
    return job_id


def _payload(d: dict, *, with_suite: bool = False) -> tuple[dict, list[str]]:
    """Pull and trim params from request body. Returns (params, missing_keys)."""
    params = {
        "story_id":            str(d.get("story_id", "")).strip(),
        "org":                 str(d.get("org", "")).strip(),
        "project":             str(d.get("project", "")).strip(),
        "pat":                 str(d.get("pat", "")).strip(),
        "plan_name_override":  str(d.get("plan_name_override",  "")).strip() or None,
        "suite_name_override": str(d.get("suite_name_override", "")).strip() or None,
    }
    required = ["story_id", "org", "project", "pat"]
    if with_suite:
        params["ref_plan_id"]  = str(d.get("ref_plan_id",  "")).strip()
        params["ref_suite_id"] = str(d.get("ref_suite_id", "")).strip()
        required += ["ref_plan_id", "ref_suite_id"]
    missing = [k for k in required if not params[k]]
    return params, missing


def _safe_path(filename: str) -> str | None:
    """Resolve a download/read filename to a safe path under OUTPUT_DIR."""
    if not filename:
        return None
    path = os.path.join(OUTPUT_DIR, secure_filename(filename))
    return path if os.path.exists(path) else None


# ── Agentic endpoints (factory) ───────────────────────────────────────────────
def _make_endpoint(fn, *, with_suite: bool):
    def handler():
        body = request.get_json(silent=True) or {}
        params, missing = _payload(body, with_suite=with_suite)
        if missing:
            return _err(f"Missing required field(s): {', '.join(missing)}")
        return jsonify({"status": "started", "job_id": _start_job(fn, **params)})
    handler.__name__ = f"route_{fn.__name__}"
    return handler


app.add_url_rule("/agentic_flow",      view_func=_make_endpoint(agentic_flow,      with_suite=False), methods=["POST"])
app.add_url_rule("/agentic_reference", view_func=_make_endpoint(agentic_reference, with_suite=True),  methods=["POST"])
app.add_url_rule("/agentic_clone",     view_func=_make_endpoint(agentic_clone,     with_suite=True),  methods=["POST"])


# ── SSE log stream ────────────────────────────────────────────────────────────
@app.route("/agentic_flow_logs/<job_id>", methods=["GET"])
def route_logs(job_id: str):
    job = _jobs.get(job_id)
    if not job:
        return _err("Job not found", 404)

    def _stream():
        q = job["queue"]
        while True:
            try:
                msg = q.get(timeout=_LOG_TIMEOUT_S)
            except queue.Empty:
                yield "data: ⚠️ Timed out waiting for next log line.\n\n"
                break

            if msg is None:
                if job.get("result"):
                    yield f"data: __RESULT__{json.dumps(job['result'])}\n\n"
                else:
                    yield f"data: __ERROR__{job.get('error') or 'Unknown error'}\n\n"
                _jobs.pop(job_id, None)
                break

            # Strip newlines so each msg is a single SSE record
            clean = str(msg).replace("\n", " ").replace("\r", "")
            yield f"data: {clean}\n\n"

    return Response(
        stream_with_context(_stream()),
        mimetype="text/event-stream",
        headers={
            "Cache-Control":             "no-cache",
            "X-Accel-Buffering":         "no",
            "Access-Control-Allow-Origin": "*",
        },
    )


# ── File routes ───────────────────────────────────────────────────────────────
@app.route("/download", methods=["GET"])
def route_download():
    filename = request.args.get("filename")
    path     = _safe_path(filename or "")
    if not path:
        return _err("File not found", 404)
    return send_file(path, as_attachment=True, download_name=filename)


@app.route("/get-test-cases", methods=["GET"])
def route_get_test_cases():
    filename = request.args.get("filename")
    path     = _safe_path(filename or "")
    if not path:
        return _err("File not found", 404)

    df        = pd.read_excel(path).where(lambda x: x.notnull(), None)
    col_order = ["S.No.", "User Story", "Title", "Steps", "Priority", "Test Type"]
    cols      = [c for c in col_order if c in df.columns]
    test_cases = [OrderedDict((c, row[c]) for c in cols) for _, row in df.iterrows()]

    return app.response_class(
        response=json.dumps({"status": "success", "test_cases": test_cases}, ensure_ascii=False),
        mimetype="application/json",
    )


# ── Health ────────────────────────────────────────────────────────────────────
@app.route("/", methods=["GET"])
def route_health():
    return jsonify({"status": "ok", "service": "tcs-synthesizer"})


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)