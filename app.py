#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Taxis & Logico-Semantic Shift Tool — Web Version
Flask backend: PDF→TXT, EN Analysis, AR Analysis, Shifts
"""

import os
import sys
import uuid
import json
import queue
import threading
import traceback
from pathlib import Path
from flask import (Flask, render_template, request, jsonify,
                   send_file, Response, stream_with_context)
from werkzeug.utils import secure_filename

# ── ensure modules are importable ────────────────────────────────────────────
HERE = os.path.dirname(os.path.abspath(__file__))
MODULES = os.path.join(HERE, "modules")
if MODULES not in sys.path:
    sys.path.insert(0, MODULES)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB

UPLOAD_DIR = os.path.join(HERE, "uploads")
OUTPUT_DIR = os.path.join(HERE, "outputs")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# job_id → queue of SSE messages
_job_queues = {}


# ── helpers ───────────────────────────────────────────────────────────────────

def _new_job():
    job_id = str(uuid.uuid4())
    _job_queues[job_id] = queue.Queue()
    return job_id


def _push(job_id, event, data):
    if job_id in _job_queues:
        _job_queues[job_id].put({"event": event, "data": data})


def _save_upload(file_obj, suffix=""):
    fname = secure_filename(file_obj.filename)
    stem  = Path(fname).stem
    ext   = Path(fname).suffix
    uid   = str(uuid.uuid4())[:8]
    path  = os.path.join(UPLOAD_DIR, f"{stem}_{uid}{suffix}{ext}")
    file_obj.save(path)
    return path


def _out_path(stem, suffix):
    uid = str(uuid.uuid4())[:8]
    return os.path.join(OUTPUT_DIR, f"{stem}_{uid}{suffix}.xlsx")


# ── SSE stream ────────────────────────────────────────────────────────────────

@app.route("/stream/<job_id>")
def stream(job_id):
    def _generate():
        q = _job_queues.get(job_id)
        if not q:
            yield f"event: error\ndata: Job not found\n\n"
            return
        while True:
            try:
                msg = q.get(timeout=120)
                yield f"event: {msg['event']}\ndata: {json.dumps(msg['data'])}\n\n"
                if msg["event"] in ("done", "error"):
                    _job_queues.pop(job_id, None)
                    break
            except queue.Empty:
                yield "event: ping\ndata: {}\n\n"

    return Response(stream_with_context(_generate()),
                    mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache",
                             "X-Accel-Buffering": "no"})


# ── download ──────────────────────────────────────────────────────────────────

@app.route("/download/<path:filename>")
def download(filename):
    full = os.path.join(OUTPUT_DIR, os.path.basename(filename))
    if not os.path.exists(full):
        return jsonify({"error": "File not found"}), 404
    return send_file(full, as_attachment=True)


# ── PDF → TXT ─────────────────────────────────────────────────────────────────

@app.route("/api/pdf", methods=["POST"])
def api_pdf():
    if "pdf" not in request.files:
        return jsonify({"error": "No PDF uploaded"}), 400

    pdf_file = request.files["pdf"]
    lang     = request.form.get("lang", "eng")
    job_id   = _new_job()

    pdf_path  = _save_upload(pdf_file)
    stem      = Path(pdf_file.filename).stem
    txt_path  = os.path.join(OUTPUT_DIR, f"{stem}_{job_id[:8]}.txt")
    xlsx_path = _out_path(stem, "_sentences")

    def _run():
        try:
            def _cb(cur, tot):
                pct = int(cur / tot * 100)
                _push(job_id, "progress",
                      {"msg": f"Page {cur} of {tot}", "pct": pct})

            if lang == "eng":
                from pdf_text_eng import pdf_to_text_excel_english
                pdf_to_text_excel_english(pdf_path, txt_path, xlsx_path,
                                          progress_cb=_cb)
            else:
                from pdf_text_ar import pdf_to_text_excel
                pdf_to_text_excel(pdf_path, txt_path, xlsx_path,
                                  progress_cb=_cb)

            _push(job_id, "done", {
                "msg": "Conversion complete!",
                "file": os.path.basename(xlsx_path)
            })
        except Exception as e:
            _push(job_id, "error", {"msg": str(e)})
        finally:
            if os.path.exists(pdf_path):
                os.remove(pdf_path)

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"job_id": job_id})


# ── EN Analysis ───────────────────────────────────────────────────────────────

@app.route("/api/en_analysis", methods=["POST"])
def api_en_analysis():
    if "excel" not in request.files:
        return jsonify({"error": "No Excel file uploaded"}), 400

    excel_file     = request.files["excel"]
    include_embed  = request.form.get("embedded", "false") == "true"
    job_id         = _new_job()

    in_path  = _save_upload(excel_file)
    stem     = Path(excel_file.filename).stem
    out_path = _out_path(stem, "_taxis_en")

    def _run():
        try:
            _push(job_id, "progress", {"msg": "Loading model …", "pct": 10})
            from Eng_Analysis import process_excel_file
            _push(job_id, "progress", {"msg": "Analyzing sentences …", "pct": 30})
            process_excel_file(in_path, out_path,
                               include_embedded=include_embed)
            _push(job_id, "done", {
                "msg": "Analysis complete!",
                "file": os.path.basename(out_path)
            })
        except Exception as e:
            _push(job_id, "error", {"msg": str(e)})
        finally:
            if os.path.exists(in_path):
                os.remove(in_path)

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"job_id": job_id})


# ── AR Analysis ───────────────────────────────────────────────────────────────

@app.route("/api/ar_analysis", methods=["POST"])
def api_ar_analysis():
    if "excel" not in request.files:
        return jsonify({"error": "No Excel file uploaded"}), 400

    excel_file = request.files["excel"]
    job_id     = _new_job()

    in_path  = _save_upload(excel_file)
    stem     = Path(excel_file.filename).stem
    out_path = _out_path(stem, "_taxis_ar")

    def _run():
        try:
            _push(job_id, "progress", {"msg": "Analyzing Arabic sentences …", "pct": 20})
            from Ar_Analysis import process_excel
            process_excel(in_path, out_path)
            _push(job_id, "done", {
                "msg": "Analysis complete!",
                "file": os.path.basename(out_path)
            })
        except Exception as e:
            _push(job_id, "error", {"msg": str(e)})
        finally:
            if os.path.exists(in_path):
                os.remove(in_path)

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"job_id": job_id})


# ── Shifts ────────────────────────────────────────────────────────────────────

@app.route("/api/shifts", methods=["POST"])
def api_shifts():
    if "en_excel" not in request.files or "ar_excel" not in request.files:
        return jsonify({"error": "Both Excel files required"}), 400

    en_file = request.files["en_excel"]
    ar_file = request.files["ar_excel"]
    job_id  = _new_job()

    en_path  = _save_upload(en_file, "_en")
    ar_path  = _save_upload(ar_file, "_ar")
    stem     = Path(en_file.filename).stem
    out_path = _out_path(stem, "_shifts")

    def _run():
        try:
            _push(job_id, "progress", {"msg": "Running shift analysis …", "pct": 20})
            from Shift import run_shift_analysis
            run_shift_analysis(en_path, ar_path, out_path)
            _push(job_id, "done", {
                "msg": "Shift analysis complete!",
                "file": os.path.basename(out_path)
            })
        except Exception as e:
            _push(job_id, "error", {"msg": str(e)})
        finally:
            for p in (en_path, ar_path):
                if os.path.exists(p):
                    os.remove(p)

    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"job_id": job_id})


# ── Index ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=5001)
