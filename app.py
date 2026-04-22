"""
Bulk Resume Downloader for Placement Teams
==========================================
A Flask web application that reads an Excel file with student names and
Google Drive resume links, then downloads all publicly accessible PDFs.
"""

import os
import re
import io
import time
import zipfile
import threading

import pandas as pd
import requests
from flask import Flask, render_template, request, jsonify, session
from werkzeug.utils import secure_filename

# ─── App Configuration ────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = "bulk_resume_downloader_secret_2024"

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
ALLOWED_EXTENSIONS = {"xlsx", "xls"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# In-memory progress store (thread-safe via lock)
progress_store: dict = {}
progress_lock = threading.Lock()


# ─── Utility Functions ─────────────────────────────────────────────────────────

def allowed_file(filename: str) -> bool:
    """Check if the uploaded file has an allowed extension."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def extract_drive_file_id(url: str) -> str | None:
    """
    Extract the Google Drive file ID from various Drive URL formats.
    Supports:
      - https://drive.google.com/file/d/FILE_ID/view
      - https://drive.google.com/open?id=FILE_ID
      - https://docs.google.com/...?id=FILE_ID
      - https://drive.google.com/uc?id=FILE_ID
    Returns the file ID string or None if not found.
    """
    if not url or not isinstance(url, str):
        return None

    url = url.strip()

    # Pattern 1: /file/d/FILE_ID/
    match = re.search(r"/file/d/([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)

    # Pattern 2: id=FILE_ID (query param)
    match = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)

    # Pattern 3: /d/FILE_ID (shorthand)
    match = re.search(r"/d/([a-zA-Z0-9_-]+)", url)
    if match:
        return match.group(1)

    return None


def build_download_url(file_id: str) -> str:
    """Convert a Google Drive file ID into a direct download URL."""
    return f"https://drive.google.com/uc?export=download&id={file_id}"


def sanitize_filename(name: str) -> str:
    """Remove characters that are invalid in file names."""
    # Remove any character that is not alphanumeric, space, dash, or underscore
    name = re.sub(r'[\\/*?:"<>|]', "", name).strip()
    return name if name else "unknown"


def unique_filepath(folder: str, base_name: str, ext: str = ".pdf") -> str:
    """
    Return a unique file path inside `folder`.
    If <base_name>.pdf already exists, append _1, _2, … until unique.
    """
    path = os.path.join(folder, base_name + ext)
    counter = 1
    while os.path.exists(path):
        path = os.path.join(folder, f"{base_name}_{counter}{ext}")
        counter += 1
    return path


def download_resume_bytes(url: str, timeout: int = 30) -> tuple[bool, str, bytes | None]:
    """
    Attempt to download a file from `url` into memory.
    Returns (success: bool, message: str, data: bytes | None).
    """
    try:
        response = requests.get(url, timeout=timeout, stream=True, allow_redirects=True)

        # Google Drive may redirect to a confirmation page for large files
        if response.status_code == 200:
            content_type = response.headers.get("Content-Type", "")
            # If we get an HTML page instead of a PDF, it's likely a restriction page
            if "text/html" in content_type:
                return False, "No Access (restricted or requires sign-in)", None

            buf = io.BytesIO()
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    buf.write(chunk)
            return True, "Downloaded", buf.getvalue()

        elif response.status_code in (401, 403, 404):
            return False, f"No Access (HTTP {response.status_code})", None
        else:
            return False, f"Failed (HTTP {response.status_code})", None

    except requests.exceptions.Timeout:
        return False, "Failed (Timeout)", None
    except requests.exceptions.ConnectionError:
        return False, "Failed (Connection Error)", None
    except requests.exceptions.RequestException as e:
        return False, f"Failed ({str(e)[:60]})", None


# ─── Background Download Worker ───────────────────────────────────────────────

MAX_WORKERS = 20   # simultaneous download threads


def _download_one(record: dict, name_counter: dict, counter_lock: threading.Lock) -> dict:
    """
    Download a single resume into memory. Called in a thread-pool worker.
    Returns a result dict with name, status, icon, link, and (on success) pdf_name + pdf_bytes.
    """
    name = record["name"]
    link = record["link"]

    file_id = extract_drive_file_id(link)
    if not file_id:
        return {"name": name, "status": "Invalid Link", "icon": "❌", "link": link,
                "pdf_name": None, "pdf_bytes": None}

    download_url = build_download_url(file_id)
    success, message, pdf_bytes = download_resume_bytes(download_url)

    if success and pdf_bytes:
        safe_name = sanitize_filename(name)
        # Generate a unique archive name without touching the disk
        with counter_lock:
            count = name_counter.get(safe_name, 0)
            name_counter[safe_name] = count + 1
        pdf_name = f"{safe_name}.pdf" if count == 0 else f"{safe_name}_{count}.pdf"

        return {
            "name": name,
            "status": f"Downloaded → {pdf_name}",
            "icon": "✅",
            "link": link,
            "pdf_name": pdf_name,
            "pdf_bytes": pdf_bytes,
        }

    icon = "⚠️" if "No Access" in message else "❌"
    return {"name": name, "status": message, "icon": icon, "link": link,
            "pdf_name": None, "pdf_bytes": None}


def run_downloads(task_id: str, records: list[dict]):
    """
    Background thread: spawns a pool of worker threads to download
    all resumes concurrently into memory, updating progress_store as each finishes.
    The final in-memory ZIP is built here and stored in progress_store.
    """
    from concurrent.futures import ThreadPoolExecutor, as_completed

    total      = len(records)
    results    = []          # display-safe results for the frontend
    pdf_files  = []          # [(pdf_name, pdf_bytes), …] for ZIP assembly
    downloaded = 0
    skipped    = 0
    failed     = 0
    lock       = threading.Lock()

    # Shared counter so duplicate names get _1, _2 suffixes
    name_counter: dict = {}
    counter_lock = threading.Lock()

    with progress_lock:
        progress_store[task_id] = {
            "status":   "running",
            "total":    total,
            "done":     0,
            "results": [],
            "summary": {},
            "zip_bytes": None,   # will hold final ZIP bytes when done
        }

    with ThreadPoolExecutor(max_workers=min(MAX_WORKERS, total)) as pool:
        futures = {
            pool.submit(_download_one, rec, name_counter, counter_lock): rec
            for rec in records
        }

        for future in as_completed(futures):
            result = future.result()

            with lock:
                # Strip internal bytes before sending to frontend
                display_result = {k: v for k, v in result.items()
                                  if k not in ("pdf_bytes", "pdf_name")}
                results.append(display_result)

                if result["icon"] == "✅":
                    downloaded += 1
                    if result.get("pdf_bytes") and result.get("pdf_name"):
                        pdf_files.append((result["pdf_name"], result["pdf_bytes"]))
                elif result["icon"] == "⚠️":
                    skipped += 1
                else:
                    failed += 1

                done_count = len(results)

            with progress_lock:
                progress_store[task_id]["done"]    = done_count
                progress_store[task_id]["results"] = list(results)

    # ── Build in-memory ZIP ─────────────────────────────────────────────────
    zip_bytes = None
    if pdf_files:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for pdf_name, pdf_data in pdf_files:
                zf.writestr(pdf_name, pdf_data)
        zip_bytes = zip_buf.getvalue()

    # ── Mark complete ───────────────────────────────────────────────────────
    with progress_lock:
        progress_store[task_id]["status"]    = "done"
        progress_store[task_id]["zip_bytes"] = zip_bytes
        progress_store[task_id]["summary"]   = {
            "total":      total,
            "downloaded": downloaded,
            "skipped":    skipped,
            "failed":     failed,
        }



# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    """Render the main upload page."""
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    """
    Handle Excel file upload. No folder path needed — files go into a
    per-task temp directory and are served back as a ZIP download.
    """
    # ── Validate file presence ──────────────────────────────────────────────
    if "excel_file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    file = request.files["excel_file"]

    if file.filename == "":
        return jsonify({"error": "No file selected."}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Invalid file type. Please upload an .xlsx or .xls file."}), 400

    # ── Save uploaded file ──────────────────────────────────────────────────
    filename = secure_filename(file.filename)
    excel_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(excel_path)

    # ── Parse Excel ─────────────────────────────────────────────────────────
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:
        return jsonify({"error": f"Failed to read Excel file: {e}"}), 400

    # ── Normalise column names: lowercase, strip whitespace, collapse spaces ──
    df.columns = [re.sub(r"\s+", " ", c.lower().strip()) for c in df.columns]
    actual_cols = [str(c) for c in df.columns]

    # Aliases accepted for the "name" column
    NAME_ALIASES = {
        "name", "student name", "student_name", "full name", "full_name",
        "candidate name", "candidate_name", "applicant", "applicant name",
        "applicant_name", "sname", "s.name", "roll name", "rollname",
    }
    # Aliases accepted for the "resume_link" column
    LINK_ALIASES = {
        "resume_link", "resume link", "resumelink", "resume url", "resume_url",
        "drive link", "drive_link", "drivelink", "google drive link",
        "google drive", "gdrive link", "gdrive_link", "link", "url",
        "resume", "cv link", "cv_link", "cv url", "cv_url",
        "file link", "file_link", "download link", "download_link",
    }

    name_col = next((c for c in actual_cols if c in NAME_ALIASES), None)
    link_col = next((c for c in actual_cols if c in LINK_ALIASES), None)

    if not name_col or not link_col:
        missing = []
        if not name_col:
            missing.append("name  (or: 'Student Name', 'Full Name', …)")
        if not link_col:
            missing.append("resume_link  (or: 'Resume Link', 'Drive Link', 'URL', …)")
        return jsonify({
            "error": (
                f"Excel is missing required column(s):\n• " +
                "\n• ".join(missing) +
                f"\n\nColumns found in your file: {', '.join(actual_cols)}"
            )
        }), 400

    if df.empty:
        return jsonify({"error": "The Excel file has no data rows."}), 400

    # Build list of (name, link) dicts, skipping fully empty rows
    records = []
    for _, row in df.iterrows():
        name = str(row.get(name_col, "")).strip()
        link = str(row.get(link_col, "")).strip()
        if name and link and name.lower() != "nan" and link.lower() != "nan":
            records.append({"name": name, "link": link})

    if not records:
        return jsonify({"error": "No valid rows found in the Excel file."}), 400

    # ── Start background task (all in-memory, no temp folder needed) ─────────
    task_id = str(time.time_ns())

    thread = threading.Thread(
        target=run_downloads,
        args=(task_id, records),
        daemon=True,
    )
    thread.start()

    return jsonify({"task_id": task_id, "total": len(records)})


@app.route("/progress/<task_id>")
def progress(task_id: str):
    """
    Polling endpoint: returns current progress for the given task.
    The frontend polls this every second.
    """
    with progress_lock:
        data = progress_store.get(task_id)

    if data is None:
        return jsonify({"error": "Task not found."}), 404

    return jsonify(data)


@app.route("/download-zip/<task_id>")
def download_zip(task_id: str):
    """
    Serve the pre-built in-memory ZIP for this task as 'resumes.zip'.
    The ZIP was assembled during the download phase, so no temp files are needed.
    """
    from flask import send_file

    with progress_lock:
        data = progress_store.get(task_id)

    if data is None:
        return jsonify({"error": "Task not found."}), 404

    if data.get("status") != "done":
        return jsonify({"error": "Downloads are still in progress."}), 400

    zip_bytes = data.get("zip_bytes")
    if not zip_bytes:
        return jsonify({"error": "No files were downloaded successfully."}), 400

    return send_file(
        io.BytesIO(zip_bytes),
        mimetype="application/zip",
        as_attachment=True,
        download_name="resumes.zip",
    )


# ─── Entry Point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
