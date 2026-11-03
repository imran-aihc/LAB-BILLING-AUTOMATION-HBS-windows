"""
HBS HIMS — Flask API Server
Run: python server.py
"""
import asyncio, os, threading
from pathlib import Path
from flask import Flask, jsonify, request, send_file, Response
from flask_cors import CORS
from automation import run_pipeline, load_admission_excel

app = Flask(__name__)
CORS(app)

UPLOAD_DIR  = Path("uploads");  UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_FILE = Path("output/HBS_Lab_Bill.xlsx")

# ── Shared job state ──────────────────────────────────────────────────────────
job = {
    "running": False, "total": 0, "done": 0,
    "current": "", "logs": [], "finished": False,
    "output_ready": False, "success": 0, "no_labs": 0, "load_failed": 0, "failed": 0,
}

def reset_job():
    job.update(running=False, total=0, done=0, current="",
               logs=[], finished=False, output_ready=False,
               success=0, no_labs=0, load_failed=0, failed=0)

# ── Endpoints ─────────────────────────────────────────────────────────────────
@app.route("/api/health")
def health():
    return jsonify({"ok": True})

@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file"}), 400
    f = request.files["file"]
    path = UPLOAD_DIR / "admissions.xlsx"
    f.save(str(path))
    try:
        patients = load_admission_excel(str(path))
        return jsonify({"ok": True, "count": len(patients),
                        "preview": patients[:8]})
    except Exception as e:
        return jsonify({"error": str(e)}), 400

@app.route("/api/status")
def status():
    return jsonify(job)

@app.route("/api/start", methods=["POST"])
def start():
    if job["running"]:
        return jsonify({"error": "Already running"}), 400
    data  = request.json or {}
    month = data.get("month", "May 26")
    excel = str(UPLOAD_DIR / "admissions.xlsx")
    if not Path(excel).exists():
        return jsonify({"error": "Upload Excel first"}), 400

    reset_job()
    job["running"] = True

    # Count patients for progress bar
    try:
        from automation import load_admission_excel as lae
        job["total"] = len(lae(excel))
    except Exception:
        pass

    def progress_cb(msg: str):
        job["logs"].append(msg)
        # Parse progress signals
        if msg.startswith("["):
            try:
                done = int(msg.split("[")[1].split("/")[0])
                job["done"] = done
            except Exception:
                pass
        if "Processing:" in msg:
            job["current"] = msg.split("Processing:")[-1].strip()
        if "with labs" in msg.lower() and "Done" in msg:
            try:
                job["success"]     = int(msg.split("with labs")[0].split("|")[-1].strip())
                job["no_labs"]     = int(msg.split("no labs")[0].split("|")[-1].strip())
                job["load_failed"] = int(msg.split("need recheck")[0].split("|")[-1].strip())
                job["failed"]      = int(msg.split("failed")[0].split("|")[-1].strip())
            except Exception:
                pass

    def run():
        try:
            asyncio.run(run_pipeline(excel, month, cb=progress_cb))
            job["output_ready"] = True
        except Exception as e:
            job["logs"].append(f"❌ Fatal: {e}")
        finally:
            job["running"]  = False
            job["finished"] = True

    threading.Thread(target=run, daemon=True).start()
    return jsonify({"ok": True})

@app.route("/api/download")
def download():
    if not OUTPUT_FILE.exists():
        return jsonify({"error": "No file yet"}), 404
    return send_file(str(OUTPUT_FILE.resolve()),
                     download_name="HBS_Lab_Bill.xlsx",
                     as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port, debug=False)
