"""UAR Testing API - FastAPI backend (runs separately from the frontend).

Start:  uvicorn app:app --host 0.0.0.0 --port 8000
Config: backend/.env (Azure OpenAI + UAR_* settings) - see .env.example
"""
import json
import os
import pickle
import re
import shutil
import threading
import traceback
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from typing import List, Optional

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

import uar_engine as E
from chat import RunDB, answer_question

BASE = os.path.dirname(os.path.abspath(__file__))
E.load_env([os.path.join(BASE, ".env")])

RUNS_DIR = os.environ.get("UAR_RUNS_DIR") or os.path.join(BASE, "runs")
ORIGINS = [o.strip() for o in os.environ.get(
    "UAR_ALLOWED_ORIGINS", "http://localhost:5500,http://127.0.0.1:5500").split(",") if o.strip()]
MAX_MB = int(os.environ.get("UAR_MAX_UPLOAD_MB", "300"))
RETENTION_DAYS = int(os.environ.get("UAR_RETENTION_DAYS", "30"))
os.makedirs(RUNS_DIR, exist_ok=True)

# upload slot (form field) -> engine key, allowed extensions, required, label
SLOTS = {
    "start_listing":     ("user_report_xlsx", {".xlsx", ".xlsm"}, True, "Start-of-review user listing (Excel)"),
    "start_listing_pdf": ("user_report_pdf", {".pdf"}, False, "Start-of-review user listing (PDF)"),
    "review_file":       ("review_xlsx", {".xlsx", ".xlsm"}, True, "Manager review file"),
    "post_listing":      ("after_xlsx", {".xlsx", ".xlsm"}, True, "Post-review user listing (Excel)"),
    "post_listing_pdf":  ("after_pdf", {".pdf"}, False, "Post-review user listing (PDF)"),
    "removed_accounts":  ("removed_xlsx", {".xlsx", ".xlsm"}, False, "Removed accounts analysis"),
    "extraction_steps":  ("steps_doc", {".docx", ".doc", ".pdf", ".txt"}, False, "Report extraction steps"),
}
EMAIL_EXT = {".msg", ".eml"}
RUN_ID = re.compile(r"^\d{8}-\d{6}-[0-9a-f]{6}$")

app = FastAPI(title="UAR Testing API", version="1.0")
app.add_middleware(CORSMiddleware, allow_origins=ORIGINS, allow_methods=["GET", "POST", "DELETE"],
                   allow_headers=["*"])

EXECUTOR = ThreadPoolExecutor(max_workers=1)      # one run at a time: engine settings are process-wide
_state_lock = threading.Lock()
_results_cache, _db_cache = {}, {}


# ----------------------------------------------------------------------------- helpers
def run_dir(run_id):
    if not RUN_ID.match(run_id or ""):
        raise HTTPException(400, "Invalid run id")
    d = os.path.join(RUNS_DIR, run_id)
    if not os.path.isdir(d):
        raise HTTPException(404, "Run not found")
    return d


def read_meta(d):
    with open(os.path.join(d, "meta.json"), encoding="utf-8") as fh:
        return json.load(fh)


def write_meta(d, meta):
    tmp = os.path.join(d, "meta.json.tmp")
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(meta, fh, indent=1, default=str)
    os.replace(tmp, os.path.join(d, "meta.json"))


def update_meta(d, **kw):
    with _state_lock:
        m = read_meta(d)
        m.update(kw)
        write_meta(d, m)


def append_log(d, line):
    with _state_lock:
        with open(os.path.join(d, "log.txt"), "a", encoding="utf-8") as fh:
            fh.write(line.rstrip() + "\n")


def safe_name(name):
    name = os.path.basename(name or "file").replace("\\", "_")
    name = re.sub(r"[^\w.\- ()&,+]", "_", name).strip(" .") or "file"
    return name[:150]


def save_upload(up: UploadFile, dest_dir):
    os.makedirs(dest_dir, exist_ok=True)
    path = os.path.join(dest_dir, safe_name(up.filename))
    size, limit = 0, MAX_MB * 1024 * 1024
    with open(path, "wb") as out:
        while True:
            chunk = up.file.read(1024 * 1024)
            if not chunk:
                break
            size += len(chunk)
            if size > limit:
                out.close()
                os.remove(path)
                raise HTTPException(413, f"{up.filename} is larger than {MAX_MB} MB")
            out.write(chunk)
    return path


def extract_emails_zip(zpath, dest):
    """Extract only .msg/.eml files, flattened, with no path traversal and a size cap."""
    n, total, limit = 0, 0, MAX_MB * 1024 * 1024
    with zipfile.ZipFile(zpath) as z:
        for info in z.infolist():
            if info.is_dir() or os.path.splitext(info.filename)[1].lower() not in EMAIL_EXT:
                continue
            total += info.file_size
            if total > limit:
                raise HTTPException(413, "E-mail zip expands beyond the upload limit")
            base, ext = os.path.splitext(safe_name(info.filename))
            target, k = os.path.join(dest, base + ext), 1
            while os.path.exists(target):
                target, k = os.path.join(dest, f"{base} ({k}){ext}"), k + 1
            with z.open(info) as src, open(target, "wb") as dst:
                shutil.copyfileobj(src, dst)
            n += 1
    return n


def js(v):
    """JSON-safe cell value; dates shown MM-DD-YYYY (client format)."""
    if v is None or v is pd.NaT or v is pd.NA:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        t = pd.Timestamp(v)
        return t.strftime("%m-%d-%Y") if (t.hour, t.minute) == (0, 0) else t.strftime("%m-%d-%Y %H:%M")
    if hasattr(v, "item"):
        return v.item()
    return v


def table_payload(df):
    return {"columns": [str(c) for c in df.columns], "rows": [[js(v) for v in row] for row in df.values.tolist()]}


def load_results(run_id):
    if run_id not in _results_cache:
        p = os.path.join(run_dir(run_id), "results.pkl")
        if not os.path.isfile(p):
            raise HTTPException(409, "Results are not ready for this run")
        with open(p, "rb") as fh:
            _results_cache[run_id] = pickle.load(fh)
    return _results_cache[run_id]


def ai_status():
    ai = E.AzureAI(None)
    return {"enabled": ai.enabled, "deployment": ai.deployment if ai.enabled else None}


def cleanup_old_runs():
    cutoff = datetime.now() - timedelta(days=RETENTION_DAYS)
    for name in os.listdir(RUNS_DIR):
        d = os.path.join(RUNS_DIR, name)
        if RUN_ID.match(name) and os.path.isdir(d):
            try:
                if datetime.strptime(name[:15], "%Y%m%d-%H%M%S") < cutoff:
                    shutil.rmtree(d, ignore_errors=True)
            except ValueError:
                pass


# ----------------------------------------------------------------------------- processing
def process(run_id):
    d = os.path.join(RUNS_DIR, run_id)
    meta = read_meta(d)
    update_meta(d, status="running", started=datetime.now().isoformat(timespec="seconds"))
    E._LOG[0] = lambda m: append_log(d, str(m))
    try:
        opts = dict(meta["options"])
        opts["cache_path"] = os.path.join(d, "ai_cache.json")
        out = os.path.join(d, "output", f"UAR Testing Result - {safe_name(opts.get('app_name') or 'UAR')}.xlsx")
        os.makedirs(os.path.dirname(out), exist_ok=True)
        files = {k: v for k, v in meta["inputs"].items() if v}
        res = E.run_files(files, out, opts)
        with open(os.path.join(d, "results.pkl"), "wb") as fh:
            pickle.dump(res, fh)
        conclusion = next((s["value"] for s in res["summary"] if s["metric"] == "Conclusion"), "")
        update_meta(d, status="done", finished=datetime.now().isoformat(timespec="seconds"),
                    output=os.path.basename(out), conclusion=conclusion)
        append_log(d, "Done.")
    except Exception as e:
        append_log(d, f"ERROR: {e}")
        with open(os.path.join(d, "error.txt"), "w", encoding="utf-8") as fh:
            fh.write(traceback.format_exc())
        update_meta(d, status="failed", error=str(e), finished=datetime.now().isoformat(timespec="seconds"))
    finally:
        E._LOG[0] = print
        cleanup_old_runs()


# ----------------------------------------------------------------------------- API
@app.get("/api/health")
def health():
    return {"ok": True, "ai": ai_status(), "max_upload_mb": MAX_MB, "retention_days": RETENTION_DAYS,
            "slots": {k: {"label": v[3], "required": v[2], "extensions": sorted(v[1])} for k, v in SLOTS.items()},
            "approvals_extensions": sorted(EMAIL_EXT | {".zip"})}


def _clean_options(raw):
    try:
        o = json.loads(raw or "{}")
    except ValueError:
        raise HTTPException(400, "Settings are not valid JSON")
    out = {}
    for k in ("app_name", "control_owner", "status_approve", "status_remove"):
        v = str(o.get(k) or "").strip()
        if v:
            out[k] = v[:80]
    for k in ("due_date", "start_date", "after_date"):
        v = str(o.get(k) or "").strip()
        if v:
            if not re.match(r"^\d{4}-\d{2}-\d{2}$", v):
                raise HTTPException(400, f"{k} must be YYYY-MM-DD")
            out[k] = v
    if str(o.get("response_days") or "").strip():
        try:
            out["response_days"] = max(1, min(365, int(o["response_days"])))
        except (TypeError, ValueError):
            raise HTTPException(400, "Response window must be a whole number of days")
    if str(o.get("dormant_days") or "").strip():
        try:
            out["dormant_days"] = max(1, min(3650, int(o["dormant_days"])))
        except (TypeError, ValueError):
            raise HTTPException(400, "Dormant threshold must be a whole number of days")
    codes = o.get("elevated_codes")
    if isinstance(codes, str):
        codes = [c for c in re.split(r"[,\s;]+", codes) if c]
    if codes:
        out["elevated_codes"] = [str(c)[:20] for c in codes][:50]
    out["use_ai"] = bool(o.get("use_ai", True))
    return out


@app.post("/api/runs")
def create_run(
    start_listing: UploadFile = File(...),
    review_file: UploadFile = File(...),
    post_listing: UploadFile = File(...),
    start_listing_pdf: Optional[UploadFile] = File(None),
    post_listing_pdf: Optional[UploadFile] = File(None),
    removed_accounts: Optional[UploadFile] = File(None),
    extraction_steps: Optional[UploadFile] = File(None),
    approvals: Optional[List[UploadFile]] = File(None),
    options: str = Form("{}"),
):
    opts = _clean_options(options)
    uploads = {"start_listing": start_listing, "review_file": review_file, "post_listing": post_listing,
               "start_listing_pdf": start_listing_pdf, "post_listing_pdf": post_listing_pdf,
               "removed_accounts": removed_accounts, "extraction_steps": extraction_steps}
    for slot, up in uploads.items():
        if up is not None and up.filename:
            ext = os.path.splitext(up.filename)[1].lower()
            if ext not in SLOTS[slot][1]:
                raise HTTPException(400, f"{SLOTS[slot][3]}: {up.filename} is not one of "
                                         f"{', '.join(sorted(SLOTS[slot][1]))}")
    for up in approvals or []:
        if up.filename and os.path.splitext(up.filename)[1].lower() not in EMAIL_EXT | {".zip"}:
            raise HTTPException(400, f"Manager approval e-mails: {up.filename} must be .msg, .eml or .zip")

    run_id = f"{datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:6]}"
    d = os.path.join(RUNS_DIR, run_id)
    os.makedirs(d)
    try:
        inputs, names = {}, {}
        for slot, up in uploads.items():
            if up is not None and up.filename:
                inputs[SLOTS[slot][0]] = save_upload(up, os.path.join(d, "inputs", slot))
                names[slot] = up.filename
        edir = os.path.join(d, "inputs", "approvals")
        n_mail = 0
        for up in approvals or []:
            if not up.filename:
                continue
            p = save_upload(up, edir)
            if p.lower().endswith(".zip"):
                n_mail += extract_emails_zip(p, edir)
                os.remove(p)
            else:
                n_mail += 1
        inputs["emails_dir"] = edir if n_mail else None
        names["approvals"] = f"{n_mail} e-mail file(s)"
    except HTTPException:
        shutil.rmtree(d, ignore_errors=True)
        raise
    write_meta(d, {"run_id": run_id, "status": "queued", "created": datetime.now().isoformat(timespec="seconds"),
                   "options": opts, "inputs": inputs, "file_names": names,
                   "app_name": opts.get("app_name", E.CONFIG["app_name"])})
    open(os.path.join(d, "log.txt"), "w").close()
    append_log(d, f"Run {run_id} received {len(names)} evidence item(s).")
    EXECUTOR.submit(process, run_id)
    return {"run_id": run_id, "status": "queued"}


@app.get("/api/runs")
def list_runs(limit: int = 20):
    rows = []
    for name in sorted(os.listdir(RUNS_DIR), reverse=True):
        d = os.path.join(RUNS_DIR, name)
        if RUN_ID.match(name) and os.path.isfile(os.path.join(d, "meta.json")):
            m = read_meta(d)
            rows.append({k: m.get(k) for k in ("run_id", "created", "status", "app_name", "conclusion")})
        if len(rows) >= limit:
            break
    return rows


@app.get("/api/runs/{run_id}")
def get_run(run_id: str):
    d = run_dir(run_id)
    meta = read_meta(d)
    with open(os.path.join(d, "log.txt"), encoding="utf-8") as fh:
        log = [l.rstrip("\n") for l in fh.readlines()][-400:]
    payload = {k: meta.get(k) for k in ("run_id", "status", "created", "started", "finished", "error", "app_name",
                                         "file_names", "conclusion")}
    payload["options"] = {k: v for k, v in meta.get("options", {}).items()}
    payload["log"] = log
    if meta.get("status") == "done":
        res = load_results(run_id)
        payload["meta"] = {k: js(v) for k, v in res["meta"].items()}
        payload["summary"] = [{**s, "value": js(s["value"])} for s in res["summary"]]
        payload["tables"] = {k: table_payload(v) for k, v in res["tables"].items()}
    return payload


@app.get("/api/runs/{run_id}/workbook")
def download_workbook(run_id: str):
    d = run_dir(run_id)
    meta = read_meta(d)
    if meta.get("status") != "done":
        raise HTTPException(409, "The workbook is not ready yet")
    path = os.path.join(d, "output", meta["output"])
    return FileResponse(path, filename=meta["output"],
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


class ChatIn(BaseModel):
    message: str
    history: Optional[list] = None


@app.post("/api/runs/{run_id}/chat")
def chat(run_id: str, body: ChatIn):
    d = run_dir(run_id)
    q = (body.message or "").strip()
    if not q:
        raise HTTPException(400, "Type a question first")
    if len(q) > 2000:
        raise HTTPException(400, "Question is too long (2,000 characters max)")
    ai = E.AzureAI(None)
    if not ai.enabled:
        raise HTTPException(503, "Chat needs Azure OpenAI. Add the AZURE_OPENAI_* settings to backend/.env "
                                 "and restart the backend.")
    res = load_results(run_id)
    if run_id not in _db_cache:
        _db_cache[run_id] = RunDB(res)
    out = answer_question(ai, _db_cache[run_id], res, q, body.history)
    with _state_lock:
        with open(os.path.join(d, "chat_log.jsonl"), "a", encoding="utf-8") as fh:
            fh.write(json.dumps({"at": datetime.now().isoformat(timespec="seconds"), "question": q,
                                 "answer": out["answer"], "sql": [s["sql"] for s in out["steps"]]},
                                default=str) + "\n")
    return {"answer": out["answer"],
            "steps": [{**s, "rows": [[js(v) for v in r] for r in s["rows"]]} for s in out["steps"]]}


@app.delete("/api/runs/{run_id}")
def delete_run(run_id: str):
    d = run_dir(run_id)
    if read_meta(d).get("status") in ("queued", "running"):
        raise HTTPException(409, "This run is still processing")
    _results_cache.pop(run_id, None)
    _db_cache.pop(run_id, None)
    shutil.rmtree(d, ignore_errors=True)
    return {"deleted": run_id}


cleanup_old_runs()
