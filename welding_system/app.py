import html
import io
import json
import logging
import os
import re
import sqlite3
import tempfile
import uuid
from pathlib import Path

from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, HTMLResponse
from jinja2 import TemplateError
from PIL import Image, ImageOps, UnidentifiedImageError
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool
from starlette.datastructures import UploadFile

app = FastAPI(title="BSC Welding Inspection")
logger = logging.getLogger("welding")
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("WELDING_DATA_DIR", str(BASE_DIR)))
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "welding_database.db"
PHOTO_SLOTS = [f"img_{i}_{j}" for i in range(1, 4) for j in range(1, 3)]
GROUPS = ["ga", "ro", "rf", "cl", "ph", "rd", "vi", "dim", "suf", "pt", "p1", "p2", "p3", "p4", "p5"]
MAX_IMAGE_BYTES = 10 * 1024 * 1024


def project_file(name):
    for root in (BASE_DIR, BASE_DIR.parent):
        path = root / name
        if path.is_file():
            return path
    return BASE_DIR / name


def init_database():
    with sqlite3.connect(DB_PATH, timeout=30) as conn:
        conn.execute("""CREATE TABLE IF NOT EXISTS inspections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            project_title TEXT, joint_no TEXT, welder_name TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP, data_json TEXT
        )""")


init_database()


def error_page(message, status=500):
    return HTMLResponse(
        '<!doctype html><html lang="th"><meta charset="utf-8">'
        '<body><h2>ไม่สามารถสร้างรายงานได้</h2><p>'
        + html.escape(message) + '</p><a href="/">กลับไปกรอกข้อมูล</a></body></html>',
        status_code=status,
    )


def build_report(fields, images):
    template = project_file("template.docx")
    if not template.is_file():
        raise FileNotFoundError("ไม่พบ template.docx กรุณาวางไฟล์ไว้ในโฟลเดอร์เดียวกับ app.py")
    doc = DocxTemplate(str(template))
    context = dict(fields)
    for group in GROUPS:
        choice = next((fields[k] for k in (
            f"{group}_res", f"result_{group}", f"{group}_result", group
        ) if fields.get(k)), "").strip().upper()
        if choice in ("NA", "N.A.", "NOT APPLICABLE"):
            choice = "N/A"
        if choice not in ("", "PASS", "FAIL", "N/A"):
            raise ValueError(f"ผลตรวจ {group} ไม่ถูกต้อง")
        for suffix, value in (("na", "N/A"), ("pass", "PASS"), ("fail", "FAIL")):
            context[f"c_{group}_{suffix}"] = "☑" if choice == value else "☐"
    # เก็บข้อมูลข้อความก่อนเพิ่ม InlineImage
    database_data = dict(context)
    streams = []
    for slot in PHOTO_SLOTS:
        context[slot] = ""
        context.setdefault(slot.replace("img", "cap", 1), "")
        if slot not in images:
            continue
        try:
            with Image.open(io.BytesIO(images[slot])) as original:
                original.load()
                image = ImageOps.exif_transpose(original).convert("RGB")
                image.thumbnail((1600, 1600))
                stream = io.BytesIO()
                image.save(stream, format="JPEG", quality=90)
                stream.seek(0)
                streams.append(stream)
                # จำกัดทั้งความกว้างและสูง โดยรักษาสัดส่วนของภาพ
                width = min(2.5, 1.55 * image.width / image.height)
                context[slot] = InlineImage(doc, stream, width=Inches(width))
        except (UnidentifiedImageError, OSError, Image.DecompressionBombError) as exc:
            raise ValueError(f"รูป {slot} เปิดไม่ได้ กรุณาใช้ไฟล์ JPG หรือ PNG") from exc
    for variable in doc.get_undeclared_template_variables():
        context.setdefault(variable, "")
    doc.render(context, autoescape=True)
    joint = re.sub(r"[^\w-]+", "_", fields.get("joint_no", "report"), flags=re.UNICODE).strip("_")[:60] or "report"
    filename = f"Inspection_{joint}_{uuid.uuid4().hex[:8]}.docx"
    fd, path = tempfile.mkstemp(suffix=".docx")
    os.close(fd)
    try:
        doc.save(path)
        with sqlite3.connect(DB_PATH, timeout=30) as conn:
            conn.execute(
                "INSERT INTO inspections (project_title, joint_no, welder_name, data_json) VALUES (?, ?, ?, ?)",
                (fields.get("project_title", ""), fields.get("joint_no", ""),
                 fields.get("welder_name", ""), json.dumps(database_data, ensure_ascii=False)),
            )
    except Exception:
        Path(path).unlink(missing_ok=True)
        raise
    return path, filename


@app.get("/", response_class=HTMLResponse)
def home():
    index = project_file("index.html")
    if not index.is_file():
        return error_page("ไม่พบ index.html", 404)
    return HTMLResponse(index.read_text(encoding="utf-8"))


@app.get("/health")
def health():
    return {"status": "ok", "template_exists": project_file("template.docx").is_file(),
            "index_exists": project_file("index.html").is_file()}


@app.post("/submit")
async def submit_inspection(request: Request):
    try:
        fields, images = {}, {}
        async with request.form() as form:
            for key, value in form.multi_items():
                # request.form() คืน UploadFile ของ Starlette
                if isinstance(value, UploadFile):
                    if key not in PHOTO_SLOTS or not value.filename:
                        continue
                    content = await value.read(MAX_IMAGE_BYTES + 1)
                    if len(content) > MAX_IMAGE_BYTES:
                        return error_page("รูปแต่ละรูปต้องมีขนาดไม่เกิน 10 MB", 413)
                    if content:
                        images[key] = content
                else:
                    fields[key] = str(value).strip()
        if not fields.get("project_title") or not fields.get("joint_no"):
            return error_page("กรุณากรอกโครงการและหมายเลขรอยต่อ", 400)
        path, filename = await run_in_threadpool(build_report, fields, images)
        return FileResponse(
            path, filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            background=BackgroundTask(Path(path).unlink, missing_ok=True),
        )
    except ValueError as exc:
        return error_page(str(exc), 400)
    except FileNotFoundError as exc:
        return error_page(str(exc))
    except TemplateError:
        logger.exception("Word template error")
        return error_page("ตัวแปรใน template.docx ไม่ถูกต้อง กรุณาใช้แม่แบบที่มาพร้อมชุดนี้")
    except Exception:
        logger.exception("Inspection generation failed")
        return error_page("เกิดข้อผิดพลาด กรุณาดูรายละเอียดใน Logs ของ Render")

