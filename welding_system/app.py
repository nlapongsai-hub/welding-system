import json
import os
import sqlite3
from pathlib import Path
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

# ล็อกพาทให้อ้างอิงจากตำแหน่งของไฟล์ app.py เสมอ
BASE_DIR = Path(__file__).resolve().parent
UPLOADS_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"
DB_PATH = BASE_DIR / "welding_database.db"
TEMPLATE_PATH = BASE_DIR / "template.docx"
INDEX_PATH = BASE_DIR / "index.html"

app = FastAPI()

os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# เชื่อมต่อฐานข้อมูล SQLite
conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
cursor = conn.cursor()
cursor.execute("""
CREATE TABLE IF NOT EXISTS inspections (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_title TEXT,
    joint_no TEXT,
    welder_name TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    data_json TEXT
)
""")
conn.commit()

@app.get("/", response_class=HTMLResponse)
async def read_root():
    if not INDEX_PATH.exists():
        return HTMLResponse(content="<h1>ไม่พบไฟล์ index.html ในระบบ</h1>", status_code=404)
    with open(INDEX_PATH, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.post("/submit")
async def submit_inspection(
    project_title: str = Form(""),
    joint_no: str = Form(""),
    welder_name: str = Form(""),
    inspector: str = Form(""),
    date: str = Form(""),
    welding_process: str = Form(""),
    wps_no: str = Form(""),
    drawing_no: str = Form(""),
    fit_up_result: str = Form("PASS"),
    visual_result: str = Form("PASS"),
    remarks: str = Form(""),
    photo_fit_up: UploadFile = File(None),
    photo_visual: UploadFile = File(None),
):
    fit_up_path = None
    visual_path = None

    if photo_fit_up and photo_fit_up.filename:
        fit_up_path = UPLOADS_DIR / f"fitup_{photo_fit_up.filename}"
        with open(fit_up_path, "wb") as f:
            f.write(await photo_fit_up.read())

    if photo_visual and photo_visual.filename:
        visual_path = UPLOADS_DIR / f"visual_{photo_visual.filename}"
        with open(visual_path, "wb") as f:
            f.write(await photo_visual.read())

    # บันทึกลงฐานข้อมูล SQLite
    record_data = {
        "project_title": project_title,
        "joint_no": joint_no,
        "welder_name": welder_name,
        "inspector": inspector,
        "date": date,
        "welding_process": welding_process,
        "wps_no": wps_no,
        "drawing_no": drawing_no,
        "fit_up_result": fit_up_result,
        "visual_result": visual_result,
        "remarks": remarks,
    }
    cursor.execute(
        "INSERT INTO inspections (project_title, joint_no, welder_name, data_json) VALUES (?, ?, ?, ?)",
        (project_title, joint_no, welder_name, json.dumps(record_data, ensure_ascii=False))
    )
    conn.commit()

    # สร้างรายงาน Word จากเทมเพลต
    if not TEMPLATE_PATH.exists():
        return HTMLResponse(content="<h1>ไม่พบไฟล์ template.docx ในระบบ</h1>", status_code=500)

    doc = DocxTemplate(str(TEMPLATE_PATH))
    context = dict(record_data)

    if fit_up_path and fit_up_path.exists():
        context["photo_fit_up"] = InlineImage(doc, str(fit_up_path), width=Inches(2.5))
    else:
        context["photo_fit_up"] = "ไม่มีรูปภาพ"

    if visual_path and visual_path.exists():
        context["photo_visual"] = InlineImage(doc, str(visual_path), width=Inches(2.5))
    else:
        context["photo_visual"] = "ไม่มีรูปภาพ"

    output_filename = f"Inspection_{joint_no or 'report'}.docx"
    output_filepath = OUTPUT_DIR / output_filename
    doc.save(str(output_filepath))

    return FileResponse(
        path=str(output_filepath),
        filename=output_filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
