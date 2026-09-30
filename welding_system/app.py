import json
import os
import sqlite3
from pathlib import Path
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from fastapi import FastAPI, Request, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

app = FastAPI()

CURRENT_FILE_DIR = Path(__file__).resolve().parent

def find_project_file(possible_names: list) -> Path:
    for name in possible_names:
        if (CURRENT_FILE_DIR / name).exists():
            return CURRENT_FILE_DIR / name
        if (CURRENT_FILE_DIR.parent / name).exists():
            return CURRENT_FILE_DIR.parent / name
        for path in CURRENT_FILE_DIR.parent.rglob(name):
            return path
    return CURRENT_FILE_DIR / possible_names[0]

TEMPLATE_PATH = find_project_file(["template.docx", " template.docx"])
INDEX_PATH = find_project_file(["index.html", " index.html"])
DB_PATH = CURRENT_FILE_DIR / "welding_database.db"

UPLOADS_DIR = CURRENT_FILE_DIR / "uploads"
OUTPUT_DIR = CURRENT_FILE_DIR / "output"
os.makedirs(UPLOADS_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# เชื่อมต่อ SQLite
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
        return HTMLResponse(content="<h1>ไม่พบ index.html</h1>", status_code=404)
    with open(INDEX_PATH, "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read())

@app.post("/submit")
async def submit_inspection(request: Request):
    # รับข้อมูลทุกฟิลด์ที่ส่งมาจากแบบฟอร์ม HTML
    form = await request.form()
    form_data = dict(form)

    # แยกรูปภาพออกจากฟิลด์ข้อความ
    photo_fit_up = form.get("photo_fit_up")
    photo_visual = form.get("photo_visual")

    fit_up_path = None
    visual_path = None

    if isinstance(photo_fit_up, UploadFile) and photo_fit_up.filename:
        fit_up_path = UPLOADS_DIR / f"fitup_{photo_fit_up.filename}"
        with open(fit_up_path, "wb") as f:
            f.write(await photo_fit_up.read())

    if isinstance(photo_visual, UploadFile) and photo_visual.filename:
        visual_path = UPLOADS_DIR / f"visual_{photo_visual.filename}"
        with open(visual_path, "wb") as f:
            f.write(await photo_visual.read())

    # บันทึกลงฐานข้อมูล SQLite
    project_title = str(form_data.get("project_title", ""))
    joint_no = str(form_data.get("joint_no", "report"))
    welder_name = str(form_data.get("welder_name", ""))

    cursor.execute(
        "INSERT INTO inspections (project_title, joint_no, welder_name, data_json) VALUES (?, ?, ?, ?)",
        (project_title, joint_no, welder_name, json.dumps({k: str(v) for k, v in form_data.items() if not isinstance(v, UploadFile)}, ensure_ascii=False))
    )
    conn.commit()

    # ตรวจสอบไฟล์ template.docx
    actual_template = TEMPLATE_PATH
    if not actual_template.exists():
        docx_files = list(CURRENT_FILE_DIR.parent.rglob("*.docx"))
        if docx_files:
            actual_template = docx_files[0]
        else:
            return HTMLResponse(content="<h1>หาไฟล์ template.docx ไม่พบในระบบ</h1>", status_code=500)

    doc = DocxTemplate(str(actual_template))

    # เตรียม Context สำหรับแทนที่ตัวแปรใน Word
    context = {}
    for key, val in form_data.items():
        if not isinstance(val, UploadFile):
            context[key] = val

    # รองรับการแมปรูปภาพลงตำแหน่ง {{ img_1_1 }} หรือ {{ img_3_1 }}
    if fit_up_path and fit_up_path.exists():
        img_obj = InlineImage(doc, str(fit_up_path), width=Inches(2.5))
        context["img_1_1"] = img_obj
        context["photo_fit_up"] = img_obj
        context["cap_1_1"] = "Fit-up Inspection"
    else:
        context["img_1_1"] = ""
        context["cap_1_1"] = ""

    if visual_path and visual_path.exists():
        img_obj_v = InlineImage(doc, str(visual_path), width=Inches(2.5))
        context["img_3_1"] = img_obj_v
        context["photo_visual"] = img_obj_v
        context["cap_3_1"] = "Visual Inspection"
    else:
        context["img_3_1"] = ""
        context["cap_3_1"] = ""

    # ดักค่าเริ่มต้นให้ตัวแปรอื่นๆ ไม่ว่างเปล่า
    for i in range(1, 4):
        for j in range(1, 3):
            context.setdefault(f"img_{i}_{j}", "")
            context.setdefault(f"cap_{i}_{j}", "")

    # **จุดสำคัญที่สุด**: สั่ง Render แทนที่ตัวแปรทั้งหมดลงในเอกสารจริง
    doc.render(context)

    # บันทึกไฟล์ที่แทนค่าเรียบร้อยแล้ว
    output_filename = f"Inspection_{joint_no}.docx"
    output_filepath = OUTPUT_DIR / output_filename
    doc.save(str(output_filepath))

    return FileResponse(
        path=str(output_filepath),
        filename=output_filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
