import json
import os
import sqlite3
import traceback
from pathlib import Path
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from fastapi import FastAPI, Request
from fastapi.responses import FileResponse, HTMLResponse
from starlette.datastructures import UploadFile

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
    try:
        form = await request.form()
        
        saved_images = {}
        context = {}

        CHECKED = "☑"
        UNCHECKED = "☐"

        # 1. จัดการแยกไฟล์รูปและข้อความ
        for key, val in form.items():
            if isinstance(val, UploadFile):
                if val.filename and len(val.filename.strip()) > 0:
                    content = await val.read()
                    if len(content) > 0:
                        img_path = UPLOADS_DIR / f"{key}_{val.filename}"
                        with open(img_path, "wb") as f:
                            f.write(content)
                        saved_images[key] = img_path
                    else:
                        context[key] = " "
                else:
                    context[key] = " "
            else:
                context[key] = str(val)

        # 2. จัดการตัวแปรติ๊กถูก Pass / Fail / N/A
        check_groups = ["ga", "ro", "rf", "cl", "ph", "vi", "dim", "suf", "pt", "p1", "p2", "p3", "p4", "p5"]
        for group in check_groups:
            choice = form.get(f"result_{group}") or form.get(f"{group}_result") or form.get(group)
            choice_str = str(choice).upper() if choice else ""

            context[f"c_{group}_na"] = CHECKED if ("NA" in choice_str or "N/A" in choice_str) else UNCHECKED
            context[f"c_{group}_pass"] = CHECKED if "PASS" in choice_str else UNCHECKED
            context[f"c_{group}_fail"] = CHECKED if "FAIL" in choice_str else UNCHECKED

        # 3. บันทึกข้อมูลลงฐานข้อมูล SQLite
        project_title = str(context.get("project_title", ""))
        joint_no = str(context.get("joint_no", "report"))
        welder_name = str(context.get("welder_name", ""))

        db_dict = {k: v for k, v in context.items() if not isinstance(v, (Path, InlineImage))}
        cursor.execute(
            "INSERT INTO inspections (project_title, joint_no, welder_name, data_json) VALUES (?, ?, ?, ?)",
            (project_title, joint_no, welder_name, json.dumps(db_dict, ensure_ascii=False))
        )
        conn.commit()

        # 4. ตรวจสอบไฟล์ template.docx
        actual_template = TEMPLATE_PATH
        if not actual_template.exists():
            docx_files = list(CURRENT_FILE_DIR.parent.rglob("*.docx"))
            if docx_files:
                actual_template = docx_files[0]
            else:
                return HTMLResponse(content="<h1>หาไฟล์ template.docx ไม่พบในระบบ</h1>", status_code=500)

        doc = DocxTemplate(str(actual_template))

        # 5. แมปรูปภาพลงตัวแปร Word (ใช้ช่องว่างแทน None/Empty เพื่อป้องกัน docxtpl พัง)
        photo_slots = ["img_1_1", "img_1_2", "img_2_1", "img_2_2", "img_3_1", "img_3_2"]
        for slot in photo_slots:
            if slot in saved_images and saved_images[slot].exists():
                try:
                    context[slot] = InlineImage(doc, str(saved_images[slot]), width=Inches(2.5))
                except Exception:
                    context[slot] = " "
            else:
                context[slot] = " "
            context.setdefault(slot.replace("img", "cap"), " ")

        # 6. ประมวลผลเอกสาร Word
        doc.render(context)

        # 7. บันทึกและส่งไฟล์กลับ
        safe_joint = "".join(c for c in joint_no if c.isalnum() or c in ("-", "_")).strip() or "report"
        output_filename = f"Inspection_{safe_joint}.docx"
        output_filepath = OUTPUT_DIR / output_filename
        doc.save(str(output_filepath))

        return FileResponse(
            path=str(output_filepath),
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        error_detail = traceback.format_exc()
        return HTMLResponse(
            content=f"<h3>เกิดข้อผิดพลาดในการประมวลผล:</h3><pre>{error_detail}</pre>",
            status_code=500
        )
