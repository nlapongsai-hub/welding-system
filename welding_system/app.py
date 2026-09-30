import json
import os
import re
import sqlite3
import traceback
import uuid
from pathlib import Path

from fastapi import FastAPI, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches


# =========================================================
# APP
# =========================================================

app = FastAPI(
    title="Welding Inspection System",
    version="2.0.0"
)

CURRENT_FILE_DIR = Path(__file__).resolve().parent


# =========================================================
# ค้นหาไฟล์ใน Project
# =========================================================

def find_project_file(possible_names):
    """
    ค้นหาไฟล์จาก:
    1. โฟลเดอร์เดียวกับ app.py
    2. โฟลเดอร์ parent
    3. subfolder ภายใน project
    """

    for name in possible_names:
        clean_name = name.strip()

        # โฟลเดอร์เดียวกับ app.py
        path = CURRENT_FILE_DIR / clean_name
        if path.exists():
            return path

        # parent directory
        path = CURRENT_FILE_DIR.parent / clean_name
        if path.exists():
            return path

    # ค้นหาแบบ recursive
    for name in possible_names:
        clean_name = name.strip()

        try:
            for path in CURRENT_FILE_DIR.parent.rglob(clean_name):
                if path.is_file():
                    return path
        except Exception:
            pass

    return CURRENT_FILE_DIR / possible_names[0].strip()


# =========================================================
# PATH
# =========================================================

TEMPLATE_PATH = find_project_file([
    "template.docx",
    "Template.docx"
])

INDEX_PATH = find_project_file([
    "index.html",
    "Index.html"
])

DB_PATH = CURRENT_FILE_DIR / "welding_database.db"

UPLOADS_DIR = CURRENT_FILE_DIR / "uploads"
OUTPUT_DIR = CURRENT_FILE_DIR / "output"

UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# =========================================================
# DATABASE
# =========================================================

def init_database():

    with sqlite3.connect(str(DB_PATH)) as conn:

        conn.execute("""
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


init_database()


# =========================================================
# Utility
# =========================================================

def safe_filename(filename):
    """
    ป้องกันชื่อไฟล์แปลก ๆ
    """

    if not filename:
        return "file"

    filename = Path(filename).name

    filename = re.sub(
        r'[^A-Za-z0-9ก-๙._-]',
        '_',
        filename
    )

    return filename or "file"


def safe_text(value):
    """
    แปลงค่าจาก Form เป็น String อย่างปลอดภัย
    """

    if value is None:
        return ""

    return str(value).strip()


# =========================================================
# HOME
# =========================================================

@app.get("/", response_class=HTMLResponse)
async def read_root():

    if not INDEX_PATH.exists():

        return HTMLResponse(
            content=f"""
            <html>
            <head>
                <meta charset="utf-8">
            </head>

            <body>
                <h1>ไม่พบ index.html</h1>

                <p>ตำแหน่งที่ระบบค้นหา:</p>

                <pre>{INDEX_PATH}</pre>

            </body>
            </html>
            """,
            status_code=404
        )

    try:

        with open(
            INDEX_PATH,
            "r",
            encoding="utf-8"
        ) as f:

            html = f.read()

        return HTMLResponse(content=html)

    except Exception as e:

        return HTMLResponse(
            content=f"<h1>อ่าน index.html ไม่สำเร็จ</h1><pre>{e}</pre>",
            status_code=500
        )


# =========================================================
# HEALTH CHECK
# =========================================================

@app.get("/health")
async def health():

    return {
        "status": "ok",
        "template_exists": TEMPLATE_PATH.exists(),
        "index_exists": INDEX_PATH.exists()
    }


# =========================================================
# SUBMIT
# =========================================================

@app.post("/submit")
async def submit_inspection(request: Request):

    saved_images = {}
    context = {}

    CHECKED = "☑"
    UNCHECKED = "☐"

    try:

        # =================================================
        # 1. อ่าน Form
        # =================================================

        form = await request.form()


        # =================================================
        # 2. แยกข้อความ / รูป
        # =================================================

        for key, value in form.multi_items():

            # -------------------------------
            # Upload File
            # -------------------------------

            if isinstance(value, UploadFile):

                if not value.filename:
                    continue

                content = await value.read()

                if not content:
                    continue

                original_name = safe_filename(value.filename)

                unique_name = (
                    f"{uuid.uuid4().hex}_"
                    f"{safe_filename(key)}_"
                    f"{original_name}"
                )

                image_path = UPLOADS_DIR / unique_name

                with open(image_path, "wb") as f:
                    f.write(content)

                saved_images[key] = image_path

            # -------------------------------
            # Normal Form Value
            # -------------------------------

            else:

                context[key] = safe_text(value)


        # =================================================
        # 3. Checkbox PASS / FAIL / N/A
        # =================================================

        check_groups = [
            "ga",
            "ro",
            "rf",
            "cl",
            "ph",
            "vi",
            "dim",
            "suf",
            "pt",
            "p1",
            "p2",
            "p3",
            "p4",
            "p5"
        ]

        for group in check_groups:

            choice = (
                form.get(f"result_{group}")
                or form.get(f"{group}_result")
                or form.get(group)
                or ""
            )

            choice_str = str(choice).strip().upper()

            # รองรับหลายรูปแบบ
            is_na = choice_str in [
                "NA",
                "N/A",
                "N.A.",
                "NOT APPLICABLE"
            ]

            is_pass = choice_str == "PASS"
            is_fail = choice_str == "FAIL"

            context[f"c_{group}_na"] = (
                CHECKED if is_na else UNCHECKED
            )

            context[f"c_{group}_pass"] = (
                CHECKED if is_pass else UNCHECKED
            )

            context[f"c_{group}_fail"] = (
                CHECKED if is_fail else UNCHECKED
            )


        # =================================================
        # 4. ข้อมูลหลัก
        # =================================================

        project_title = safe_text(
            context.get("project_title")
        )

        joint_no = safe_text(
            context.get("joint_no")
        )

        welder_name = safe_text(
            context.get("welder_name")
        )

        if not joint_no:
            joint_no = "report"


        # =================================================
        # 5. ตรวจสอบ Template
        # =================================================

        actual_template = TEMPLATE_PATH

        if not actual_template.exists():

            docx_files = []

            try:

                docx_files = list(
                    CURRENT_FILE_DIR.parent.rglob(
                        "*.docx"
                    )
                )

                # ไม่เอาไฟล์ใน output
                docx_files = [
                    p for p in docx_files
                    if "output" not in p.parts
                ]

            except Exception:
                pass

            if docx_files:

                actual_template = docx_files[0]

            else:

                return HTMLResponse(
                    content="""
                    <html>
                    <head>
                        <meta charset="utf-8">
                    </head>

                    <body>

                    <h2>ไม่พบ template.docx</h2>

                    <p>
                    กรุณาตรวจสอบว่าใน GitHub
                    มีไฟล์ชื่อ
                    <b>template.docx</b>
                    อยู่ใน Project
                    </p>

                    </body>
                    </html>
                    """,
                    status_code=500
                )


        # =================================================
        # 6. เปิด Word Template
        # =================================================

        try:

            doc = DocxTemplate(
                str(actual_template)
            )

        except Exception:

            return HTMLResponse(
                content=f"""
                <h2>เปิด template.docx ไม่สำเร็จ</h2>

                <pre>
{traceback.format_exc()}
                </pre>
                """,
                status_code=500
            )


        # =================================================
        # 7. รูปภาพ
        # =================================================

        photo_slots = [
            "img_1_1",
            "img_1_2",
            "img_2_1",
            "img_2_2",
            "img_3_1",
            "img_3_2"
        ]

        for slot in photo_slots:

            image_path = saved_images.get(slot)

            if (
                image_path
                and image_path.exists()
            ):

                try:

                    context[slot] = InlineImage(
                        doc,
                        str(image_path),
                        width=Inches(2.5)
                    )

                except Exception:

                    context[slot] = ""

            else:

                context[slot] = ""


            # caption
            caption_name = slot.replace(
                "img",
                "cap",
                1
            )

            context.setdefault(
                caption_name,
                ""
            )


        # =================================================
        # 8. ตรวจตัวแปรใน Word
        # =================================================

        try:

            template_variables = (
                doc.get_undeclared_template_variables()
            )

            # ตัวแปรที่ Word มี
            # แต่ HTML ไม่ได้ส่งมา
            # ให้เป็นช่องว่างอัตโนมัติ

            for variable in template_variables:

                if variable not in context:
                    context[variable] = ""

        except Exception as variable_error:

            print(
                "WARNING: ไม่สามารถอ่าน "
                "template variables ได้:",
                variable_error
            )


        # =================================================
        # 9. Render Word
        # =================================================

        try:

            doc.render(context)

        except Exception as render_error:

            error_trace = traceback.format_exc()

            print("========== DOCXTPL ERROR ==========")
            print(error_trace)
            print("===================================")

            # แสดงเฉพาะข้อมูลที่ช่วย Debug
            context_keys = sorted(
                list(context.keys())
            )

            return HTMLResponse(
                content=f"""
<!DOCTYPE html>

<html lang="th">

<head>

<meta charset="UTF-8">

<meta name="viewport"
content="width=device-width, initial-scale=1">

<title>Word Template Error</title>

<style>

body {{
    font-family: Arial, sans-serif;
    padding: 30px;
    background: #f5f5f5;
}}

.box {{
    max-width: 1000px;
    margin: auto;
    background: white;
    padding: 30px;
    border-radius: 12px;
}}

pre {{
    white-space: pre-wrap;
    word-break: break-word;
    background: #eee;
    padding: 15px;
    border-radius: 8px;
}}

.error {{
    color: #b00020;
}}

</style>

</head>

<body>

<div class="box">

<h2 class="error">
เกิดข้อผิดพลาดตอนสร้าง Word
</h2>

<h3>Error</h3>

<pre>{str(render_error)}</pre>


<h3>Template ที่ใช้งาน</h3>

<pre>{actual_template}</pre>


<h3>ตัวแปรที่ Backend มี</h3>

<pre>{context_keys}</pre>


<h3>Traceback</h3>

<pre>{error_trace}</pre>


<hr>

<p>
ถ้า Error ระบุว่า
TemplateSyntaxError
ให้ตรวจสอบเครื่องหมาย
{{{{ }}}}
ใน template.docx
</p>

</div>

</body>

</html>
""",
                status_code=500
            )


        # =================================================
        # 10. บันทึก Database
        # =================================================

        db_dict = {}

        for key, value in context.items():

            # ไม่บันทึก InlineImage
            if isinstance(
                value,
                InlineImage
            ):
                continue

            # ไม่บันทึก Path
            if isinstance(
                value,
                Path
            ):
                continue

            db_dict[key] = str(value)


        try:

            with sqlite3.connect(
                str(DB_PATH)
            ) as db_conn:

                db_conn.execute(
                    """
                    INSERT INTO inspections
                    (
                        project_title,
                        joint_no,
                        welder_name,
                        data_json
                    )

                    VALUES (?, ?, ?, ?)
                    """,

                    (
                        project_title,
                        joint_no,
                        welder_name,
                        json.dumps(
                            db_dict,
                            ensure_ascii=False
                        )
                    )
                )

                db_conn.commit()

        except Exception as db_error:

            # ไม่ให้ database error
            # ทำให้ Word สร้างไม่ได้

            print(
                "DATABASE ERROR:",
                db_error
            )


        # =================================================
        # 11. สร้างชื่อไฟล์
        # =================================================

        safe_joint = re.sub(
            r'[^A-Za-z0-9ก-๙_-]',
            '_',
            joint_no
        )

        safe_joint = (
            safe_joint.strip("_")
            or "report"
        )

        unique_id = uuid.uuid4().hex[:8]

        output_filename = (
            f"Inspection_"
            f"{safe_joint}_"
            f"{unique_id}.docx"
        )

        output_filepath = (
            OUTPUT_DIR /
            output_filename
        )


        # =================================================
        # 12. SAVE WORD
        # =================================================

        try:

            doc.save(
                str(output_filepath)
            )

        except Exception:

            return HTMLResponse(
                content=f"""
                <h2>
                ไม่สามารถบันทึก Word ได้
                </h2>

                <pre>
{traceback.format_exc()}
                </pre>
                """,
                status_code=500
            )


        # =================================================
        # 13. ส่ง Word กลับ
        # =================================================

        return FileResponse(

            path=str(output_filepath),

            filename=output_filename,

            media_type=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            )
        )


    # =====================================================
    # GLOBAL ERROR
    # =====================================================

    except Exception:

        error_detail = (
            traceback.format_exc()
        )

        print(
            "========== SYSTEM ERROR =========="
        )

        print(error_detail)

        print(
            "=================================="
        )

        return HTMLResponse(
            content=f"""
<!DOCTYPE html>

<html lang="th">

<head>

<meta charset="UTF-8">

<title>System Error</title>

<style>

body {{
    font-family: Arial, sans-serif;
    padding: 30px;
}}

pre {{
    white-space: pre-wrap;
    word-break: break-word;
    background: #eee;
    padding: 20px;
}}

</style>

</head>

<body>

<h2>
เกิดข้อผิดพลาดในการประมวลผล
</h2>

<pre>
{error_detail}
</pre>

</body>

</html>
""",
            status_code=500
        )
