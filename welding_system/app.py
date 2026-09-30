import json
import os
import sqlite3
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import FileResponse, HTMLResponse

app = FastAPI()

os.makedirs("uploads", exist_ok=True)
os.makedirs("output", exist_ok=True)

# ฐานข้อมูล SQLite
conn = sqlite3.connect("welding_database.db", check_same_thread=False)
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
def home():
    with open("index.html", "r", encoding="utf-8") as f:
        return f.read()


@app.post("/submit")
async def submit_form(
    project_title: str = Form(""),
    joint_no: str = Form(""),
    welder_name: str = Form(""),
    welder_id: str = Form(""),
    process: str = Form(""),
    date: str = Form(""),
    # Section 1
    ga_res: str = Form("Pass"),
    r_ga: str = Form(""),
    ro_res: str = Form("Pass"),
    r_ro: str = Form(""),
    rf_res: str = Form("Pass"),
    r_rf: str = Form(""),
    cl_res: str = Form("Pass"),
    r_cl: str = Form(""),
    ph_res: str = Form("Pass"),
    rd_res: str = Form("Pass"),
    # Section 2
    p1_type: str = Form(""),
    p1_amp: str = Form(""),
    p1_volt: str = Form(""),
    p1_spd: str = Form(""),
    p1_res: str = Form("Pass"),
    p2_type: str = Form(""),
    p2_amp: str = Form(""),
    p2_volt: str = Form(""),
    p2_spd: str = Form(""),
    p2_res: str = Form("Pass"),
    p3_type: str = Form(""),
    p3_amp: str = Form(""),
    p3_volt: str = Form(""),
    p3_spd: str = Form(""),
    p3_res: str = Form("Pass"),
    p4_type: str = Form(""),
    p4_amp: str = Form(""),
    p4_volt: str = Form(""),
    p4_spd: str = Form(""),
    p4_res: str = Form("Pass"),
    p5_type: str = Form(""),
    p5_amp: str = Form(""),
    p5_volt: str = Form(""),
    p5_spd: str = Form(""),
    p5_res: str = Form("Pass"),
    # Section 3
    vi_res: str = Form("Pass"),
    r_vi: str = Form(""),
    dim_res: str = Form("Pass"),
    r_dim: str = Form(""),
    suf_res: str = Form("Pass"),
    r_suf: str = Form(""),
    pt_res: str = Form("Pass"),
    r_pt: str = Form(""),
    # Photographs & Captions
    img_1_1: UploadFile = File(None),
    cap_1_1: str = Form(""),
    img_1_2: UploadFile = File(None),
    cap_1_2: str = Form(""),
    img_2_1: UploadFile = File(None),
    cap_2_1: str = Form(""),
    img_2_2: UploadFile = File(None),
    cap_2_2: str = Form(""),
    img_3_1: UploadFile = File(None),
    cap_3_1: str = Form(""),
    img_3_2: UploadFile = File(None),
    cap_3_2: str = Form(""),
):
    uploaded_files = {
        "1_1": img_1_1,
        "1_2": img_1_2,
        "2_1": img_2_1,
        "2_2": img_2_2,
        "3_1": img_3_1,
        "3_2": img_3_2,
    }
    saved_img_paths = {}

    for key, file_obj in uploaded_files.items():
        if file_obj and file_obj.filename:
            file_path = f"uploads/{joint_no}_{key}_{file_obj.filename}"
            with open(file_path, "wb") as f:
                f.write(await file_obj.read())
            saved_img_paths[key] = file_path

    def chk(expected, actual):
        return "☑" if expected == actual else "☐"

    doc = DocxTemplate("template.docx")

    def bind_image(key):
        if key in saved_img_paths and os.path.exists(saved_img_paths[key]):
            return InlineImage(doc, saved_img_paths[key], width=Inches(2.8))
        return ""

    context = {
        "project_title": project_title,
        "joint_no": joint_no,
        "welder_name": welder_name,
        "welder_id": welder_id,
        "process": process,
        "date": date,
        # Section 1
        "c_ga_pass": chk("Pass", ga_res),
        "c_ga_fail": chk("Fail", ga_res),
        "c_ga_na": chk("N/A", ga_res),
        "r_ga": r_ga,
        "c_ro_pass": chk("Pass", ro_res),
        "c_ro_fail": chk("Fail", ro_res),
        "c_ro_na": chk("N/A", ro_res),
        "r_ro": r_ro,
        "c_rf_pass": chk("Pass", rf_res),
        "c_rf_fail": chk("Fail", rf_res),
        "c_rf_na": chk("N/A", rf_res),
        "r_rf": r_rf,
        "c_cl_pass": chk("Pass", cl_res),
        "c_cl_fail": chk("Fail", cl_res),
        "c_cl_na": chk("N/A", cl_res),
        "r_cl": r_cl,
        "c_ph_pass": chk("Pass", ph_res),
        "c_ph_fail": chk("Fail", ph_res),
        "c_ph_na": chk("N/A", ph_res),
        "c_rd_pass": chk("Pass", rd_res),
        "c_rd_fail": chk("Fail", rd_res),
        "c_rd_na": chk("N/A", rd_res),
        # Section 2
        "p1_type": p1_type,
        "p1_amp": p1_amp,
        "p1_volt": p1_volt,
        "p1_spd": p1_spd,
        "c_p1_pass": chk("Pass", p1_res),
        "c_p1_fail": chk("Fail", p1_res),
        "p2_type": p2_type,
        "p2_amp": p2_amp,
        "p2_volt": p2_volt,
        "p2_spd": p2_spd,
        "c_p2_pass": chk("Pass", p2_res),
        "c_p2_fail": chk("Fail", p2_res),
        "p3_type": p3_type,
        "p3_amp": p3_amp,
        "p3_volt": p3_volt,
        "p3_spd": p3_spd,
        "c_p3_pass": chk("Pass", p3_res),
        "c_p3_fail": chk("Fail", p3_res),
        "p4_type": p4_type,
        "p4_amp": p4_amp,
        "p4_volt": p4_volt,
        "p4_spd": p4_spd,
        "c_p4_pass": chk("Pass", p4_res),
        "c_p4_fail": chk("Fail", p4_res),
        "p5_type": p5_type,
        "p5_amp": p5_amp,
        "p5_volt": p5_volt,
        "p5_spd": p5_spd,
        "c_p5_pass": chk("Pass", p5_res),
        "c_p5_fail": chk("Fail", p5_res),
        # Section 3
        "c_vi_pass": chk("Pass", vi_res),
        "c_vi_fail": chk("Fail", vi_res),
        "c_vi_na": chk("N/A", vi_res),
        "r_vi": r_vi,
        "c_dim_pass": chk("Pass", dim_res),
        "c_dim_fail": chk("Fail", dim_res),
        "c_dim_na": chk("N/A", dim_res),
        "r_dim": r_dim,
        "c_suf_pass": chk("Pass", suf_res),
        "c_suf_fail": chk("Fail", suf_res),
        "c_suf_na": chk("N/A", suf_res),
        "r_suf": r_suf,
        "c_pt_pass": chk("Pass", pt_res),
        "c_pt_fail": chk("Fail", pt_res),
        "c_pt_na": chk("N/A", pt_res),
        "r_pt": r_pt,
        # Images & Captions
        "img_1_1": bind_image("1_1"),
        "cap_1_1": cap_1_1,
        "img_1_2": bind_image("1_2"),
        "cap_1_2": cap_1_2,
        "img_2_1": bind_image("2_1"),
        "cap_2_1": cap_2_1,
        "img_2_2": bind_image("2_2"),
        "cap_2_2": cap_2_2,
        "img_3_1": bind_image("3_1"),
        "cap_3_1": cap_3_1,
        "img_3_2": bind_image("3_2"),
        "cap_3_2": cap_3_2,
    }

    output_filename = f"Report_{joint_no if joint_no else 'welding'}.docx"
    output_path = f"output/{output_filename}"
    doc.render(context)
    doc.save(output_path)

    raw_data = {
        k: v for k, v in context.items() if not str(k).startswith("img_")
    }
    cursor.execute(
        """
        INSERT INTO inspections (project_title, joint_no, welder_name, data_json)
        VALUES (?, ?, ?, ?)
    """,
        (
            project_title,
            joint_no,
            welder_name,
            json.dumps(raw_data, ensure_ascii=False),
        ),
    )
    conn.commit()

    return FileResponse(
        output_path,
        filename=output_filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

