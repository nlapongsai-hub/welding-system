import io
import re
import uuid
from pathlib import Path
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Inches
from PIL import Image, ImageOps, UnidentifiedImageError

BASE_DIR = Path(__file__).resolve().parent
PHOTO_SLOTS = [f"img_{i}_{j}" for i in range(1, 4) for j in range(1, 3)]
GROUPS = ["ga", "ro", "rf", "cl", "ph", "rd", "vi", "dim", "suf", "pt", "p1", "p2", "p3", "p4", "p5"]

def project_file(name):
    return BASE_DIR / name

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
                width = min(2.5, 1.55 * image.width / image.height)
                context[slot] = InlineImage(doc, stream, width=Inches(width))
        except (UnidentifiedImageError, OSError, Image.DecompressionBombError) as exc:
            raise ValueError(f"รูป {slot} เปิดไม่ได้ กรุณาใช้ไฟล์ JPG หรือ PNG") from exc
    for variable in doc.get_undeclared_template_variables():
        context.setdefault(variable, "")
    doc.render(context, autoescape=True)
    joint = re.sub(r"[^\w-]+", "_", fields.get("joint_no", "report"), flags=re.UNICODE).strip("_")[:60] or "report"
    filename = f"Inspection_{joint}_{uuid.uuid4().hex[:8]}.docx"
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue(), filename, database_data
