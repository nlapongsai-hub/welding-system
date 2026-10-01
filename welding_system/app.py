import hashlib
import hmac
import html
import json
import logging
import os
import secrets
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import quote, urlencode
from uuid import UUID, uuid4
from zoneinfo import ZoneInfo

import httpx
from fastapi import Depends, FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.security import HTTPBasic, HTTPBasicCredentials
from jinja2 import TemplateError
from starlette.concurrency import run_in_threadpool
from starlette.datastructures import UploadFile

from report_builder import build_report, PHOTO_SLOTS

app = FastAPI(title="BSC Welding Report Archive")
logger = logging.getLogger("welding")
BASE_DIR = Path(__file__).resolve().parent
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MAX_IMAGE_BYTES = 10 * 1024 * 1024
MAX_REPORT_BYTES = 25 * 1024 * 1024
security = HTTPBasic(auto_error=False)


class StorageError(Exception):
    pass


class CloudArchive:
    def __init__(self):
        self.url = os.environ.get("SUPABASE_URL", "").rstrip("/")
        self.key = os.environ.get("SUPABASE_SECRET_KEY", "")
        self.bucket = os.environ.get("SUPABASE_BUCKET", "welding-reports")
        if not self.url.startswith("https://") or not self.key:
            raise StorageError("ยังไม่ได้ตั้งค่า SUPABASE_URL และ SUPABASE_SECRET_KEY บน Render")
        self.headers = {"apikey": self.key}
        # รองรับทั้ง Secret key แบบใหม่และ service_role JWT แบบเดิม
        if not self.key.startswith("sb_secret_"):
            self.headers["Authorization"] = "Bearer " + self.key

    def request(self, method, path, **kwargs):
        extra = kwargs.pop("headers", {})
        try:
            with httpx.Client(timeout=httpx.Timeout(90, connect=15)) as client:
                response = client.request(method, self.url + path,
                                          headers={**self.headers, **extra}, **kwargs)
        except httpx.RequestError:
            logger.error("Cloud storage connection failed")
            raise StorageError("เชื่อมต่อคลังข้อมูลไม่ได้ กรุณาตรวจว่า Supabase พร้อมใช้งาน") from None
        if not response.is_success:
            # ไม่บันทึก response body ซึ่งอาจมีข้อมูลรายงานหรือ secrets
            logger.error("Cloud request failed: %s %s status=%s", method, path.split('?')[0], response.status_code)
            raise StorageError("บันทึกหรืออ่านคลังข้อมูลไม่สำเร็จ กรุณาตรวจการตั้งค่า Supabase และโควตา")
        return response

    def save(self, document, filename, fields):
        report_id = str(uuid4())
        # ลิงก์ผู้สร้างเป็น capability token: เก็บเฉพาะ hash ในฐานข้อมูล
        token = secrets.token_urlsafe(32)
        object_path = report_id + ".docx"
        path = "/storage/v1/object/" + quote(self.bucket, safe="") + "/" + object_path
        self.request("POST", path, content=document,
                     headers={"Content-Type": DOCX_MIME, "x-upsert": "false"})
        record = {
            "id": report_id, "project_title": fields.get("project_title", ""),
            "joint_no": fields.get("joint_no", ""), "welder_name": fields.get("welder_name", ""),
            "data_json": fields, "file_name": filename, "object_path": object_path,
            "file_size": len(document),
            "access_hash": hashlib.sha256(token.encode()).hexdigest(),
        }
        try:
            self.request("POST", "/rest/v1/welding_reports", json=record,
                         headers={"Prefer": "return=minimal"})
        except StorageError:
            # ตรวจกรณีตอบกลับหายทั้งที่ DB commit ไปแล้ว ก่อนลบ object
            try:
                if self.get(report_id):
                    return report_id, token
            except StorageError:
                # หากตรวจไม่ได้ คง object ไว้ให้ผู้ดูแลตรวจ แทนการลบไฟล์ที่อาจบันทึกสำเร็จ
                raise
            try:
                self.request("DELETE", "/storage/v1/object/" + quote(self.bucket, safe=""),
                             json={"prefixes": [object_path]})
            except StorageError:
                logger.error("Unlinked uploaded object needs review: %s", report_id)
            raise
        return report_id, token

    def get(self, report_id):
        rows = self.request("GET", "/rest/v1/welding_reports", params={
            "select": "*", "id": "eq." + report_id, "limit": "1"
        }).json()
        return rows[0] if rows else None

    def list(self, query="", page=1):
        params = {"select": "id,project_title,joint_no,welder_name,created_at,file_name,file_size",
                  "order": "created_at.desc", "limit": "21", "offset": str((page - 1) * 20)}
        if query.strip():
            params["search_text"] = "ilike.*" + query.strip().replace("*", "").replace("%", "") + "*"
        return self.request("GET", "/rest/v1/welding_reports", params=params).json()

    def download(self, record):
        return self.request("GET", "/storage/v1/object/authenticated/" + quote(self.bucket, safe="")
                            + "/" + quote(record["object_path"], safe="/")).content


def repository():
    return CloudArchive()


def archive_login(credentials: HTTPBasicCredentials = Depends(security)):
    password = os.environ.get("ARCHIVE_PASSWORD", "")
    username = os.environ.get("ARCHIVE_USERNAME", "admin")
    if not password:
        raise HTTPException(503, "กรุณาตั้ง ARCHIVE_PASSWORD บน Render ก่อนเข้าคลังรายงาน")
    if (credentials is None
        or not secrets.compare_digest(credentials.username.encode(), username.encode())
        or not secrets.compare_digest(credentials.password.encode(), password.encode())):
        raise HTTPException(401, "กรุณาเข้าสู่คลังรายงาน", headers={"WWW-Authenticate": 'Basic realm="Report Archive"'})
    return True


def escape(value):
    return html.escape(str(value or ""), quote=True)


def page(title, body, status=200):
    return HTMLResponse(f'''<!doctype html><html lang="th"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{escape(title)}</title>
<style>body{{font-family:system-ui,sans-serif;background:#eef2f5;color:#172b3a;margin:0;padding:20px}}
main{{max-width:950px;margin:auto;background:white;border-radius:12px;padding:24px}}
a{{color:#0056b3}}.button{{display:inline-block;background:#0056b3;color:white;padding:13px 18px;border-radius:6px;text-decoration:none;margin:8px 6px 8px 0}}
input{{padding:10px;max-width:90%;width:320px;border:1px solid #bbc6cd;border-radius:5px}}
button{{padding:10px}}table{{width:100%;border-collapse:collapse}}th,td{{text-align:left;padding:10px;border-bottom:1px solid #dce2e7}}
small{{color:#526673}}.scroll{{overflow-x:auto}}h1{{font-size:24px}}nav{{margin:16px 0}}</style></head>
<body><main><h1>{escape(title)}</h1>{body}</main></body></html>''', status_code=status,
        headers={"Cache-Control": "no-store", "Referrer-Policy": "no-referrer", "X-Content-Type-Options": "nosniff"})


def error_page(message, status=500):
    return page("ยังดำเนินการไม่สำเร็จ", f'<p>{escape(message)}</p><p><a href="/">กลับไปกรอกข้อมูล</a></p>', status)


def valid_id(report_id):
    try:
        return str(UUID(report_id))
    except ValueError:
        raise HTTPException(404, "ไม่พบรายงาน") from None


def file_response(data, filename):
    return Response(data, media_type=DOCX_MIME, headers={
        "Content-Disposition": "attachment; filename=\"Welding_Report.docx\"; filename*=UTF-8''" + quote(filename, safe=""),
        "Cache-Control": "private, no-store", "X-Content-Type-Options": "nosniff",
        "Referrer-Policy": "no-referrer",
    })


@app.get("/", response_class=HTMLResponse)
def home():
    return HTMLResponse((BASE_DIR / "index.html").read_text(encoding="utf-8"))


@app.get("/health")
def health():
    return {"status": "ok", "cloud_configured": bool(os.environ.get("SUPABASE_URL") and os.environ.get("SUPABASE_SECRET_KEY")),
            "archive_password_configured": bool(os.environ.get("ARCHIVE_PASSWORD"))}


@app.post("/submit")
async def submit(request: Request):
    try:
        # ตรวจการตั้งค่าก่อนสร้างไฟล์ เพื่อไม่แจ้งว่าบันทึกสำเร็จทั้งที่ไม่ได้เก็บเข้าคลัง
        cloud = repository()
        fields, images = {}, {}
        async with request.form() as form:
            for key, value in form.multi_items():
                if isinstance(value, UploadFile):
                    if key not in PHOTO_SLOTS or not value.filename:
                        continue
                    content = await value.read(MAX_IMAGE_BYTES + 1)
                    if len(content) > MAX_IMAGE_BYTES:
                        return error_page("รูปแต่ละรูปต้องไม่เกิน 10 MB", 413)
                    if content:
                        images[key] = content
                else:
                    text = str(value).strip()
                    if len(text) > 500:
                        return error_page("ข้อความแต่ละช่องต้องไม่เกิน 500 ตัวอักษร", 400)
                    fields[key] = text
        if not fields.get("project_title") or not fields.get("joint_no"):
            return error_page("กรุณากรอกโครงการและหมายเลขรอยต่อ", 400)
        document, filename, context = await run_in_threadpool(build_report, fields, images)
        if len(document) > MAX_REPORT_BYTES:
            return error_page("รายงานใหญ่เกิน 25 MB กรุณาลดขนาดรูป", 413)
        report_id, token = await run_in_threadpool(cloud.save, document, filename, context)
        return RedirectResponse(f"/report/{report_id}?token={token}", status_code=303,
                                headers={"Cache-Control": "no-store", "Referrer-Policy": "no-referrer"})
    except ValueError as exc:
        return error_page(str(exc), 400)
    except StorageError as exc:
        return error_page(str(exc), 503)
    except TemplateError:
        logger.error("Template syntax error")
        return error_page("แม่แบบ Word ไม่ถูกต้อง กรุณาใช้ template.docx จากชุดนี้")
    except Exception:
        logger.error("Report creation failed", exc_info=False)
        return error_page("สร้างรายงานไม่สำเร็จ กรุณาตรวจไฟล์แม่แบบและการตั้งค่าระบบ")


def archive_ticket(report_id):
    expires = str(int(time.time()) + 3600)
    secret = os.environ.get("ARCHIVE_PASSWORD", "").encode()
    signature = hmac.new(secret, (report_id + ":" + expires).encode(), hashlib.sha256).hexdigest()
    return expires + "." + signature


def ticket_valid(report_id, ticket):
    try:
        expires, signature = ticket.split(".", 1)
        if int(expires) < time.time() or not os.environ.get("ARCHIVE_PASSWORD"):
            return False
        expected = hmac.new(os.environ["ARCHIVE_PASSWORD"].encode(),
                            (report_id + ":" + expires).encode(), hashlib.sha256).hexdigest()
        return hmac.compare_digest(signature, expected)
    except (ValueError, TypeError):
        return False


def creator_record(report_id, token):
    record = repository().get(valid_id(report_id))
    digest = hashlib.sha256(token.encode()).hexdigest()
    if not record or not hmac.compare_digest(digest, record["access_hash"]):
        raise HTTPException(404, "ไม่พบรายงานหรือลิงก์ไม่ถูกต้อง")
    return record


@app.get("/report/{report_id}")
def report_page(report_id: str, token: str = ""):
    try:
        record = creator_record(report_id, token)
        url = f"/download/{record['id']}?token=" + quote(token, safe="")
        return page("บันทึกรายงานสำเร็จ", f'''
<p>รายงานและไฟล์ Word ถูกบันทึกเข้าคลังแล้ว</p>
<p><strong>โครงการ:</strong> {escape(record['project_title'])}<br>
<strong>หมายเลขรอยต่อ:</strong> {escape(record['joint_no'])}<br>
<strong>ไฟล์:</strong> {escape(record['file_name'])}</p>
<a class="button" href="{escape(url)}">ดาวน์โหลด Word</a>
<a class="button" href="/reports">เปิดคลังรายงานย้อนหลัง</a>
<p><a href="/">สร้างรายงานใหม่</a></p>
<p><small>หากเปิดผ่าน LINE แล้วดาวน์โหลดไม่ได้ ให้เปิดหน้านี้ใน Chrome หรือ Safari</small></p>''')
    except StorageError as exc:
        return error_page(str(exc), 503)


@app.get("/download/{report_id}")
def creator_download(report_id: str, token: str = "", ticket: str = ""):
    try:
        report_id = valid_id(report_id)
        if ticket_valid(report_id, ticket):
            record = repository().get(report_id)
            if not record:
                raise HTTPException(404, "ไม่พบรายงาน")
        else:
            record = creator_record(report_id, token)
        return file_response(repository().download(record), record["file_name"])
    except StorageError as exc:
        return error_page(str(exc), 503)


@app.get("/reports")
def archive(q: str = "", page_number: int = 1, logged_in=Depends(archive_login)):
    if len(q) > 100 or page_number < 1:
        raise HTTPException(400, "คำค้นหรือหมายเลขหน้าไม่ถูกต้อง")
    try:
        records = repository().list(q, page_number)
    except StorageError as exc:
        return error_page(str(exc), 503)
    rows = ""
    for record in records[:20]:
        timestamp = datetime.fromisoformat(record["created_at"].replace("Z", "+00:00")).astimezone(ZoneInfo("Asia/Bangkok"))
        rows += f'''<tr><td>{timestamp.strftime('%d/%m/%Y %H:%M')}</td>
<td>{escape(record['project_title'])}</td><td>{escape(record['joint_no'])}</td>
<td>{escape(record['welder_name'])}</td>
<td><a href="/download/{escape(record['id'])}?ticket={archive_ticket(record['id'])}">ดาวน์โหลด Word</a></td></tr>'''
    navigation = ""
    if page_number > 1:
        navigation += '<a href="/reports?' + escape(urlencode({"q": q, "page_number": page_number - 1})) + '">หน้าก่อนหน้า</a> '
    if len(records) > 20:
        navigation += '<a href="/reports?' + escape(urlencode({"q": q, "page_number": page_number + 1})) + '">หน้าถัดไป</a>'
    return page("คลังรายงานงานเชื่อม", f'''
<p><a href="/">สร้างรายงานใหม่</a></p>
<form method="get"><input name="q" value="{escape(q)}" placeholder="ค้นหาโครงการ รอยต่อ หรือช่างเชื่อม"><button>ค้นหา</button></form>
<p>หน้า {page_number} · เวลาไทย</p><div class="scroll"><table><thead><tr><th>วันที่บันทึก</th><th>โครงการ</th><th>รอยต่อ</th><th>ช่างเชื่อม</th><th>รายงาน</th></tr></thead>
<tbody>{rows or '<tr><td colspan="5">ยังไม่มีรายงานที่ตรงกับคำค้น</td></tr>'}</tbody></table></div><nav>{navigation}</nav>''')


@app.get("/reports/{report_id}/download")
def archive_download(report_id: str, logged_in=Depends(archive_login)):
    try:
        record = repository().get(valid_id(report_id))
        if not record:
            raise HTTPException(404, "ไม่พบรายงาน")
        return RedirectResponse(f"/download/{record['id']}?ticket={archive_ticket(record['id'])}",
                                status_code=303, headers={"Cache-Control": "no-store"})
    except StorageError as exc:
        return error_page(str(exc), 503)
