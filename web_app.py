"""
圖面識別量測檢表 — Web 應用程式
啟動：uvicorn drawing_checker.web_app:app --reload --port 8001
"""

import os
import sys
import base64
import tempfile

import anthropic as _anthropic
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.requests import Request
from pydantic import BaseModel
from dotenv import load_dotenv

load_dotenv()

# 加入上層目錄到 path（讓 analyzer / pdf_converter 可被 import）
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from analyzer import analyze_drawing_image, analyze_multiple_images
from pdf_converter import pdf_to_images, cleanup_temp_images
from preprocess import pdf_first_page_thumbnail
import db

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))

app = FastAPI(title="圖面識別量測檢表")

ALLOWED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp", ".gif", ".pdf"}


@app.on_event("startup")
async def startup():
    db.init_db()


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/api/thumbnail")
async def get_thumbnail(file: UploadFile = File(...)):
    """PDF 上傳後即時取得第一頁縮圖（用於預覽）"""
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext != ".pdf":
        return JSONResponse({"thumbnail": None})
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(await file.read())
        tmp_path = tmp.name
    try:
        thumb = pdf_first_page_thumbnail(tmp_path, max_w=800)
        return JSONResponse({"thumbnail": thumb})
    finally:
        os.unlink(tmp_path)


@app.post("/api/analyze")
async def analyze(
    file: UploadFile = File(...),
    all_pages: bool = False,
    page: int = 1,
):
    """
    接收圖面檔案，回傳尺寸分析結果 JSON
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY 未設定")

    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail=f"不支援的格式：{ext}")

    # 儲存上傳的檔案到臨時目錄
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        content = await file.read()
        tmp.write(content)
        tmp_path = tmp.name

    # 同時回傳圖片 base64 供前端預覽
    preview_b64 = None
    temp_images = []

    try:
        if ext == ".pdf":
            temp_images = pdf_to_images(tmp_path, dpi=250)
            if not temp_images:
                raise HTTPException(status_code=400, detail="PDF 轉換失敗")

            if all_pages and len(temp_images) > 1:
                data = analyze_multiple_images(temp_images, api_key)
                preview_path = temp_images[0]
            else:
                page_idx = max(0, min(page - 1, len(temp_images) - 1))
                data = analyze_drawing_image(temp_images[page_idx], api_key)
                preview_path = temp_images[page_idx]

            with open(preview_path, "rb") as f:
                preview_b64 = "data:image/png;base64," + base64.b64encode(f.read()).decode()

        else:
            data = analyze_drawing_image(tmp_path, api_key)
            media_type = {
                ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
                ".png": "image/png", ".webp": "image/webp", ".gif": "image/gif",
            }.get(ext, "image/png")
            preview_b64 = f"data:{media_type};base64," + base64.b64encode(content).decode()

    except _anthropic.AuthenticationError:
        raise HTTPException(status_code=401, detail="API Key 無效，請確認 ANTHROPIC_API_KEY")
    except _anthropic.BadRequestError as e:
        msg = str(e)
        if "credit balance" in msg.lower():
            raise HTTPException(status_code=402, detail="Anthropic API 帳戶餘額不足，請前往 console.anthropic.com 充值")
        raise HTTPException(status_code=400, detail=f"API 請求錯誤：{msg[:200]}")
    except _anthropic.RateLimitError:
        raise HTTPException(status_code=429, detail="API 請求頻率超限，請稍後再試")
    except _anthropic.APIError as e:
        raise HTTPException(status_code=502, detail=f"AI 服務暫時無法使用：{str(e)[:200]}")
    finally:
        os.unlink(tmp_path)
        if temp_images:
            cleanup_temp_images(temp_images)

    dims = data.get("dimensions", [])
    resp = {
        "success": True,
        "preview": preview_b64,
        "part_name": data.get("part_name", "Unknown"),
        "drawing_no": data.get("drawing_no", "N/A"),
        "has_yellow_marks": data.get("has_yellow_marks", False),
        "dimensions": dims,
    }
    # 若 0 項且有 parse error，把提示帶給前端
    if len(dims) == 0 and data.get("_parse_error"):
        resp["_warn"] = f"AI 回應解析失敗，請重新上傳或確認圖面清晰度。原始片段：{data.get('_raw','')[:100]}"

    return JSONResponse(resp)


# ── 檢表儲存 CRUD ─────────────────────────────────────

class SaveRequest(BaseModel):
    part_name: str
    drawing_no: str = ""
    internal_no: str = ""
    dimensions: list
    tools: dict = {}
    preview: str = ""


@app.post("/api/checklists")
async def save_checklist(req: SaveRequest):
    cid = db.save(
        req.part_name, req.drawing_no, req.internal_no,
        req.dimensions, req.tools, req.preview
    )
    return JSONResponse({"id": cid, "count": db.count()})


@app.get("/api/checklists")
async def list_checklists():
    rows = db.list_all()
    # preview 縮略（只取前 8000 字元給清單頁，避免傳輸過大）
    for r in rows:
        if r.get("preview_b64") and len(r["preview_b64"]) > 8000:
            r["preview_b64"] = r["preview_b64"][:8000]
    return JSONResponse({"items": rows, "count": len(rows)})


@app.get("/api/checklists/{cid}")
async def get_checklist(cid: int):
    row = db.get(cid)
    if not row:
        raise HTTPException(status_code=404, detail="找不到該檢表")
    return JSONResponse(row)


@app.delete("/api/checklists/{cid}")
async def delete_checklist(cid: int):
    db.delete(cid)
    return JSONResponse({"count": db.count()})
