"""
格式快调 - 后端 API 服务
基于 FastAPI + python-docx 实现文档格式检测与一键排版。

排版核心逻辑见 formatter.py（与腾讯云 SCF 云函数共用）。
"""

import io
import os
import uuid
import tempfile
from pathlib import Path
from datetime import datetime
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, JSONResponse, HTMLResponse
from fastapi.middleware.cors import CORSMiddleware

from docx import Document

from formatter import (
    FormatParams, PRESET_TEMPLATES,
    apply_document_format, detect_format_issues, is_valid_doc_file, resolve_params,
)

app = FastAPI(title="格式快调 API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 配置
BASE_DIR = Path(__file__).parent

def _resolve_writable_dir(local_subdir: str, env_key: str) -> Path:
    """选择可写目录：优先环境变量；其次项目内 local_subdir（本地开发可见）；
    不可写时回退系统临时目录（云端项目目录只读场景）。"""
    env_val = os.environ.get(env_key)
    if env_val:
        p = Path(env_val)
        try:
            p.mkdir(parents=True, exist_ok=True)
            return p
        except OSError:
            pass
    local = BASE_DIR / local_subdir
    try:
        local.mkdir(parents=True, exist_ok=True)
        probe = local / ".writetest"
        probe.touch()
        probe.unlink(missing_ok=True)
        return local
    except OSError:
        tmp = Path(tempfile.gettempdir()) / "fmt_kuaitiao"
        tmp.mkdir(parents=True, exist_ok=True)
        return tmp

UPLOAD_DIR = _resolve_writable_dir("uploads", "FMT_UPLOAD_DIR")
PROCESSED_DIR = _resolve_writable_dir("processed", "FMT_PROCESSED_DIR")


# ===== API 路由 =====

@app.get("/api/health")
async def health_check():
    return {"status": "ok", "timestamp": datetime.now().isoformat()}


@app.get("/api/templates")
async def list_templates():
    """列出所有预设模板"""
    templates = []
    for key, params in PRESET_TEMPLATES.items():
        templates.append({
            "id": key,
            "name": key,
            "params": params.to_dict(),
        })
    return {"templates": templates}


@app.post("/api/detect")
async def detect_format(file: UploadFile = File(...)):
    """
    上传文档并检测格式问题，返回检测报告。
    使用默认"本科毕业论文"模板参数作为检测基准。
    """
    if not is_valid_doc_file(file.filename, file.content_type):
        raise HTTPException(status_code=400, detail="仅支持 .docx / .doc 格式，请确认文件后缀正确")

    contents = await file.read()
    try:
        doc = Document(io.BytesIO(contents))
    except Exception:
        raise HTTPException(status_code=400, detail="文件格式无效，请上传有效的 .docx 文档")

    params = PRESET_TEMPLATES["undergrad_thesis"]
    report = detect_format_issues(doc, params)

    return {"filename": file.filename, "report": report.to_dict()}


@app.post("/api/format")
async def format_document(
    file: UploadFile = File(...),
    template: str = Form("undergrad_thesis"),
    font_cn: Optional[str] = Form(None),
    font_en: Optional[str] = Form(None),
    font_size: Optional[float] = Form(None),
    line_spacing: Optional[float] = Form(None),
    first_line_indent: Optional[int] = Form(None),
    alignment: Optional[str] = Form(None),
    space_before: Optional[float] = Form(None),
    space_after: Optional[float] = Form(None),
    margin_top: Optional[float] = Form(None),
    margin_bottom: Optional[float] = Form(None),
    margin_left: Optional[float] = Form(None),
    margin_right: Optional[float] = Form(None),
    h1_size: Optional[float] = Form(None),
    h1_bold: Optional[bool] = Form(None),
    h2_size: Optional[float] = Form(None),
    h2_bold: Optional[bool] = Form(None),
):
    """上传文档并按指定模板 + 可选自定义参数进行排版，返回处理后的文档。"""
    if not is_valid_doc_file(file.filename, file.content_type):
        raise HTTPException(status_code=400, detail="仅支持 .docx / .doc 格式，请确认文件后缀正确")

    contents = await file.read()
    try:
        doc = Document(io.BytesIO(contents))
    except Exception:
        raise HTTPException(status_code=400, detail="文件格式无效，请上传有效的 .docx 文档")

    overrides = {
        "font_cn": font_cn, "font_en": font_en, "font_size": font_size,
        "line_spacing": line_spacing, "first_line_indent": first_line_indent,
        "alignment": alignment, "space_before": space_before, "space_after": space_after,
        "margin_top": margin_top, "margin_bottom": margin_bottom,
        "margin_left": margin_left, "margin_right": margin_right,
        "h1_size": h1_size, "h1_bold": h1_bold, "h2_size": h2_size, "h2_bold": h2_bold,
    }
    params = resolve_params(template, overrides)

    report = detect_format_issues(doc, params)
    modified_count = apply_document_format(doc, params)

    task_id = uuid.uuid4().hex[:12]
    original_name = Path(file.filename).stem
    output_name = f"{original_name}_已排版_{template}.docx"
    output_path = PROCESSED_DIR / f"{task_id}_{output_name}"
    doc.save(str(output_path))

    return JSONResponse({
        "task_id": task_id,
        "original_filename": file.filename,
        "output_filename": output_name,
        "modified_paragraphs": modified_count,
        "report": report.to_dict(),
        "download_url": f"/api/download/{task_id}/{output_name}",
        "template_used": template,
    })


@app.get("/api/download/{task_id}/{filename}")
async def download_file(task_id: str, filename: str):
    """下载处理后的文档"""
    file_path = PROCESSED_DIR / f"{task_id}_{filename}"
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在或已过期")

    return FileResponse(
        path=str(file_path),
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ===== 静态文件服务 =====

INDEX_PATH = BASE_DIR / "index.html"

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    if INDEX_PATH.exists():
        return HTMLResponse(
            content=INDEX_PATH.read_text(encoding="utf-8"),
            headers={"Cache-Control": "no-cache, no-store, must-revalidate", "Pragma": "no-cache", "Expires": "0"}
        )
    raise HTTPException(status_code=404, detail="index.html not found")


# ===== 启动入口 =====
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("server:app", host="0.0.0.0", port=port, reload=False)
