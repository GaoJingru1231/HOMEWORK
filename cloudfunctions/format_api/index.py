"""格式快调 - 腾讯云 SCF 云函数入口（HTTP 触发器 + 集成响应）。

部署：把本文件 + formatter.py + python-docx 依赖一起打包成函数 zip。
详见 部署到腾讯云.md。

与本地 FastAPI 服务 (server.py) 共用 formatter.py 的排版逻辑，
仅在此处实现 HTTP 适配（multipart 解析、路由、base64 响应）。
"""

import os
import sys
import io
import json
import uuid
import base64
import re
import tempfile
from pathlib import Path

# 让 formatter 既能在打包后的函数目录里被导入，也能在本地从项目根目录导入
_HERE = os.path.dirname(os.path.abspath(__file__))
for _p in (_HERE, os.path.dirname(os.path.dirname(_HERE))):
    if _p and _p not in sys.path:
        sys.path.insert(0, _p)

from docx import Document  # noqa: E402
from formatter import (  # noqa: E402
    PRESET_TEMPLATES, apply_document_format, detect_format_issues,
    is_valid_doc_file, resolve_params,
)

# SCF 可写目录（/tmp）；与 server.py 的回退目录同名，便于本地联调
OUT_DIR = Path(tempfile.gettempdir()) / "fmt_kuaitiao"
OUT_DIR.mkdir(parents=True, exist_ok=True)

DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


# ===== HTTP 响应辅助 =====

CORS_HEADERS = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
}


def _json_resp(data, status=200):
    return {
        "isBase64Encoded": False,
        "statusCode": status,
        "headers": {**{"Content-Type": "application/json; charset=utf-8"}, **CORS_HEADERS},
        "body": json.dumps(data, ensure_ascii=False),
    }


def _err(detail, status=400):
    return _json_resp({"detail": detail}, status)


# ===== multipart/form-data 解析（无第三方依赖）=====

def parse_multipart(body: bytes, content_type: str):
    """解析 multipart/form-data，返回 (fields dict, files dict)。
    files[name] = {"filename": str, "data": bytes}。"""
    m = re.search(r'boundary=("?)([^";]+)\1', content_type or "")
    if not m:
        return {}, {}
    boundary = m.group(2)
    delim = ("--" + boundary).encode()
    fields, files = {}, {}
    segments = body.split(delim)
    # 首段是 preamble，末段是 epilogue (--)，跳过
    for seg in segments[1:-1]:
        if seg.startswith(b"\r\n"):
            seg = seg[2:]
        if seg.endswith(b"\r\n"):
            seg = seg[:-2]
        hsep = seg.find(b"\r\n\r\n")
        if hsep == -1:
            continue
        header = seg[:hsep].decode("utf-8", errors="ignore")
        content = seg[hsep + 4:]
        name_m = re.search(r'name="([^"]+)"', header)
        if not name_m:
            continue
        name = name_m.group(1)
        fn_m = re.search(r'filename="([^"]*)"', header)
        if fn_m:
            files[name] = {"filename": fn_m.group(1), "data": content}
        else:
            fields[name] = content.decode("utf-8", errors="ignore")
    return fields, files


# ===== 路由处理 =====

def handle_health():
    return _json_resp({"status": "ok"})


def handle_templates():
    """列出所有预设模板(与 server.py 的 /api/templates 对齐)"""
    templates = []
    for key, params in PRESET_TEMPLATES.items():
        templates.append({
            "id": key,
            "name": key,
            "params": params.to_dict(),
        })
    return _json_resp({"templates": templates})


def handle_detect(fields, files):
    upload = files.get("file")
    if not upload:
        return _err("未收到文件 (file 字段缺失)")
    filename = upload["filename"]
    if not is_valid_doc_file(filename, None):
        return _err("仅支持 .docx / .doc 格式，请确认文件后缀正确")
    try:
        doc = Document(io.BytesIO(upload["data"]))
    except Exception:
        return _err("文件格式无效，请上传有效的 .docx 文档", 400)

    params = PRESET_TEMPLATES["undergrad_thesis"]
    report = detect_format_issues(doc, params)
    return _json_resp({"filename": filename, "report": report.to_dict()})


def handle_format(fields, files):
    upload = files.get("file")
    if not upload:
        return _err("未收到文件 (file 字段缺失)")
    filename = upload["filename"]
    if not is_valid_doc_file(filename, None):
        return _err("仅支持 .docx / .doc 格式，请确认文件后缀正确")
    try:
        doc = Document(io.BytesIO(upload["data"]))
    except Exception:
        return _err("文件格式无效，请上传有效的 .docx 文档", 400)

    template = fields.get("template") or "undergrad_thesis"
    params = resolve_params(template, fields)

    report = detect_format_issues(doc, params)
    modified_count = apply_document_format(doc, params)

    task_id = uuid.uuid4().hex[:12]
    original_name = Path(filename).stem
    output_name = f"{original_name}_已排版_{template}.docx"
    output_path = OUT_DIR / f"{task_id}_{output_name}"
    doc.save(str(output_path))

    # SCF 多实例无共享 /tmp，直接把文件 base64 内联返回，前端用 Blob 下载
    file_b64 = base64.b64encode(output_path.read_bytes()).decode("ascii")

    return _json_resp({
        "task_id": task_id,
        "original_filename": filename,
        "output_filename": output_name,
        "modified_paragraphs": modified_count,
        "report": report.to_dict(),
        "download_url": f"/api/download/{task_id}/{output_name}",
        "file_base64": file_b64,
        "file_mime": DOCX_MIME,
        "template_used": template,
    })


def handle_download(path):
    rest = path.split("/api/download/")[-1].strip("/").split("/")
    if len(rest) >= 2:
        task_id, fname = rest[0], rest[1]
        fpath = OUT_DIR / f"{task_id}_{fname}"
        if fpath.exists():
            data = fpath.read_bytes()
            return {
                "isBase64Encoded": True,
                "statusCode": 200,
                "headers": {
                    "Content-Type": DOCX_MIME,
                    "Content-Disposition": f'attachment; filename="{fname}"',
                    **CORS_HEADERS,
                },
                "body": base64.b64encode(data).decode("ascii"),
            }
    return _err("文件不存在或已过期", 404)


# ===== 入口 =====

def main(event, context):
    method = (event.get("httpMethod") or event.get("method") or "GET").upper()
    path = event.get("path") or ""

    # CORS 预检
    if method == "OPTIONS":
        return {"isBase64Encoded": False, "statusCode": 204, "headers": CORS_HEADERS, "body": ""}

    # 健康检查
    if path.endswith("/api/health"):
        return handle_health()

    # 模板列表
    if path.endswith("/api/templates") and method == "GET":
        return handle_templates()

    # 排版：返回文件 base64
    if path.endswith("/api/format") and method == "POST":
        headers = event.get("headers") or {}
        ct = headers.get("Content-Type") or headers.get("content-type") or ""
        raw = event.get("body") or ""
        if event.get("isBase64Encoded"):
            body = base64.b64decode(raw)
        else:
            body = raw.encode("utf-8") if isinstance(raw, str) else raw
        fields, files = parse_multipart(body, ct)
        return handle_format(fields, files)

    # 检测
    if path.endswith("/api/detect") and method == "POST":
        headers = event.get("headers") or {}
        ct = headers.get("Content-Type") or headers.get("content-type") or ""
        raw = event.get("body") or ""
        if event.get("isBase64Encoded"):
            body = base64.b64decode(raw)
        else:
            body = raw.encode("utf-8") if isinstance(raw, str) else raw
        fields, files = parse_multipart(body, ct)
        return handle_detect(fields, files)

    # 下载（同实例 best-effort；前端主要走 file_base64）
    if "/api/download/" in path and method == "GET":
        return handle_download(path)

    return _err(f"Not Found: {method} {path}", 404)


# 兼容部分模板使用 main_handler 作为入口
main_handler = main
