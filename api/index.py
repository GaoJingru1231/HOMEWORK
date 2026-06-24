#!/usr/bin/env python3
"""
Vercel Serverless Function - 格式快调后端
"""
import sys
import os

# 添加依赖路径
sys.path.append('/var/task')
sys.path.append('/tmp')

from http.server import BaseHTTPRequestHandler
from urllib.parse import parse_qs, urlparse
import json
import io
import tempfile

# 尝试导入依赖
try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
except ImportError:
    import subprocess
    subprocess.run([sys.executable, '-m', 'pip', 'install', 'python-docx', '-t', '/tmp'], check=True)
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH

def handler(request):
    from http.server import BaseHTTPRequestHandler
    return {
        "statusCode": 200,
        "body": json.dumps({"status": "ok", "message": "格式快调 API 正常运行"}),
        "headers": {"Content-Type": "application/json"}
    }
