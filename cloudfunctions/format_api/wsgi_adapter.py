"""WSGI 适配器:把标准 WSGI 请求转成 SCF event,复用 index.main_handler。

腾讯云 Web 函数(Type=HTTP)通过 scf_bootstrap 启动一个监听 9000 端口的
HTTP server,公网请求转发到该 server。本文件用 Python 标准库 wsgiref 启动
server,无需额外依赖。

启动方式(由 scf_bootstrap 调用):
    python wsgi_adapter.py
"""
import os
import base64
from wsgiref.simple_server import make_server

from index import main_handler


def _scf_event_from_wsgi(environ):
    """WSGI environ -> SCF event dict(集成响应输入格式)。"""
    method = environ.get("REQUEST_METHOD", "GET").upper()
    path = environ.get("PATH_INFO", "/")
    query = environ.get("QUERY_STRING", "")
    full_path = f"{path}?{query}" if query else path

    # headers:WSGI 把 HTTP header 加 HTTP_ 前缀大写
    headers = {}
    for k, v in environ.items():
        if k.startswith("HTTP_"):
            headers[k[5:].replace("_", "-").title()] = v
        elif k == "CONTENT_TYPE":
            headers["Content-Type"] = v
        elif k == "CONTENT_LENGTH":
            headers["Content-Length"] = v

    # body
    try:
        length = int(environ.get("CONTENT_LENGTH") or 0)
    except (TypeError, ValueError):
        length = 0
    body = environ["wsgi.input"].read(length) if length > 0 else b""

    # multipart / 二进制用 base64,与 index.py 的解码逻辑对应
    ct = headers.get("Content-Type", "")
    is_base64 = False
    if body and (ct.startswith("multipart/") or "application/octet-stream" in ct):
        is_base64 = True
        body_str = base64.b64encode(body).decode("ascii")
    else:
        body_str = body.decode("utf-8", errors="replace") if body else ""

    return {
        "httpMethod": method,
        "path": full_path,
        "headers": headers,
        "body": body_str,
        "isBase64Encoded": is_base64,
    }


def wsgi_app(environ, start_response):
    """WSGI 应用:转 event -> 调 main_handler -> 转回 WSGI 响应。"""
    event = _scf_event_from_wsgi(environ)
    try:
        resp = main_handler(event, None)
    except Exception as e:
        resp = {
            "isBase64Encoded": False,
            "statusCode": 500,
            "headers": {"Content-Type": "text/plain; charset=utf-8"},
            "body": f"Internal Error: {e}",
        }

    status_code = resp.get("statusCode", 200)
    # WSGI 状态行需要标准 HTTP 状态码 + 原因短语
    import http.client
    try:
        reason = http.client.responses.get(status_code, "OK")
    except Exception:
        reason = "OK"
    status_line = f"{status_code} {reason}"

    headers_list = [(k, str(v)) for k, v in (resp.get("headers") or {}).items()]
    body = resp.get("body", "")
    if resp.get("isBase64Encoded"):
        try:
            body_bytes = base64.b64decode(body)
        except Exception:
            body_bytes = b""
    else:
        body_bytes = body.encode("utf-8") if isinstance(body, str) else (body or b"")
    start_response(status_line, headers_list)
    return [body_bytes]


if __name__ == "__main__":
    port = int(os.environ.get("SCF_SERVER_PORT", "9000"))
    httpd = make_server("", port, wsgi_app)
    print(f"wsgi_adapter serving on port {port}", flush=True)
    httpd.serve_forever()
