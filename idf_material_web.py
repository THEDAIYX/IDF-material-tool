"""Local web front end for the IDF material extractor.

Usage:
    python idf_material_web.py start
    python idf_material_web.py stop
    python idf_material_web.py status

The server binds to 127.0.0.1 by default and never uploads files anywhere.
"""

from __future__ import annotations

import argparse
import json
import os
import signal
import subprocess
import sys
import tempfile
import time
import urllib.request
import webbrowser
from email.parser import BytesParser
from email.policy import default
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlparse

from idf_material_extractor import (
    DETAIL_HEADERS,
    extract_idf,
    extract_idfs,
    export_new_workbook,
    export_to_custom_template,
)


ROOT = Path(__file__).resolve().parent
WEB_ROOT = ROOT / "web"
PID_FILE = ROOT / ".idf_material_web.pid"
LOG_FILE = ROOT / ".idf_material_web.log"
DEFAULT_HOST = "127.0.0.1"
DEFAULT_PORT = 8765
MAX_UPLOAD_BYTES = 30 * 1024 * 1024


def _pid_is_alive(pid: int) -> bool:
    try:
        os.kill(pid, 0)
        return True
    except (OSError, ValueError):
        return False


def _read_pid() -> int | None:
    try:
        return int(PID_FILE.read_text(encoding="ascii").strip())
    except (OSError, ValueError):
        return None


def _write_pid() -> None:
    PID_FILE.write_text(str(os.getpid()), encoding="ascii")


def _remove_pid() -> None:
    try:
        PID_FILE.unlink()
    except FileNotFoundError:
        pass


def _safe_stem(filename: str) -> str:
    stem = Path(filename or "idf").stem
    cleaned = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in stem)
    return cleaned[:80] or "idf"


def _parse_multipart(handler: BaseHTTPRequestHandler) -> dict[str, dict[str, Any]]:
    content_type = handler.headers.get("Content-Type", "")
    content_length = int(handler.headers.get("Content-Length", "0") or 0)
    if content_length <= 0:
        raise ValueError("请求中没有上传内容。")
    if content_length > MAX_UPLOAD_BYTES:
        raise ValueError(f"上传文件过大，限制为 {MAX_UPLOAD_BYTES // 1024 // 1024} MB。")
    if "multipart/form-data" not in content_type:
        raise ValueError("请求格式必须是 multipart/form-data。")
    body = handler.rfile.read(content_length)
    header = (
        f"Content-Type: {content_type}\r\n"
        "MIME-Version: 1.0\r\n\r\n"
    ).encode("utf-8")
    message = BytesParser(policy=default).parsebytes(header + body)
    fields: dict[str, dict[str, Any]] = {}
    for part in message.iter_parts():
        name = part.get_param("name", header="content-disposition")
        if not name:
            continue
        payload = part.get_payload(decode=True) or b""
        filename = part.get_filename()
        if filename:
            item = {"filename": filename, "content": payload}
            if name == "idf":
                fields.setdefault(name, []).append(item)
            else:
                fields[name] = item
        else:
            fields[name] = {"value": payload.decode("utf-8", errors="replace")}
    return fields


class WebHandler(BaseHTTPRequestHandler):
    server_version = "IDFMaterialWeb/1.0"

    def log_message(self, fmt: str, *args: Any) -> None:
        # Keep the console useful without printing uploaded file contents.
        sys.stdout.write(f"[{self.log_date_time_string()}] {fmt % args}\n")

    def _send_json(self, data: dict[str, Any], status: int = 200) -> None:
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def _send_bytes(self, payload: bytes, filename: str, content_type: str) -> None:
        ascii_name = "".join(ch if ord(ch) < 128 and ch not in '"\\' else "_" for ch in filename)
        encoded_name = quote(filename, safe="")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header(
            "Content-Disposition",
            f'attachment; filename="{ascii_name}"; filename*=UTF-8\'\'{encoded_name}',
        )
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        if path == "/api/health":
            self._send_json({"ok": True, "service": "idf-material"})
            return
        if path == "/":
            path = "/index.html"
        assets = {
            "/index.html": (WEB_ROOT / "index.html", "text/html; charset=utf-8"),
            "/static/app.js": (WEB_ROOT / "app.js", "text/javascript; charset=utf-8"),
            "/static/styles.css": (WEB_ROOT / "styles.css", "text/css; charset=utf-8"),
        }
        target = assets.get(path)
        if not target or not target[0].is_file():
            self._send_json({"ok": False, "error": "页面不存在。"}, HTTPStatus.NOT_FOUND)
            return
        payload = target[0].read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", target[1])
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)

    def do_POST(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        if path not in {"/api/preview", "/api/export"}:
            self._send_json({"ok": False, "error": "接口不存在。"}, HTTPStatus.NOT_FOUND)
            return
        try:
            fields = _parse_multipart(self)
            idfs = fields.get("idf", [])
            if isinstance(idfs, dict):
                idfs = [idfs]
            if not idfs or any(not item.get("content") for item in idfs):
                raise ValueError("请上传至少一个 IDF 文件。")
            with tempfile.TemporaryDirectory(prefix="idf-material-web-") as temp_dir:
                temp_paths = []
                for index, idf in enumerate(idfs, 1):
                    safe_name = Path(idf.get("filename") or f"input_{index}.idf").name
                    temp_path = Path(temp_dir) / f"{index}_{safe_name}"
                    temp_path.write_bytes(idf["content"])
                    temp_paths.append(temp_path)
                data = extract_idfs(temp_paths)
                if path == "/api/preview":
                    self._send_json(
                        {
                            "ok": True,
                            "file_name": data["file_name"],
                            "file_count": len(idfs),
                            "encoding": data["encoding"],
                            "material_count": len(data["summary_rows"]),
                            "detail_count": len(data["detail_rows"]),
                            "headers": DETAIL_HEADERS,
                            "rows": data["detail_rows"],
                        }
                    )
                    return

                mode = fields.get("mode", {}).get("value", "new")
                output_name = "批量IDF_材料清单.xlsx" if len(idfs) > 1 else f"{_safe_stem(idfs[0].get('filename', 'idf'))}_材料清单.xlsx"
                output_path = Path(temp_dir) / output_name
                if mode == "custom":
                    template = fields.get("template")
                    if not template or not template.get("content"):
                        raise ValueError("自定义模板模式必须上传 Excel 模板。")
                    template_path = Path(temp_dir) / (template.get("filename") or "template.xlsx")
                    template_path.write_bytes(template["content"])
                    sheet_name = fields.get("sheet_name", {}).get("value", "")
                    header_text = fields.get("header_row", {}).get("value", "")
                    header_row = int(header_text) if header_text.strip().isdigit() else 0
                    prefix = "批量IDF" if len(idfs) > 1 else _safe_stem(idfs[0].get('filename', 'idf'))
                    output_name = f"{prefix}_填充结果{template_path.suffix.lower()}"
                    output_path = Path(temp_dir) / output_name
                    export_to_custom_template(data, template_path, output_path, sheet_name, header_row)
                else:
                    export_new_workbook(data, output_path)
                payload = output_path.read_bytes()
                content_type = (
                    "application/vnd.ms-excel.sheet.macroEnabled.12"
                    if output_path.suffix.lower() == ".xlsm"
                    else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                self._send_bytes(payload, output_name, content_type)
        except Exception as exc:  # return user-readable errors to the page
            self._send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)


def serve(host: str, port: int) -> None:
    _write_pid()
    httpd = ThreadingHTTPServer((host, port), WebHandler)
    print(f"IDF材料网页工具已启动：http://{host}:{port}")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
        _remove_pid()


def start_server(host: str, port: int, open_browser: bool = True) -> None:
    existing = _read_pid()
    if existing and _pid_is_alive(existing):
        print(f"服务已在运行：http://{host}:{port}（PID {existing}）")
        if open_browser:
            webbrowser.open(f"http://{host}:{port}")
        return
    _remove_pid()
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    log_handle = LOG_FILE.open("a", encoding="utf-8")
    command = [sys.executable, "-u", str(Path(__file__).resolve()), "serve", "--host", host, "--port", str(port)]
    # CREATE_NO_WINDOW keeps the local switch quiet on Windows while the
    # process remains independently stoppable through the PID file.
    creationflags = getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0) | getattr(subprocess, "CREATE_NO_WINDOW", 0)
    subprocess.Popen(
        command,
        stdin=subprocess.DEVNULL,
        stdout=log_handle,
        stderr=log_handle,
        cwd=str(ROOT),
        creationflags=creationflags,
        close_fds=True,
    )
    for _ in range(60):
        time.sleep(0.1)
        pid = _read_pid()
        if pid and _pid_is_alive(pid):
            print(f"已启动：http://{host}:{port}（PID {pid}）")
            if open_browser:
                webbrowser.open(f"http://{host}:{port}")
            return
    raise RuntimeError("本地服务启动失败，请查看 .idf_material_web.log。")


def stop_server() -> None:
    pid = _read_pid()
    if not pid or not _pid_is_alive(pid):
        _remove_pid()
        print("本地服务当前未运行。")
        return
    try:
        if os.name == "nt":
            subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], check=False, capture_output=True)
        else:
            os.kill(pid, signal.SIGTERM)
    finally:
        _remove_pid()
    print(f"已停止本地服务（PID {pid}）。")


def status_server(host: str, port: int) -> None:
    pid = _read_pid()
    alive = bool(pid and _pid_is_alive(pid))
    print({"running": alive, "pid": pid, "url": f"http://{host}:{port}"})


def main() -> None:
    parser = argparse.ArgumentParser(description="IDF材料提取本地网页工具")
    parser.add_argument("action", nargs="?", choices=("start", "serve", "stop", "status"), default="start")
    parser.add_argument("--host", default=DEFAULT_HOST)
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()
    if args.action == "serve":
        serve(args.host, args.port)
    elif args.action == "start":
        start_server(args.host, args.port, open_browser=not args.no_browser)
    elif args.action == "stop":
        stop_server()
    else:
        status_server(args.host, args.port)


if __name__ == "__main__":
    main()
