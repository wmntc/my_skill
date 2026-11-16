# -*- coding: utf-8 -*-
"""
本地查询服务（零依赖，仅用 Python 标准库 http.server）

提供两类能力：
  1. 静态托管 web/ 目录下的查询页面（index.html）；
  2. /api/* JSON 接口，从 SQLite 读取分析结果，供页面做班级/校区对比与可视化。

用法（由 score_analysis.py --serve 调用，也可单独运行）：
    python web_app.py --db data/analysis.db --port 8765
"""

import os
import json
import urllib.parse
from http.server import HTTPServer, SimpleHTTPRequestHandler

import db_store

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
WEB_DIR = os.path.join(SKILL_DIR, "web")
DEFAULT_DB = os.path.join(SKILL_DIR, "data", "analysis.db")


class QueryHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, db_path=DEFAULT_DB, **kwargs):
        self.db_path = db_path
        super().__init__(*args, **kwargs)

    def _send_json(self, obj, status=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def end_headers(self):
        # 禁止浏览器缓存页面与接口，确保修正后的 index.html / API 始终被重新拉取，
        # 避免用户看到旧版本（例如校区对比图表曾因缓存显示 undefined 校区名）。
        self.send_header("Cache-Control", "no-store, max-age=0")
        super().end_headers()

    def _api(self, path, query):
        if path == "/api/exams":
            return db_store.list_exams(self.db_path)
        if path == "/api/subjects":
            return db_store.get_subjects(self.db_path, int(query.get("exam_id", [0])[0]))
        if path == "/api/classes":
            subject = query.get("subject", [None])[0]
            return db_store.get_classes(self.db_path, int(query.get("exam_id", [0])[0]), subject)
        if path == "/api/class_metrics":
            subject = query.get("subject", [None])[0]
            return db_store.get_class_metrics(self.db_path, int(query.get("exam_id", [0])[0]), subject)
        if path == "/api/teacher_scores":
            return db_store.get_teacher_scores(self.db_path, int(query.get("exam_id", [0])[0]))
        if path == "/api/campus":
            subject = query.get("subject", [None])[0]
            return db_store.get_campus_agg(self.db_path, int(query.get("exam_id", [0])[0]), subject)
        if path == "/api/compare":
            subject = query.get("subject", [""])[0]
            classes = [c for c in query.get("classes", [""])[0].split(",") if c]
            exam_ids = [int(x) for x in query.get("exam_ids", [""])[0].split(",") if x]
            return db_store.compare_classes(self.db_path, classes, subject, exam_ids)
        return {"error": "unknown api: " + path}, 404

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path.startswith("/api/"):
            query = urllib.parse.parse_qs(parsed.query)
            try:
                resp = self._api(parsed.path, query)
            except Exception as e:  # noqa: BLE001
                self._send_json({"error": str(e)}, status=500)
                return
            # _api 成功时返回数据本身（list/dict），失败时返回 (data, status)
            if isinstance(resp, tuple) and len(resp) == 2 and isinstance(resp[1], int):
                data, status = resp
            else:
                data, status = resp, 200
            self._send_json(data, status=status)
            return
        # 静态文件：默认首页
        if parsed.path in ("/", ""):
            self.path = "/index.html"
        return SimpleHTTPRequestHandler.do_GET(self)

    def translate_path(self, path):
        # 把请求映射到 WEB_DIR 而不是当前工作目录
        rel = urllib.parse.urlparse(self.path).path.lstrip("/")
        rel = urllib.parse.unquote(rel)
        return os.path.join(WEB_DIR, rel)

    def log_message(self, fmt, *args):
        pass  # 静默日志，避免刷屏

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path == "/api/exam_delete":
            try:
                length = int(self.headers.get("Content-Length", 0))
                body = json.loads(self.rfile.read(length) or b"{}")
                exam_id = int(body.get("exam_id"))
                db_store.delete_exam(self.db_path, exam_id)
                self._send_json({"ok": True, "exam_id": exam_id})
            except Exception as e:  # noqa: BLE001
                self._send_json({"ok": False, "error": str(e)}, status=500)
            return
        self._send_json({"error": "unknown api: " + parsed.path}, status=404)


def run_server(db_path: str = DEFAULT_DB, port: int = 8765, host: str = "127.0.0.1"):
    db_store.init_db(db_path)
    handler = lambda *a, **k: QueryHandler(*a, db_path=db_path, **k)
    server = HTTPServer((host, port), handler)
    url = f"http://{host}:{port}/"
    print(f"本地查询服务已启动：{url}")
    print(f"数据库：{db_path}")
    print("按 Ctrl+C 停止服务。")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务已停止。")
    finally:
        server.server_close()


if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="呼市实验中学成绩分析 - 本地查询服务")
    ap.add_argument("--db", default=DEFAULT_DB, help="SQLite 数据库路径")
    ap.add_argument("--port", type=int, default=8765, help="服务端口，默认 8765")
    ap.add_argument("--host", default="127.0.0.1", help="监听地址，默认 127.0.0.1")
    a = ap.parse_args()
    run_server(a.db, a.port, a.host)
