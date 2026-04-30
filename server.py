from __future__ import annotations

import cgi
import csv
import json
import os
import re
import shutil
import uuid
from datetime import datetime, timezone
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "app_data" / "tasks"
WEB_DIR = BASE_DIR / "webapp"


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def parse_csv(path: Path) -> list[dict[str, str]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def summary_from_payload(master_rows, review_rows, candidate_rows):
    return {
        "master_count": len(master_rows),
        "enriched_count": sum(1 for row in master_rows if row.get("node_status") == "enriched"),
        "review_count": len(review_rows),
        "chart1_only_count": sum(1 for row in master_rows if row.get("node_status") == "chart1_only"),
        "candidate_count": len(candidate_rows),
    }


def load_sample_payload() -> dict:
    master_rows = parse_csv(BASE_DIR / "reconciliation_outputs" / "master_nodes_enriched.csv")
    review_rows = parse_csv(BASE_DIR / "reconciliation_outputs" / "reconciliation_report.csv")
    candidate_rows = parse_csv(BASE_DIR / "reconciliation_outputs" / "chart2_only_candidates.csv")
    graph_nodes = parse_csv(BASE_DIR / "qcc_nodes.csv")
    graph_edges = parse_csv(BASE_DIR / "qcc_edges.csv")
    return {
        "master_rows": master_rows,
        "review_rows": review_rows,
        "candidate_rows": candidate_rows,
        "graph": {
            "nodes": graph_nodes,
            "edges": graph_edges,
            "stage2": {
                "status": "reserved",
                "ready_after_review": True,
                "target_output": "equity_structure_chart",
                "note": "Second stage should consume resolved master rows plus accepted candidate nodes to generate a final equity structure chart.",
            },
        },
        "summary": summary_from_payload(master_rows, review_rows, candidate_rows),
    }


def task_dir(task_id: str) -> Path:
    return DATA_DIR / task_id


def task_json_path(task_id: str) -> Path:
    return task_dir(task_id) / "task.json"


def write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def read_task(task_id: str) -> dict | None:
    path = task_json_path(task_id)
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def save_task(task: dict) -> None:
    task["updated_at"] = now_iso()
    write_json(task_json_path(task["id"]), task)


def sanitize_filename(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9._\-\u4e00-\u9fff()（）]+", "_", name.strip())
    return name or "upload.bin"


def build_task_payload(task_name: str, chart1_file: str, chart2_file: str) -> dict:
    sample = load_sample_payload()
    task_id = uuid.uuid4().hex[:12]
    return {
        "id": task_id,
        "name": task_name or f"任務-{task_id}",
        "status": "ready",
        "created_at": now_iso(),
        "updated_at": now_iso(),
        "analysis_mode": "sample_seed",
        "source_files": {
            "chart1": chart1_file,
            "chart2": chart2_file,
        },
        "summary": sample["summary"],
        "master_rows": sample["master_rows"],
        "review_rows": sample["review_rows"],
        "candidate_rows": sample["candidate_rows"],
        "review_decisions": {},
        "candidate_decisions": {},
        "graph": sample["graph"],
    }


def build_task_payload_from_analysis(task_name: str, chart1_file: str, chart2_file: str, analysis: dict) -> dict:
    task_id = uuid.uuid4().hex[:12]
    return {
        "id": task_id,
        "name": task_name or f"任務-{task_id}",
        "status": "ready",
        "created_at": now_iso(),
        "updated_at": now_iso(),
        "analysis_mode": "qwen_vl",
        "source_files": {
            "chart1": chart1_file,
            "chart2": chart2_file,
        },
        "summary": analysis["summary"],
        "master_rows": analysis["master_rows"],
        "review_rows": analysis["review_rows"],
        "candidate_rows": analysis["candidate_rows"],
        "review_decisions": {},
        "candidate_decisions": {},
        "graph": analysis["graph"],
    }


class EquityReviewHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def log_message(self, format, *args):  # noqa: A003
        return super().log_message(format, *args)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.path = "/webapp/index.html"
            return super().do_GET()

        if parsed.path == "/api/health":
            return self.respond_json({"ok": True, "time": now_iso()})

        if parsed.path == "/api/demo-task":
            task = build_task_payload("示範任務", "demo_chart1.png", "demo_chart2.jpg")
            save_task(task)
            return self.respond_json(task, HTTPStatus.CREATED)

        if parsed.path.startswith("/api/tasks/"):
            task_id = parsed.path.split("/")[3] if len(parsed.path.split("/")) > 3 else ""
            task = read_task(task_id)
            if not task:
                return self.respond_json({"error": "task_not_found"}, HTTPStatus.NOT_FOUND)
            return self.respond_json(task)

        return super().do_GET()

    def do_POST(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/tasks/analyze":
            return self.handle_task_analyze()

        if parsed.path == "/api/review-decision":
            return self.handle_review_decision()

        if parsed.path == "/api/candidate-decision":
            return self.handle_candidate_decision()

        return self.respond_json({"error": "not_found"}, HTTPStatus.NOT_FOUND)

    def handle_task_analyze(self):
        ctype, _ = cgi.parse_header(self.headers.get("content-type", ""))
        if ctype != "multipart/form-data":
            return self.respond_json({"error": "multipart_required"}, HTTPStatus.BAD_REQUEST)

        environ = {
            "REQUEST_METHOD": "POST",
            "CONTENT_TYPE": self.headers.get("content-type"),
        }
        form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ=environ)
        task_name = form.getfirst("task_name", "").strip()
        chart1 = form["chart1"] if "chart1" in form else None
        chart2 = form["chart2"] if "chart2" in form else None
        if not chart1 or not chart2 or not getattr(chart1, "file", None) or not getattr(chart2, "file", None):
            return self.respond_json({"error": "chart1_and_chart2_required"}, HTTPStatus.BAD_REQUEST)

        # 先建立暫存資料夾，存放上傳圖片
        tmp_id = uuid.uuid4().hex[:12]
        folder = DATA_DIR / tmp_id / "uploads"
        folder.mkdir(parents=True, exist_ok=True)

        saved_paths: dict[str, Path] = {}
        for field_name, field in (("chart1", chart1), ("chart2", chart2)):
            filename = sanitize_filename(field.filename)
            dest = folder / f"{field_name}_{filename}"
            with dest.open("wb") as handle:
                shutil.copyfileobj(field.file, handle)
            saved_paths[field_name] = dest

        # 若有設定 DASHSCOPE_API_KEY，使用 Qwen-VL 真實分析；否則 fallback 示範資料
        use_qwen = bool(os.environ.get("DASHSCOPE_API_KEY", "").strip())
        if use_qwen:
            try:
                from analyzer import run_analysis
                analysis = run_analysis(saved_paths["chart1"], saved_paths["chart2"])
                task = build_task_payload_from_analysis(task_name, chart1.filename, chart2.filename, analysis)
            except Exception as exc:
                task = build_task_payload(task_name, chart1.filename, chart2.filename)
                task["analysis_warning"] = f"Qwen-VL 分析失敗，已 fallback 示範資料：{exc}"
        else:
            task = build_task_payload(task_name, chart1.filename, chart2.filename)
            task["analysis_warning"] = "未設定 DASHSCOPE_API_KEY，目前使用示範資料。"

        # 把暫存資料夾重命名為正式 task_id
        (DATA_DIR / tmp_id).rename(task_dir(task["id"]))
        save_task(task)
        return self.respond_json(task, HTTPStatus.CREATED)

    def handle_review_decision(self):
        payload = self.read_json_body()
        if not payload:
            return self.respond_json({"error": "invalid_json"}, HTTPStatus.BAD_REQUEST)
        task = read_task(payload.get("task_id", ""))
        if not task:
            return self.respond_json({"error": "task_not_found"}, HTTPStatus.NOT_FOUND)
        key = payload.get("key")
        if not key:
            return self.respond_json({"error": "key_required"}, HTTPStatus.BAD_REQUEST)

        task["review_decisions"][key] = {
            "decision": payload.get("decision", ""),
            "corrected_name": payload.get("corrected_name", ""),
            "corrected_level": payload.get("corrected_level", ""),
            "corrected_parent": payload.get("corrected_parent", ""),
            "note": payload.get("note", ""),
        }
        save_task(task)
        return self.respond_json({"ok": True, "review_decisions": task["review_decisions"]})

    def handle_candidate_decision(self):
        payload = self.read_json_body()
        if not payload:
            return self.respond_json({"error": "invalid_json"}, HTTPStatus.BAD_REQUEST)
        task = read_task(payload.get("task_id", ""))
        if not task:
            return self.respond_json({"error": "task_not_found"}, HTTPStatus.NOT_FOUND)
        key = payload.get("key")
        if not key:
            return self.respond_json({"error": "key_required"}, HTTPStatus.BAD_REQUEST)

        task["candidate_decisions"][key] = {
            "decision": payload.get("decision", ""),
            "parent": payload.get("parent", ""),
            "corrected_name": payload.get("corrected_name", ""),
            "note": payload.get("note", ""),
        }
        save_task(task)
        return self.respond_json({"ok": True, "candidate_decisions": task["candidate_decisions"]})

    def read_json_body(self) -> dict | None:
        length = int(self.headers.get("content-length", "0") or 0)
        if length <= 0:
            return None
        body = self.rfile.read(length)
        try:
            return json.loads(body.decode("utf-8"))
        except Exception:
            return None

    def do_OPTIONS(self):
        self.send_response(HTTPStatus.NO_CONTENT)
        self._send_cors_headers()
        self.end_headers()

    def _send_cors_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def respond_json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self._send_cors_headers()
        self.end_headers()
        self.wfile.write(body)


def run(host: str = "0.0.0.0", port: int = 8000):
    port = int(os.environ.get("PORT", port))
    ensure_data_dir()
    server = ThreadingHTTPServer((host, port), EquityReviewHandler)
    print(f"Serving equity review app on http://{host}:{port}")
    server.serve_forever()


if __name__ == "__main__":
    run()
