from __future__ import annotations

import csv
import json
import os
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "app_data" / "tasks"
WEB_DIR = BASE_DIR / "webapp"

app = Flask(__name__, static_folder=str(WEB_DIR), static_url_path="")
CORS(app)


# ── helpers ──────────────────────────────────────────────────────────────────

def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def parse_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def sanitize_filename(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9._\-一-鿿()（）]+", "_", name.strip())
    return name or "upload.bin"


def write_json(path: Path, payload: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def read_task(task_id: str) -> dict | None:
    path = DATA_DIR / task_id / "task.json"
    if not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def save_task(task: dict) -> None:
    task["updated_at"] = now_iso()
    write_json(DATA_DIR / task["id"] / "task.json", task)


def summary_from_rows(master_rows, review_rows, candidate_rows) -> dict:
    return {
        "master_count": len(master_rows),
        "enriched_count": sum(1 for r in master_rows if r.get("node_status") == "enriched"),
        "review_count": len(review_rows),
        "chart1_only_count": sum(1 for r in master_rows if r.get("node_status") == "chart1_only"),
        "candidate_count": len(candidate_rows),
    }


def load_sample_payload() -> dict:
    out_dir = BASE_DIR / "reconciliation_outputs"
    master_rows = parse_csv(out_dir / "master_nodes_enriched.csv")
    review_rows = parse_csv(out_dir / "reconciliation_report.csv")
    candidate_rows = parse_csv(out_dir / "chart2_only_candidates.csv")
    nodes_path = BASE_DIR / "qcc_nodes.csv"
    edges_path = BASE_DIR / "qcc_edges.csv"
    return {
        "master_rows": master_rows,
        "review_rows": review_rows,
        "candidate_rows": candidate_rows,
        "graph": {
            "nodes": parse_csv(nodes_path) if nodes_path.exists() else [],
            "edges": parse_csv(edges_path) if edges_path.exists() else [],
            "stage2": {
                "status": "reserved",
                "ready_after_review": True,
                "target_output": "equity_structure_chart",
                "note": "第二階段：審核完成後生成最終股權架構圖。",
            },
        },
        "summary": summary_from_rows(master_rows, review_rows, candidate_rows),
    }


def build_task(task_name: str, chart1_file: str, chart2_file: str) -> dict:
    sample = load_sample_payload()
    task_id = uuid.uuid4().hex[:12]
    return {
        "id": task_id,
        "name": task_name or f"任務-{task_id}",
        "status": "ready",
        "created_at": now_iso(),
        "updated_at": now_iso(),
        "analysis_mode": "sample_seed",
        "source_files": {"chart1": chart1_file, "chart2": chart2_file},
        "summary": sample["summary"],
        "master_rows": sample["master_rows"],
        "review_rows": sample["review_rows"],
        "candidate_rows": sample["candidate_rows"],
        "review_decisions": {},
        "candidate_decisions": {},
        "graph": sample["graph"],
    }


def build_task_from_analysis(task_name: str, chart1_file: str, chart2_file: str, analysis: dict) -> dict:
    task_id = uuid.uuid4().hex[:12]
    return {
        "id": task_id,
        "name": task_name or f"任務-{task_id}",
        "status": "ready",
        "created_at": now_iso(),
        "updated_at": now_iso(),
        "analysis_mode": "qwen_vl",
        "source_files": {"chart1": chart1_file, "chart2": chart2_file},
        "summary": analysis["summary"],
        "master_rows": analysis["master_rows"],
        "review_rows": analysis["review_rows"],
        "candidate_rows": analysis["candidate_rows"],
        "review_decisions": {},
        "candidate_decisions": {},
        "graph": analysis["graph"],
    }


# ── routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory(str(WEB_DIR), "index.html")


@app.route("/<path:filename>")
def static_files(filename):
    return send_from_directory(str(WEB_DIR), filename)


@app.route("/api/health")
def health():
    return jsonify({"ok": True, "time": now_iso()})


@app.route("/api/demo-task")
def demo_task():
    task = build_task("示範任務", "demo_chart1.png", "demo_chart2.jpg")
    save_task(task)
    return jsonify(task), 201


@app.route("/api/tasks/<task_id>")
def get_task(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    return jsonify(task)


@app.route("/api/tasks/analyze", methods=["POST"])
def analyze():
    import shutil

    chart1 = request.files.get("chart1")
    chart2 = request.files.get("chart2")
    if not chart1 or not chart2:
        return jsonify({"error": "chart1_and_chart2_required", "message": "請同時上傳圖一和圖二。"}), 400

    if not os.environ.get("DASHSCOPE_API_KEY", "").strip():
        return jsonify({"error": "no_api_key", "message": "伺服器尚未設定 AI 辨識 API Key，無法分析。請聯絡管理員。"}), 422

    task_name = request.form.get("task_name", "").strip()

    # 儲存上傳圖片
    tmp_id = uuid.uuid4().hex[:12]
    upload_dir = DATA_DIR / tmp_id / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)
    c1_path = upload_dir / f"chart1_{sanitize_filename(chart1.filename or 'upload.png')}"
    c2_path = upload_dir / f"chart2_{sanitize_filename(chart2.filename or 'upload.jpg')}"
    chart1.save(str(c1_path))
    chart2.save(str(c2_path))

    # Qwen-VL 分析（失敗直接回傳錯誤，不 fallback）
    try:
        from analyzer import run_analysis
        analysis = run_analysis(c1_path, c2_path)
        task = build_task_from_analysis(task_name, chart1.filename or "", chart2.filename or "", analysis)
    except Exception as exc:
        shutil.rmtree(DATA_DIR / tmp_id, ignore_errors=True)
        return jsonify({"error": "analysis_failed", "message": f"AI 辨識失敗，請重新上傳或確認圖片是否清晰：{exc}"}), 500

    (DATA_DIR / tmp_id).rename(DATA_DIR / task["id"])
    save_task(task)
    return jsonify(task), 201


@app.route("/api/review-decision", methods=["POST"])
def review_decision():
    payload = request.get_json(silent=True) or {}
    task = read_task(payload.get("task_id", ""))
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    key = payload.get("key")
    if not key:
        return jsonify({"error": "key_required"}), 400
    task["review_decisions"][key] = {
        "decision": payload.get("decision", ""),
        "corrected_name": payload.get("corrected_name", ""),
        "corrected_level": payload.get("corrected_level", ""),
        "corrected_parent": payload.get("corrected_parent", ""),
        "note": payload.get("note", ""),
    }
    save_task(task)
    return jsonify({"ok": True, "review_decisions": task["review_decisions"]})


@app.route("/api/tasks/<task_id>/update-row", methods=["POST"])
def update_row(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    payload = request.get_json(silent=True) or {}

    editable = ["canonical_name", "legal_representative", "registered_capital",
                "established_date", "actual_controller_share", "company_status",
                "chart1_parent_name", "subsidiary_level_label"]

    # 連動更新模式：同欄位相同原始值的列全部更新
    if payload.get("cascade") and payload.get("field") and "original_value" in payload:
        field = payload["field"]
        original = payload["original_value"]
        new_val  = payload.get("new_value", "")
        if field in editable:
            for row in task.get("master_rows", []):
                if row.get(field) == original:
                    row[field] = new_val
    else:
        node_id = payload.get("node_id")
        if not node_id:
            return jsonify({"error": "node_id_required"}), 400
        for row in task.get("master_rows", []):
            if row.get("node_id") == node_id:
                for field in editable:
                    if field in payload:
                        row[field] = payload[field]
                break

    save_task(task)
    return jsonify({"ok": True, "master_rows": task["master_rows"]})


@app.route("/api/tasks/<task_id>/delete-row", methods=["POST"])
def delete_row(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    payload = request.get_json(silent=True) or {}
    node_id = payload.get("node_id")
    if not node_id:
        return jsonify({"error": "node_id_required"}), 400
    task["master_rows"] = [r for r in task.get("master_rows", []) if r.get("node_id") != node_id]
    save_task(task)
    return jsonify({"ok": True, "master_rows": task["master_rows"]})


@app.route("/api/candidate-decision", methods=["POST"])
def candidate_decision():
    payload = request.get_json(silent=True) or {}
    task = read_task(payload.get("task_id", ""))
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    key = payload.get("key")
    if not key:
        return jsonify({"error": "key_required"}), 400
    task["candidate_decisions"][key] = {
        "decision": payload.get("decision", ""),
        "parent": payload.get("parent", ""),
        "corrected_name": payload.get("corrected_name", ""),
        "note": payload.get("note", ""),
    }
    save_task(task)
    return jsonify({"ok": True, "candidate_decisions": task["candidate_decisions"]})


# ── entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
