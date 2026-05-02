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


def level_label(level: int) -> str:
    if level <= 0:
        return "集團本級"
    labels = {
        1: "一級子公司",
        2: "二級子公司",
        3: "三級子公司",
        4: "四級子公司",
    }
    return labels.get(level, f"{level}級子公司")


def parse_level_value(value) -> int | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    match = re.search(r"\d+", text)
    if match:
        return int(match.group(0))
    chinese = {
        "集團本級": 0,
        "頂層主體": 0,
        "一級子公司": 1,
        "二級子公司": 2,
        "三級子公司": 3,
        "四級子公司": 4,
        "五級子公司": 5,
    }
    return chinese.get(text)


def find_row_by_name(master_rows: list[dict], name: str) -> dict | None:
    target = str(name or "").strip()
    if not target:
        return None
    for row in master_rows:
        if target in {str(row.get("canonical_name", "")).strip(), str(row.get("chart1_name", "")).strip()}:
            return row
    return None


def refresh_children_parent_names(master_rows: list[dict], parent_id: str, parent_name: str) -> None:
    for row in master_rows:
        if row.get("chart1_parent") == parent_id:
            row["chart1_parent_name"] = parent_name


def rebuild_task_state(task: dict) -> None:
    from analyzer import _make_graph

    task["summary"] = summary_from_rows(
        task.get("master_rows", []),
        task.get("review_rows", []),
        task.get("candidate_rows", []),
    )
    task["graph"] = _make_graph(task.get("master_rows", []))


def update_review_status(task: dict, key: str, decision: str, note: str = "") -> None:
    for row in task.get("review_rows", []):
        row_key = row.get("candidate_node_id") or row.get("chart2_name")
        if row_key == key:
            row["review_status"] = "done" if decision and decision != "暫不處理" else "pending"
            if note:
                row["review_note"] = note
            break


def apply_chart2_attrs_to_row(target_row: dict, candidate_row: dict) -> None:
    target_row["matched_chart2_name"] = candidate_row.get("chart2_name", "") or candidate_row.get("company", "")
    target_row["legal_representative"] = candidate_row.get("legal_representative", "")
    target_row["established_date"] = candidate_row.get("established_date", "")
    target_row["registered_capital"] = candidate_row.get("registered_capital", "")
    target_row["actual_controller_share"] = candidate_row.get("actual_controller_share", "") or target_row.get("actual_controller_share", "")
    target_row["subsidiary_level_label"] = candidate_row.get("subsidiary_level_label", "") or target_row.get("subsidiary_level_label", "")
    target_row["company_status"] = candidate_row.get("company_status", "")


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
    import threading

    chart1 = request.files.get("chart1")
    chart2 = request.files.get("chart2")
    if not chart1 or not chart2:
        return jsonify({"error": "chart1_and_chart2_required", "message": "請同時上傳圖一和圖二。"}), 400

    if not os.environ.get("DASHSCOPE_API_KEY", "").strip():
        return jsonify({"error": "no_api_key", "message": "伺服器尚未設定 AI 辨識 API Key，無法分析。請聯絡管理員。"}), 422

    task_name = request.form.get("task_name", "").strip()

    # 儲存上傳圖片，立刻建立 processing 狀態的任務
    task_id = uuid.uuid4().hex[:12]
    upload_dir = DATA_DIR / task_id / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)
    c1_path = upload_dir / f"chart1_{sanitize_filename(chart1.filename or 'upload.png')}"
    c2_path = upload_dir / f"chart2_{sanitize_filename(chart2.filename or 'upload.jpg')}"
    chart1.save(str(c1_path))
    chart2.save(str(c2_path))

    c1_name = chart1.filename or ""
    c2_name = chart2.filename or ""

    task: dict = {
        "id": task_id,
        "name": task_name or f"任務-{task_id}",
        "status": "processing",
        "created_at": now_iso(),
        "updated_at": now_iso(),
        "analysis_mode": "qwen_vl",
        "source_files": {"chart1": c1_name, "chart2": c2_name},
        "summary": {},
        "master_rows": [],
        "review_rows": [],
        "candidate_rows": [],
        "review_decisions": {},
        "candidate_decisions": {},
        "graph": {},
        "error": "",
    }
    save_task(task)

    # 背景執行兩段式 Qwen-VL 分析
    def run_async():
        # ── 第一階段：圖一骨架 ──────────────────────────────────
        try:
            from analyzer import run_chart1_stage
            stage1 = run_chart1_stage(c1_path)
            task["status"] = "chart1_ready"
            task["analysis_mode"] = "qwen_vl"
            task["summary"] = stage1["summary"]
            task["master_rows"] = stage1["master_rows"]
            task["review_rows"] = []
            task["candidate_rows"] = []
            task["graph"] = stage1["graph"]
            task["error"] = ""
            save_task(task)
        except Exception as exc:
            task["status"] = "error"
            task["error"] = f"圖一辨識失敗：{exc}"
            save_task(task)
            return  # 圖一失敗就停在這裡

        # ── 第二階段：圖二補充 ──────────────────────────────────
        try:
            from analyzer import enrich_with_chart2
            from pathlib import Path as _Path
            stage2 = enrich_with_chart2(task["master_rows"], c2_path)
            task["status"] = "ready"
            task["summary"] = stage2["summary"]
            task["master_rows"] = stage2["master_rows"]
            task["review_rows"] = stage2["review_rows"]
            task["candidate_rows"] = stage2["candidate_rows"]
            task["graph"] = stage2["graph"]
            task["error"] = ""
        except Exception as exc:
            # 圖二失敗：主表保留圖一骨架，提示用戶重新上傳圖二
            task["status"] = "chart2_error"
            task["error"] = f"圖二辨識失敗：{exc}"
        finally:
            save_task(task)

    threading.Thread(target=run_async, daemon=True).start()

    # 立刻回傳 202，前端輪詢 /api/tasks/<task_id>
    return jsonify({"id": task_id, "status": "processing"}), 202


@app.route("/api/tasks/<task_id>/analyze-chart2", methods=["POST"])
def analyze_chart2_only(task_id: str):
    """單獨重新上傳圖二，保留現有圖一骨架與用戶調整。"""
    import threading

    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404

    chart2 = request.files.get("chart2")
    if not chart2:
        return jsonify({"error": "chart2_required"}), 400

    # 存新的圖二
    upload_dir = DATA_DIR / task_id / "uploads"
    upload_dir.mkdir(parents=True, exist_ok=True)
    c2_path = upload_dir / f"chart2_retry_{sanitize_filename(chart2.filename or 'upload.jpg')}"
    chart2.save(str(c2_path))

    # 標記為處理中，保留現有骨架
    task["status"] = "processing_chart2"
    task["error"] = ""
    task["source_files"]["chart2"] = chart2.filename or ""
    save_task(task)

    existing_master = list(task.get("master_rows", []))

    def run_async():
        try:
            from analyzer import enrich_with_chart2
            stage2 = enrich_with_chart2(existing_master, c2_path)
            task["status"] = "ready"
            task["summary"] = stage2["summary"]
            task["master_rows"] = stage2["master_rows"]
            task["review_rows"] = stage2["review_rows"]
            task["candidate_rows"] = stage2["candidate_rows"]
            task["graph"] = stage2["graph"]
            task["error"] = ""
        except Exception as exc:
            task["status"] = "chart2_error"
            task["error"] = f"圖二辨識失敗：{exc}"
        finally:
            save_task(task)

    threading.Thread(target=run_async, daemon=True).start()
    return jsonify({"id": task_id, "status": "processing_chart2"}), 202


@app.route("/api/review-decision", methods=["POST"])
def review_decision():
    payload = request.get_json(silent=True) or {}
    task = read_task(payload.get("task_id", ""))
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    key = payload.get("key")
    if not key:
        return jsonify({"error": "key_required"}), 400

    decision = payload.get("decision", "")
    corrected_name = payload.get("corrected_name", "").strip()
    corrected_level = payload.get("corrected_level", "")
    corrected_parent = payload.get("corrected_parent", "").strip()
    note = payload.get("note", "")

    task["review_decisions"][key] = {
        "decision": decision,
        "corrected_name": corrected_name,
        "corrected_level": corrected_level,
        "corrected_parent": corrected_parent,
        "note": note,
    }

    target_row = next((row for row in task.get("master_rows", []) if row.get("node_id") == key), None)
    if target_row:
        matched_candidate = next(
            (
                row for row in task.get("candidate_rows", [])
                if row.get("chart2_name") and row.get("chart2_name") == target_row.get("matched_chart2_name")
            ),
            None,
        )
        if corrected_name:
            target_row["canonical_name"] = corrected_name
            refresh_children_parent_names(task["master_rows"], target_row["node_id"], corrected_name)

        parent_row = find_row_by_name(task.get("master_rows", []), corrected_parent) if corrected_parent else None
        if corrected_parent:
            target_row["chart1_parent"] = parent_row.get("node_id", "") if parent_row else ""
            target_row["chart1_parent_name"] = parent_row.get("canonical_name") if parent_row else corrected_parent
        parsed_level = parse_level_value(corrected_level)
        if parsed_level is not None:
            target_row["chart1_level"] = parsed_level
            target_row["subsidiary_level_label"] = level_label(parsed_level)
        elif parent_row:
            parent_level = parse_level_value(parent_row.get("chart1_level")) or 0
            target_row["chart1_level"] = parent_level + 1
            target_row["subsidiary_level_label"] = level_label(parent_level + 1)

        if decision == "確認一致":
            if matched_candidate:
                apply_chart2_attrs_to_row(target_row, matched_candidate)
                task["candidate_rows"] = [
                    row for row in task.get("candidate_rows", [])
                    if row.get("chart2_name") != matched_candidate.get("chart2_name")
                ]
            target_row["node_status"] = "enriched" if target_row.get("matched_chart2_name") else target_row.get("node_status", "review_match")
            target_row["review_flag"] = ""
        elif decision == "不是同一家公司":
            target_row["matched_chart2_name"] = ""
            target_row["node_status"] = "chart1_only"
            target_row["review_flag"] = "yes"
        if note:
            target_row["review_note"] = note

    update_review_status(task, key, decision, note)
    rebuild_task_state(task)
    save_task(task)
    return jsonify({
        "ok": True,
        "review_decisions": task["review_decisions"],
        "master_rows": task["master_rows"],
        "review_rows": task["review_rows"],
        "candidate_rows": task["candidate_rows"],
        "summary": task["summary"],
        "graph": task["graph"],
    })


@app.route("/api/tasks/<task_id>/update-row", methods=["POST"])
def update_row(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    payload = request.get_json(silent=True) or {}

    editable = ["canonical_name", "legal_representative", "registered_capital",
                "established_date", "actual_controller_share", "company_status",
                "chart1_parent_name", "chart1_parent", "chart1_level",
                "subsidiary_level_label"]

    # 連動更新模式：同欄位相同原始值的列全部更新
    if payload.get("cascade") and payload.get("field") and "original_value" in payload:
        field = payload["field"]
        original = payload["original_value"]
        new_val  = payload.get("new_value", "")
        if field in editable:
            for row in task.get("master_rows", []):
                if row.get(field) == original:
                    row[field] = new_val
                    if field == "canonical_name":
                        refresh_children_parent_names(task["master_rows"], row["node_id"], new_val)
    else:
        node_id = payload.get("node_id")
        if not node_id:
            return jsonify({"error": "node_id_required"}), 400
        for row in task.get("master_rows", []):
            if row.get("node_id") == node_id:
                for field in editable:
                    if field in payload:
                        row[field] = payload[field]
                        if field == "canonical_name":
                            refresh_children_parent_names(task["master_rows"], row["node_id"], payload[field])
                break

    rebuild_task_state(task)
    save_task(task)
    return jsonify({
        "ok": True,
        "master_rows": task["master_rows"],
        "review_rows": task["review_rows"],
        "candidate_rows": task["candidate_rows"],
        "summary": task["summary"],
        "graph": task["graph"],
    })


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
    for row in task["master_rows"]:
        if row.get("chart1_parent") == node_id:
            row["chart1_parent"] = ""
            row["chart1_parent_name"] = ""
            row["chart1_level"] = 0
            row["subsidiary_level_label"] = level_label(0)
    rebuild_task_state(task)
    save_task(task)
    return jsonify({
        "ok": True,
        "master_rows": task["master_rows"],
        "review_rows": task["review_rows"],
        "candidate_rows": task["candidate_rows"],
        "summary": task["summary"],
        "graph": task["graph"],
    })


@app.route("/api/candidate-decision", methods=["POST"])
def candidate_decision():
    payload = request.get_json(silent=True) or {}
    task = read_task(payload.get("task_id", ""))
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    key = payload.get("key")
    if not key:
        return jsonify({"error": "key_required"}), 400

    decision = payload.get("decision", "")
    parent_name = payload.get("parent", "").strip()
    corrected_name = payload.get("corrected_name", "").strip()
    note = payload.get("note", "")

    task["candidate_decisions"][key] = {
        "decision": decision,
        "parent": parent_name,
        "corrected_name": corrected_name,
        "note": note,
    }

    if decision == "加入主表":
        candidate = next((row for row in task.get("candidate_rows", []) if row.get("chart2_name") == key), None)
        if candidate:
            parent_row = find_row_by_name(task.get("master_rows", []), parent_name) if parent_name else None
            parent_level = parse_level_value(parent_row.get("chart1_level")) if parent_row else None
            level = (parent_level + 1) if parent_level is not None else parse_level_value(candidate.get("subsidiary_level_label")) or 0
            final_name = corrected_name or candidate.get("chart2_name") or candidate.get("company") or ""
            new_row = {
                "node_id": f"A{uuid.uuid4().hex[:6].upper()}",
                "chart1_name": final_name,
                "canonical_name": final_name,
                "chart1_level": level,
                "chart1_parent": parent_row.get("node_id", "") if parent_row else "",
                "chart1_parent_name": parent_row.get("canonical_name", "") if parent_row else parent_name,
                "matched_chart2_name": candidate.get("chart2_name", ""),
                "legal_representative": candidate.get("legal_representative", ""),
                "established_date": candidate.get("established_date", ""),
                "registered_capital": candidate.get("registered_capital", ""),
                "actual_controller_share": candidate.get("actual_controller_share", ""),
                "subsidiary_level_label": candidate.get("subsidiary_level_label") or level_label(level),
                "company_status": candidate.get("company_status", ""),
                "match_status": "matched",
                "node_status": "enriched",
                "review_flag": "",
                "review_note": note,
            }
            task["master_rows"].append(new_row)
            task["candidate_rows"] = [row for row in task.get("candidate_rows", []) if row.get("chart2_name") != key]

    rebuild_task_state(task)
    save_task(task)
    return jsonify({
        "ok": True,
        "candidate_decisions": task["candidate_decisions"],
        "master_rows": task["master_rows"],
        "review_rows": task["review_rows"],
        "candidate_rows": task["candidate_rows"],
        "summary": task["summary"],
        "graph": task["graph"],
    })


# ── entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
