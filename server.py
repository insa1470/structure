from __future__ import annotations

import csv
import copy
import os
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path
from tempfile import NamedTemporaryFile

from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS

from storage import ensure_storage_ready, list_ocr_tests, list_tasks, read_task, save_ocr_test, save_task, task_upload_dir

BASE_DIR = Path(__file__).resolve().parent
WEB_DIR = BASE_DIR / "webapp"
MAX_CHART1_MB = 3
MAX_CHART1_LONG_EDGE = 9000
MAX_CHART2_CHUNKS = 9

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


def _estimate_chart2_chunks(width: int, height: int) -> int:
    scaled_height = round(height * (900 / width)) if width > 900 else height
    if scaled_height <= 3600:
        return 1
    return ((scaled_height - 1500) + (1500 - 180) - 1) // (1500 - 180) + 1


def _inspect_image(file_storage):
    try:
        from PIL import Image
    except ImportError:
        return None
    stream = file_storage.stream
    pos = stream.tell()
    stream.seek(0)
    try:
        img = Image.open(stream)
        width, height = img.size
        return {"width": width, "height": height}
    finally:
        stream.seek(pos)


def _file_size_mb(file_storage) -> float:
    stream = file_storage.stream
    pos = stream.tell()
    stream.seek(0, 2)
    size = stream.tell()
    stream.seek(pos)
    return size / 1024 / 1024


def _compact_task(task: dict) -> dict:
    summary = task.get("summary") or {}
    return {
        "id": task.get("id"),
        "name": task.get("name"),
        "status": task.get("status"),
        "created_at": task.get("created_at"),
        "updated_at": task.get("updated_at"),
        "source_files": task.get("source_files") or {},
        "master_count": summary.get("master_count", len(task.get("master_rows", []) or [])),
        "review_count": summary.get("review_count", len(task.get("review_rows", []) or [])),
        "candidate_count": summary.get("candidate_count", len(task.get("candidate_rows", []) or [])),
    }


def _normalize_company_text(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text)
    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("，", ",").replace("：", ":")
    return text


def is_admin_request() -> bool:
    password = os.environ.get("ADMIN_TEST_PASSWORD", "").strip()
    if not password:
        return True
    supplied = (
        request.headers.get("X-Admin-Test-Password")
        or request.form.get("admin_password")
        or request.args.get("admin_password")
        or ""
    ).strip()
    return supplied == password


def admin_required_response():
    return jsonify({"error": "admin_required", "message": "請先輸入管理測試密碼。"}), 403


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


def _draft_payload_from_task(task: dict) -> dict:
    return {
        "master_rows": copy.deepcopy(task.get("master_rows", [])),
        "review_rows": copy.deepcopy(task.get("review_rows", [])),
        "candidate_rows": copy.deepcopy(task.get("candidate_rows", [])),
        "review_decisions": copy.deepcopy(task.get("review_decisions", {})),
        "candidate_decisions": copy.deepcopy(task.get("candidate_decisions", {})),
    }


def make_chart2_progress_saver(task: dict):
    def save_progress(progress: dict) -> None:
        task["status"] = "processing_chart2"
        task["chart2_progress"] = {
            **task.get("chart2_progress", {}),
            **progress,
            "updated_at": now_iso(),
        }
        save_task(task)

    return save_progress


def is_task_cancel_requested(task_id: str) -> bool:
    latest = read_task(task_id)
    return bool(latest and latest.get("cancel_requested"))


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


@app.route("/api/ocr/probe", methods=["POST"])
def ocr_probe():
    started_at = datetime.now(timezone.utc)
    image = request.files.get("image") or request.files.get("chart1") or request.files.get("chart2")
    if not image:
        return jsonify({"error": "image_required", "message": "請上傳 image、chart1 或 chart2 圖片。"}), 400

    provider = (request.form.get("provider") or request.args.get("provider") or "").strip() or None
    save_record = (request.form.get("save") or request.args.get("save") or "").lower() in {"1", "true", "yes"}
    if save_record and not is_admin_request():
        return admin_required_response()

    original_filename = image.filename or "upload.png"
    note = (request.form.get("note") or "").strip()
    suffix = Path(image.filename or "upload.png").suffix or ".png"
    temp_path = None
    try:
        with NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            image.save(tmp.name)
            temp_path = Path(tmp.name)
        from ocr_engine import run_ocr_probe
        result = run_ocr_probe(temp_path, provider)
        elapsed_ms = int((datetime.now(timezone.utc) - started_at).total_seconds() * 1000)
        if save_record:
            record = {
                "id": f"T{uuid.uuid4().hex[:10].upper()}",
                "created_at": now_iso(),
                "provider": result.get("provider") or provider or "",
                "model": result.get("model") or "",
                "filename": original_filename,
                "file_size": temp_path.stat().st_size if temp_path and temp_path.exists() else 0,
                "elapsed_ms": elapsed_ms,
                "text_count": result.get("text_count", 0),
                "company_candidate_count": result.get("company_candidate_count", 0),
                "items": result.get("items", [])[:120],
                "company_candidates": result.get("company_candidates", [])[:120],
                "note": note,
                "raw_response": result.get("raw_response") or result.get("raw_text") or "",
            }
            save_ocr_test(record)
            result["saved_test_id"] = record["id"]
        result["elapsed_ms"] = elapsed_ms
        return jsonify({"ok": True, **result})
    except RuntimeError as exc:
        return jsonify({"error": "ocr_unavailable", "message": str(exc)}), 503
    except Exception as exc:
        return jsonify({"error": "ocr_failed", "message": str(exc)}), 500
    finally:
        if temp_path and temp_path.exists():
            temp_path.unlink(missing_ok=True)


@app.route("/api/ocr/tests")
def ocr_tests():
    if not is_admin_request():
        return admin_required_response()
    limit = request.args.get("limit", "50")
    try:
        limit_value = max(1, min(int(limit), 200))
    except ValueError:
        limit_value = 50
    return jsonify({"ok": True, "tests": list_ocr_tests(limit_value)})


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


@app.route("/api/tasks/<task_id>/save-draft", methods=["POST"])
def save_task_draft(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    payload = request.get_json(silent=True) or {}
    draft_state = payload.get("state") if isinstance(payload, dict) else None
    if not isinstance(draft_state, dict):
        draft_state = _draft_payload_from_task(task)
    task["draft"] = {
        "saved_at": now_iso(),
        "state": draft_state,
    }
    try:
        save_task(task)
    except Exception as exc:
        return jsonify({
            "error": "draft_save_failed",
            "message": f"暫存寫入失敗：{exc}",
        }), 500
    return jsonify({"ok": True, "saved_at": task["draft"]["saved_at"]})


@app.route("/api/tasks/<task_id>/restore-draft", methods=["POST"])
def restore_task_draft(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    draft = task.get("draft") or {}
    state = draft.get("state") if isinstance(draft, dict) else None
    if not isinstance(state, dict):
        return jsonify({"error": "draft_not_found", "message": "目前沒有可還原的草稿。"}), 404
    if isinstance(state.get("master_rows"), list):
        task["master_rows"] = state["master_rows"]
    if isinstance(state.get("review_rows"), list):
        task["review_rows"] = state["review_rows"]
    if isinstance(state.get("candidate_rows"), list):
        task["candidate_rows"] = state["candidate_rows"]
    if isinstance(state.get("review_decisions"), dict):
        task["review_decisions"] = state["review_decisions"]
    if isinstance(state.get("candidate_decisions"), dict):
        task["candidate_decisions"] = state["candidate_decisions"]
    rebuild_task_state(task)
    save_task(task)
    return jsonify({
        "ok": True,
        "master_rows": task["master_rows"],
        "review_rows": task["review_rows"],
        "candidate_rows": task["candidate_rows"],
        "review_decisions": task["review_decisions"],
        "candidate_decisions": task["candidate_decisions"],
        "summary": task["summary"],
        "graph": task["graph"],
    })


@app.route("/api/tasks")
def query_tasks():
    keyword = (request.args.get("q") or "").strip().lower()
    status = (request.args.get("status") or "").strip().lower()
    try:
        limit = max(1, min(int(request.args.get("limit", 200) or 200), 500))
    except ValueError:
        limit = 200
    items = [_compact_task(task) for task in list_tasks(limit=500)]

    def matches(item: dict) -> bool:
        if status and item.get("status", "").lower() != status:
            return False
        if keyword:
            hay = " ".join([
                str(item.get("id", "")),
                str(item.get("name", "")),
                str((item.get("source_files") or {}).get("chart1", "")),
                str((item.get("source_files") or {}).get("chart2", "")),
            ]).lower()
            if keyword not in hay:
                return False
        return True

    filtered = [item for item in items if matches(item)]
    filtered = sorted(filtered, key=lambda x: x.get("updated_at") or "", reverse=True)[:limit]
    return jsonify({"ok": True, "tasks": filtered, "count": len(filtered)})


@app.route("/api/tasks/<task_id>/clone", methods=["POST"])
def clone_task(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404

    new_task = copy.deepcopy(task)
    new_id = uuid.uuid4().hex[:12]
    new_name = (request.get_json(silent=True) or {}).get("name")
    new_task["id"] = new_id
    new_task["name"] = (str(new_name).strip() if new_name else f"{task.get('name') or '任務'}（複製）")
    new_task["status"] = "ready" if task.get("master_rows") else "processing"
    new_task["created_at"] = now_iso()
    new_task["updated_at"] = now_iso()
    new_task["review_decisions"] = {}
    new_task["candidate_decisions"] = {}
    new_task["error"] = ""
    save_task(new_task)
    return jsonify({"ok": True, "task": new_task}), 201


@app.route("/api/tasks/analyze", methods=["POST"])
def analyze():
    import shutil
    import threading

    chart1 = request.files.get("chart1")
    chart2 = request.files.get("chart2")
    if not chart1 or not chart2:
        missing = [name for name, value in (("chart1", chart1), ("chart2", chart2)) if not value]
        print(f"[API] /api/tasks/analyze rejected: missing files={missing}", flush=True)
        return jsonify({"error": "chart1_and_chart2_required", "message": "請同時上傳圖一和圖二。"}), 400

    chart1_size_mb = _file_size_mb(chart1)
    if chart1_size_mb > MAX_CHART1_MB:
        return jsonify({"error": "chart1_too_large", "message": f"圖一檔案不可超過 {MAX_CHART1_MB}MB。"}), 400
    c1_shape = _inspect_image(chart1)
    if c1_shape and max(c1_shape["width"], c1_shape["height"]) > MAX_CHART1_LONG_EDGE:
        return jsonify({"error": "chart1_edge_too_large", "message": f"圖一最長邊不可超過 {MAX_CHART1_LONG_EDGE}px。"}), 400

    c2_shape = _inspect_image(chart2)
    if c2_shape:
        c2_chunks = _estimate_chart2_chunks(c2_shape["width"], c2_shape["height"])
        if c2_chunks > MAX_CHART2_CHUNKS:
            return jsonify({"error": "chart2_too_many_chunks", "message": f"圖二預估分塊為 {c2_chunks} 段，超過上限 {MAX_CHART2_CHUNKS} 段，請先裁切後再上傳。"}), 400

    if not os.environ.get("DASHSCOPE_API_KEY", "").strip():
        return jsonify({"error": "no_api_key", "message": "伺服器尚未設定 AI 辨識 API Key，無法分析。請聯絡管理員。"}), 422

    task_name = request.form.get("task_name", "").strip()

    # 儲存上傳圖片，立刻建立 processing 狀態的任務
    task_id = uuid.uuid4().hex[:12]
    upload_dir = task_upload_dir(task_id)
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
        "chart2_progress": {},
        "error": "",
        "cancel_requested": False,
    }
    save_task(task)

    # 背景執行兩段式 Qwen-VL 分析
    def run_async():
        # ── 第一階段：圖一骨架 ──────────────────────────────────
        try:
            from analyzer import run_chart1_stage
            stage1 = run_chart1_stage(c1_path)
            if is_task_cancel_requested(task_id):
                task["status"] = "cancelled"
                task["error"] = "任務已取消"
                save_task(task)
                return
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

        # ── 第二階段：圖二 OCR（辨識完成後暫停，等用戶確認再 merge）
        try:
            from analyzer import analyze_chart2
            task["status"] = "processing_chart2"
            task["chart2_progress"] = {
                "status": "queued",
                "current_chunk": 0,
                "total_chunks": 0,
                "rows_so_far": 0,
                "failed_chunks": [],
                "updated_at": now_iso(),
            }
            save_task(task)
            chart2_attrs = analyze_chart2(
                c2_path,
                progress_callback=make_chart2_progress_saver(task),
                should_continue=lambda: not is_task_cancel_requested(task_id),
            )
            task["status"] = "chart2_ocr_done"
            task["chart2_raw"] = chart2_attrs
            task["chart2_progress"] = {
                **task.get("chart2_progress", {}),
                "status": task.get("chart2_progress", {}).get("status") or "done",
                "deduped_count": len(chart2_attrs),
                "updated_at": now_iso(),
            }
            task["error"] = ""
        except InterruptedError:
            task["status"] = "cancelled"
            task["error"] = "任務已取消"
        except Exception as exc:
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
        print(f"[API] /api/tasks/{task_id}/analyze-chart2 rejected: missing chart2 file", flush=True)
        return jsonify({"error": "chart2_required", "message": "請重新選擇圖二圖片後再上傳。"}), 400

    c2_shape = _inspect_image(chart2)
    if c2_shape:
        c2_chunks = _estimate_chart2_chunks(c2_shape["width"], c2_shape["height"])
        if c2_chunks > MAX_CHART2_CHUNKS:
            return jsonify({"error": "chart2_too_many_chunks", "message": f"圖二預估分塊為 {c2_chunks} 段，超過上限 {MAX_CHART2_CHUNKS} 段，請先裁切後再上傳。"}), 400

    # 存新的圖二
    upload_dir = task_upload_dir(task_id)
    c2_path = upload_dir / f"chart2_retry_{sanitize_filename(chart2.filename or 'upload.jpg')}"
    chart2.save(str(c2_path))

    # 標記為處理中，保留現有骨架
    task["status"] = "processing_chart2"
    task["error"] = ""
    task["cancel_requested"] = False
    task["source_files"]["chart2"] = chart2.filename or ""
    task["chart2_progress"] = {
        "status": "queued",
        "current_chunk": 0,
        "total_chunks": 0,
        "rows_so_far": 0,
        "failed_chunks": [],
        "updated_at": now_iso(),
    }
    save_task(task)

    def run_async():
        try:
            from analyzer import analyze_chart2
            chart2_attrs = analyze_chart2(
                c2_path,
                progress_callback=make_chart2_progress_saver(task),
                should_continue=lambda: not is_task_cancel_requested(task_id),
            )
            task["status"] = "chart2_ocr_done"
            task["chart2_raw"] = chart2_attrs
            task["chart2_progress"] = {
                **task.get("chart2_progress", {}),
                "status": task.get("chart2_progress", {}).get("status") or "done",
                "deduped_count": len(chart2_attrs),
                "updated_at": now_iso(),
            }
            task["error"] = ""
        except InterruptedError:
            task["status"] = "cancelled"
            task["error"] = "任務已取消"
        except Exception as exc:
            task["status"] = "chart2_error"
            task["error"] = f"圖二辨識失敗：{exc}"
        finally:
            save_task(task)

    threading.Thread(target=run_async, daemon=True).start()
    return jsonify({"id": task_id, "status": "processing_chart2"}), 202


@app.route("/api/tasks/<task_id>/cancel", methods=["POST"])
def cancel_task(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    if task.get("status") in {"ready", "chart2_ocr_done", "chart2_error", "error", "cancelled"}:
        return jsonify({"ok": True, "status": task.get("status"), "message": "任務目前無需取消。"})
    task["cancel_requested"] = True
    task["status"] = "cancel_requested"
    task["error"] = "取消中：將於當前分塊完成後停止。"
    save_task(task)
    return jsonify({"ok": True, "status": "cancel_requested"})


@app.route("/api/tasks/<task_id>/confirm-chart2", methods=["POST"])
def confirm_chart2(task_id: str):
    """用戶確認圖二辨識結果後，觸發 merge（純 Python，瞬間完成）。"""
    payload = request.get_json(silent=True) or {}
    task = read_task(task_id)
    if not task:
        snapshot = payload.get("task_snapshot") if isinstance(payload, dict) else None
        if (
            isinstance(snapshot, dict)
            and snapshot.get("id") == task_id
            and isinstance(snapshot.get("master_rows"), list)
            and isinstance(snapshot.get("chart2_raw"), list)
        ):
            task = snapshot
            task["status"] = "chart2_ocr_done"
            task.setdefault("review_decisions", {})
            task.setdefault("candidate_decisions", {})
            task.setdefault("review_rows", [])
            task.setdefault("candidate_rows", [])
            task.setdefault("graph", {})
            task.setdefault("summary", {})
            task.setdefault("error", "")
            save_task(task)
        else:
            return jsonify({"error": "task_not_found"}), 404
    if task.get("status") != "chart2_ocr_done":
        return jsonify({"error": "invalid_status", "message": "任務狀態不是 chart2_ocr_done，無法確認。"}), 400

    chart2_attrs = task.get("chart2_raw", [])

    try:
        from analyzer import enrich_with_chart2_precomputed
        stage2 = enrich_with_chart2_precomputed(task["master_rows"], chart2_attrs)
        task["status"] = "ready"
        task["summary"] = stage2["summary"]
        task["master_rows"] = stage2["master_rows"]
        task["review_rows"] = stage2["review_rows"]
        task["candidate_rows"] = stage2["candidate_rows"]
        task["graph"] = stage2["graph"]
        task["error"] = ""
        save_task(task)
        return jsonify({
            "ok": True,
            "master_rows": task["master_rows"],
            "review_rows": task["review_rows"],
            "candidate_rows": task["candidate_rows"],
            "summary": task["summary"],
            "graph": task["graph"],
        })
    except Exception as exc:
        task["status"] = "chart2_error"
        task["error"] = f"配對失敗：{exc}"
        save_task(task)
        return jsonify({"error": "match_failed", "message": str(exc)}), 500


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


@app.route("/api/tasks/<task_id>/review-confirm-all", methods=["POST"])
def review_confirm_all(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404

    confirmed = 0
    for review in list(task.get("review_rows", [])):
        key = review.get("candidate_node_id") or review.get("chart2_name")
        if not key:
            continue
        task.setdefault("review_decisions", {})[key] = {
            "decision": "確認一致",
            "corrected_name": "",
            "corrected_level": "",
            "corrected_parent": "",
            "note": "批次確認：圖一已辨識，視為可用。",
        }
        target_row = next((row for row in task.get("master_rows", []) if row.get("node_id") == key), None)
        if not target_row:
            continue
        matched_candidate = next(
            (
                row for row in task.get("candidate_rows", [])
                if row.get("chart2_name") and row.get("chart2_name") == target_row.get("matched_chart2_name")
            ),
            None,
        )
        if matched_candidate:
            apply_chart2_attrs_to_row(target_row, matched_candidate)
            task["candidate_rows"] = [
                row for row in task.get("candidate_rows", [])
                if row.get("chart2_name") != matched_candidate.get("chart2_name")
            ]
        target_row["node_status"] = "enriched"
        target_row["match_status"] = target_row.get("match_status") or "confirmed_chart1"
        target_row["review_flag"] = ""
        target_row["review_note"] = "批次確認：圖一已辨識，視為可用。"
        confirmed += 1

    task["review_rows"] = []
    rebuild_task_state(task)
    save_task(task)
    return jsonify({
        "ok": True,
        "confirmed_count": confirmed,
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
                "subsidiary_level_label", "role_label", "chart_note", "sort_index"]

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


@app.route("/api/tasks/<task_id>/replace-state", methods=["POST"])
def replace_task_state(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    payload = request.get_json(silent=True) or {}
    master_rows = payload.get("master_rows")
    if not isinstance(master_rows, list):
        return jsonify({"error": "master_rows_required"}), 400

    task["master_rows"] = master_rows
    if isinstance(payload.get("review_rows"), list):
        task["review_rows"] = payload.get("review_rows")
    if isinstance(payload.get("candidate_rows"), list):
        task["candidate_rows"] = payload.get("candidate_rows")
    if isinstance(payload.get("review_decisions"), dict):
        task["review_decisions"] = payload.get("review_decisions")
    if isinstance(payload.get("candidate_decisions"), dict):
        task["candidate_decisions"] = payload.get("candidate_decisions")

    rebuild_task_state(task)
    save_task(task)
    return jsonify({
        "ok": True,
        "master_rows": task["master_rows"],
        "review_rows": task["review_rows"],
        "candidate_rows": task["candidate_rows"],
        "review_decisions": task["review_decisions"],
        "candidate_decisions": task["candidate_decisions"],
        "summary": task["summary"],
        "graph": task["graph"],
    })


@app.route("/api/tasks/<task_id>/batch-update", methods=["POST"])
def batch_update_rows(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    payload = request.get_json(silent=True) or {}
    node_ids = payload.get("node_ids")
    field = str(payload.get("field") or "").strip()
    value = payload.get("value")
    if not isinstance(node_ids, list) or not node_ids:
        return jsonify({"error": "node_ids_required", "message": "請先勾選要批次更新的公司。"}), 400

    editable = {
        "canonical_name", "legal_representative", "registered_capital",
        "established_date", "actual_controller_share", "role_label", "chart_note",
    }
    if field not in editable:
        return jsonify({"error": "invalid_field", "message": "這個欄位目前不支援批次編輯。"}), 400

    id_set = {str(node_id) for node_id in node_ids if str(node_id).strip()}
    updated = 0
    for row in task.get("master_rows", []):
        if str(row.get("node_id")) not in id_set:
            continue
        row[field] = str(value or "").strip()
        if field == "canonical_name":
            refresh_children_parent_names(task["master_rows"], row["node_id"], row[field])
        updated += 1

    rebuild_task_state(task)
    save_task(task)
    return jsonify({
        "ok": True,
        "updated_count": updated,
        "master_rows": task["master_rows"],
        "review_rows": task["review_rows"],
        "candidate_rows": task["candidate_rows"],
        "summary": task["summary"],
        "graph": task["graph"],
    })


@app.route("/api/tasks/<task_id>/governance-preview", methods=["POST"])
def governance_preview(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    payload = request.get_json(silent=True) or {}
    apply_changes = bool(payload.get("apply"))

    updates: list[dict] = []
    dup_map: dict[str, list[str]] = {}
    for row in task.get("master_rows", []):
        before = str(row.get("canonical_name") or row.get("chart1_name") or "")
        after = _normalize_company_text(before)
        if after:
            dup_map.setdefault(after.lower(), []).append(str(row.get("node_id") or ""))
        if after != before:
            updates.append({
                "node_id": row.get("node_id"),
                "field": "canonical_name",
                "before": before,
                "after": after,
            })
            if apply_changes:
                row["canonical_name"] = after
                refresh_children_parent_names(task["master_rows"], row.get("node_id"), after)

    duplicate_groups = []
    for norm_name, ids in dup_map.items():
        if len(ids) < 2:
            continue
        duplicate_groups.append({
            "name": norm_name,
            "node_ids": ids,
        })

    if apply_changes and updates:
        rebuild_task_state(task)
        save_task(task)

    return jsonify({
        "ok": True,
        "applied": apply_changes,
        "changes": updates,
        "change_count": len(updates),
        "duplicate_groups": duplicate_groups,
        "master_rows": task.get("master_rows", []) if apply_changes else None,
        "review_rows": task.get("review_rows", []) if apply_changes else None,
        "candidate_rows": task.get("candidate_rows", []) if apply_changes else None,
        "summary": task.get("summary", {}) if apply_changes else None,
        "graph": task.get("graph", {}) if apply_changes else None,
    })


@app.route("/api/tasks/<task_id>/add-row", methods=["POST"])
def add_row(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    payload = request.get_json(silent=True) or {}
    name = str(payload.get("canonical_name") or payload.get("chart1_name") or "").strip()
    if not name:
        return jsonify({"error": "company_name_required"}), 400

    parent_id = str(payload.get("chart1_parent") or "").strip()
    parent_row = next((row for row in task.get("master_rows", []) if row.get("node_id") == parent_id), None)
    root_row = next((row for row in task.get("master_rows", []) if parse_level_value(row.get("chart1_level")) == 0), None)
    if not parent_row and root_row:
        parent_row = root_row

    parent_level = parse_level_value(parent_row.get("chart1_level")) if parent_row else None
    level = (parent_level + 1) if parent_level is not None else 0
    sibling_count = sum(1 for row in task.get("master_rows", []) if (row.get("chart1_parent") or "") == (parent_row.get("node_id", "") if parent_row else ""))
    new_row = {
        "node_id": f"M{uuid.uuid4().hex[:6].upper()}",
        "chart1_name": name,
        "canonical_name": name,
        "chart1_level": level,
        "chart1_parent": parent_row.get("node_id", "") if parent_row else "",
        "chart1_parent_name": parent_row.get("canonical_name", "") or parent_row.get("chart1_name", "") if parent_row else "",
        "sort_index": sibling_count + 1,
        "matched_chart2_name": "",
        "legal_representative": str(payload.get("legal_representative") or "").strip(),
        "established_date": str(payload.get("established_date") or "").strip(),
        "registered_capital": str(payload.get("registered_capital") or "").strip(),
        "actual_controller_share": str(payload.get("actual_controller_share") or "").strip(),
        "subsidiary_level_label": level_label(level),
        "company_status": "",
        "role_label": str(payload.get("role_label") or "").strip(),
        "chart_note": str(payload.get("chart_note") or "").strip(),
        "match_status": "manual",
        "node_status": "manual_added",
        "review_flag": "manual_added",
        "review_note": "人工新增",
    }
    task.setdefault("master_rows", []).append(new_row)
    rebuild_task_state(task)
    save_task(task)
    return jsonify({
        "ok": True,
        "added_node_id": new_row["node_id"],
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


@app.route("/api/tasks/<task_id>/batch-delete", methods=["POST"])
def batch_delete_rows(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    payload = request.get_json(silent=True) or {}
    node_ids = payload.get("node_ids")
    if not isinstance(node_ids, list) or not node_ids:
        return jsonify({"error": "node_ids_required", "message": "請先勾選要刪除的公司。"}), 400

    delete_ids = {str(node_id) for node_id in node_ids if str(node_id).strip()}
    deleted = 0
    kept_rows = []
    for row in task.get("master_rows", []):
        if str(row.get("node_id")) in delete_ids:
            deleted += 1
            continue
        kept_rows.append(row)

    for row in kept_rows:
        if str(row.get("chart1_parent")) in delete_ids:
            row["chart1_parent"] = ""
            row["chart1_parent_name"] = ""
            row["chart1_level"] = 0
            row["subsidiary_level_label"] = level_label(0)

    task["master_rows"] = kept_rows
    rebuild_task_state(task)
    save_task(task)
    return jsonify({
        "ok": True,
        "deleted_count": deleted,
        "master_rows": task["master_rows"],
        "review_rows": task["review_rows"],
        "candidate_rows": task["candidate_rows"],
        "summary": task["summary"],
        "graph": task["graph"],
    })


@app.route("/api/tasks/<task_id>/chart-shareholders", methods=["POST"])
def update_chart_shareholders(task_id: str):
    task = read_task(task_id)
    if not task:
        return jsonify({"error": "task_not_found"}), 404
    payload = request.get_json(silent=True) or {}
    rows = payload.get("chart_shareholders")
    if not isinstance(rows, list):
        return jsonify({"error": "chart_shareholders_required", "message": "請提供上層股東清單。"}), 400

    valid_targets = {str(row.get("node_id")) for row in task.get("master_rows", []) if row.get("node_id")}
    cleaned = []
    for idx, row in enumerate(rows):
        if not isinstance(row, dict):
            continue
        name = str(row.get("name") or "").strip()
        target_node_id = str(row.get("target_node_id") or "").strip()
        if not name or target_node_id not in valid_targets:
            continue
        holder_type = str(row.get("type") or "company").strip()
        if holder_type not in {"company", "person"}:
            holder_type = "company"
        cleaned.append({
            "id": str(row.get("id") or f"holder_{idx + 1}").strip(),
            "name": name,
            "type": holder_type,
            "share": str(row.get("share") or "").strip(),
            "target_node_id": target_node_id,
            "note": str(row.get("note") or "").strip(),
        })

    task["chart_shareholders"] = cleaned
    save_task(task)
    return jsonify({
        "ok": True,
        "chart_shareholders": task["chart_shareholders"],
        "master_rows": task.get("master_rows", []),
        "review_rows": task.get("review_rows", []),
        "candidate_rows": task.get("candidate_rows", []),
        "summary": task.get("summary", {}),
        "graph": task.get("graph", {}),
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
            exists = next(
                (
                    row for row in task.get("master_rows", [])
                    if row.get("matched_chart2_name") == key
                    or row.get("canonical_name") == (corrected_name or candidate.get("chart2_name") or candidate.get("company") or "")
                ),
                None,
            )
            if exists:
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
            parent_row = find_row_by_name(task.get("master_rows", []), parent_name) if parent_name else None
            parent_level = parse_level_value(parent_row.get("chart1_level")) if parent_row else None
            level = (parent_level + 1) if parent_level is not None else parse_level_value(candidate.get("subsidiary_level_label")) or 0
            final_name = corrected_name or candidate.get("chart2_name") or candidate.get("company") or ""
            sibling_count = sum(1 for row in task.get("master_rows", []) if (row.get("chart1_parent") or "") == (parent_row.get("node_id", "") if parent_row else ""))
            new_row = {
                "node_id": f"A{uuid.uuid4().hex[:6].upper()}",
                "chart1_name": final_name,
                "canonical_name": final_name,
                "chart1_level": level,
                "chart1_parent": parent_row.get("node_id", "") if parent_row else "",
                "chart1_parent_name": parent_row.get("canonical_name", "") if parent_row else parent_name,
                "sort_index": sibling_count + 1,
                "matched_chart2_name": candidate.get("chart2_name", ""),
                "legal_representative": candidate.get("legal_representative", ""),
                "established_date": candidate.get("established_date", ""),
                "registered_capital": candidate.get("registered_capital", ""),
                "actual_controller_share": candidate.get("actual_controller_share", ""),
                "subsidiary_level_label": candidate.get("subsidiary_level_label") or level_label(level),
                "company_status": candidate.get("company_status", ""),
                "role_label": "",
                "chart_note": "",
                "match_status": "matched",
                "node_status": "enriched",
                "review_flag": "",
                "review_note": note,
            }
            task["master_rows"].append(new_row)

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
    ensure_storage_ready()
    port = int(os.environ.get("PORT", 8000))
    app.run(host="0.0.0.0", port=port)
