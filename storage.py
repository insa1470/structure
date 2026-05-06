from __future__ import annotations

import json
import os
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "app_data" / "tasks"
OCR_TEST_DIR = BASE_DIR / "app_data" / "ocr_tests"


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


class JsonTaskStore:
    """Current task storage adapter.

    This keeps the existing JSON files, while isolating task persistence so a
    database-backed adapter can replace it later without changing every route.
    """

    def __init__(self, data_dir: Path = DATA_DIR):
        self.data_dir = data_dir

    def task_path(self, task_id: str) -> Path:
        return self.data_dir / task_id / "task.json"

    def upload_dir(self, task_id: str) -> Path:
        path = self.data_dir / task_id / "uploads"
        path.mkdir(parents=True, exist_ok=True)
        return path

    def read_task(self, task_id: str) -> dict | None:
        path = self.task_path(task_id)
        if not path.exists():
            return None
        return json.loads(path.read_text(encoding="utf-8"))

    def save_task(self, task: dict) -> None:
        task["updated_at"] = now_iso()
        path = self.task_path(task["id"])
        path.parent.mkdir(parents=True, exist_ok=True)
        payload = json.dumps(task, ensure_ascii=False, indent=2)
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            dir=path.parent,
            prefix=".task.",
            suffix=".tmp",
            delete=False,
        ) as f:
            tmp_name = f.name
            f.write(payload)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_name, path)

    def ensure_ready(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)

    def iter_tasks(self) -> Iterable[dict]:
        self.ensure_ready()
        for task_dir in sorted(self.data_dir.iterdir(), reverse=True):
            if not task_dir.is_dir():
                continue
            path = task_dir / "task.json"
            if not path.exists():
                continue
            try:
                yield json.loads(path.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                continue


task_store = JsonTaskStore()


class JsonOcrTestStore:
    def __init__(self, data_dir: Path = OCR_TEST_DIR):
        self.data_dir = data_dir

    def ensure_ready(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)

    def save_test(self, record: dict) -> None:
        self.ensure_ready()
        path = self.data_dir / f"{record['id']}.json"
        payload = json.dumps(record, ensure_ascii=False, indent=2)
        with tempfile.NamedTemporaryFile(
            "w",
            encoding="utf-8",
            dir=path.parent,
            prefix=".ocr.",
            suffix=".tmp",
            delete=False,
        ) as f:
            tmp_name = f.name
            f.write(payload)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_name, path)

    def list_tests(self, limit: int = 50) -> list[dict]:
        self.ensure_ready()
        records: list[dict] = []
        for path in sorted(self.data_dir.glob("*.json"), reverse=True):
            try:
                data = json.loads(path.read_text(encoding="utf-8"))
            except json.JSONDecodeError:
                continue
            raw = data.pop("raw_response", None)
            if raw is not None:
                data["raw_response_preview"] = str(raw)[:800]
            records.append(data)
            if len(records) >= limit:
                break
        return records


ocr_test_store = JsonOcrTestStore()


def read_task(task_id: str) -> dict | None:
    return task_store.read_task(task_id)


def save_task(task: dict) -> None:
    task_store.save_task(task)


def task_upload_dir(task_id: str) -> Path:
    return task_store.upload_dir(task_id)


def ensure_storage_ready() -> None:
    task_store.ensure_ready()
    ocr_test_store.ensure_ready()


def save_ocr_test(record: dict) -> None:
    ocr_test_store.save_test(record)


def list_ocr_tests(limit: int = 50) -> list[dict]:
    return ocr_test_store.list_tests(limit)


def list_tasks(limit: int = 200) -> list[dict]:
    tasks: list[dict] = []
    for task in task_store.iter_tasks():
        tasks.append(task)
        if len(tasks) >= limit:
            break
    return tasks
