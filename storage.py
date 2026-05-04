from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "app_data" / "tasks"


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
        path.write_text(json.dumps(task, ensure_ascii=False, indent=2), encoding="utf-8")

    def ensure_ready(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)


task_store = JsonTaskStore()


def read_task(task_id: str) -> dict | None:
    return task_store.read_task(task_id)


def save_task(task: dict) -> None:
    task_store.save_task(task)


def task_upload_dir(task_id: str) -> Path:
    return task_store.upload_dir(task_id)


def ensure_storage_ready() -> None:
    task_store.ensure_ready()
