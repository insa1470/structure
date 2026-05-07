from __future__ import annotations

import copy
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


class PostgresTaskStore:
    """PostgreSQL-backed task storage.

    This first database adapter stores the complete task JSON to preserve the
    current API contract. Later migrations can split master/review/candidate
    rows into normalized tables without changing route handlers again.
    """

    def __init__(self, database_url: str | None = None, data_dir: Path = DATA_DIR):
        self.database_url = database_url or os.environ.get("DATABASE_URL", "")
        self.data_dir = data_dir
        self._schema_ready = False

    def upload_dir(self, task_id: str) -> Path:
        path = self.data_dir / task_id / "uploads"
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _connect(self):
        if not self.database_url:
            raise RuntimeError("TASK_STORE=postgres requires DATABASE_URL")
        try:
            import psycopg
        except ImportError as exc:
            raise RuntimeError("TASK_STORE=postgres requires psycopg[binary]") from exc
        return psycopg.connect(self.database_url)

    def ensure_ready(self) -> None:
        self.data_dir.mkdir(parents=True, exist_ok=True)
        if self._schema_ready:
            return
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS tasks (
                        id TEXT PRIMARY KEY,
                        name TEXT,
                        status TEXT,
                        created_at TIMESTAMPTZ,
                        updated_at TIMESTAMPTZ,
                        source_files JSONB NOT NULL DEFAULT '{}'::jsonb,
                        summary JSONB NOT NULL DEFAULT '{}'::jsonb,
                        error TEXT NOT NULL DEFAULT '',
                        task JSONB NOT NULL
                    )
                """)
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS task_files (
                        id BIGSERIAL PRIMARY KEY,
                        task_id TEXT NOT NULL REFERENCES tasks(id) ON DELETE CASCADE,
                        file_type TEXT NOT NULL,
                        original_filename TEXT NOT NULL DEFAULT '',
                        storage_path TEXT NOT NULL DEFAULT '',
                        width INTEGER,
                        height INTEGER,
                        file_size BIGINT,
                        created_at TIMESTAMPTZ NOT NULL DEFAULT now()
                    )
                """)
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS task_snapshots (
                        id BIGSERIAL PRIMARY KEY,
                        task_id TEXT NOT NULL REFERENCES tasks(id) ON DELETE CASCADE,
                        task JSONB NOT NULL,
                        created_at TIMESTAMPTZ NOT NULL DEFAULT now()
                    )
                """)
                cur.execute("CREATE INDEX IF NOT EXISTS idx_tasks_updated_at ON tasks(updated_at DESC)")
                cur.execute("CREATE INDEX IF NOT EXISTS idx_task_snapshots_task_id ON task_snapshots(task_id, created_at DESC)")
            conn.commit()
        self._schema_ready = True

    def read_task(self, task_id: str) -> dict | None:
        self.ensure_ready()
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT task FROM tasks WHERE id = %s", (task_id,))
                row = cur.fetchone()
                return row[0] if row else None

    def save_task(self, task: dict) -> None:
        self.ensure_ready()
        task["updated_at"] = now_iso()
        payload = json.dumps(task, ensure_ascii=False)
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO tasks (id, name, status, created_at, updated_at, source_files, summary, error, task)
                    VALUES (%s, %s, %s, %s, %s, %s::jsonb, %s::jsonb, %s, %s::jsonb)
                    ON CONFLICT (id) DO UPDATE SET
                        name = EXCLUDED.name,
                        status = EXCLUDED.status,
                        updated_at = EXCLUDED.updated_at,
                        source_files = EXCLUDED.source_files,
                        summary = EXCLUDED.summary,
                        error = EXCLUDED.error,
                        task = EXCLUDED.task
                    """,
                    (
                        task["id"],
                        task.get("name", ""),
                        task.get("status", ""),
                        task.get("created_at") or task.get("updated_at"),
                        task.get("updated_at"),
                        json.dumps(task.get("source_files") or {}, ensure_ascii=False),
                        json.dumps(task.get("summary") or {}, ensure_ascii=False),
                        task.get("error", ""),
                        payload,
                    ),
                )
                cur.execute(
                    "INSERT INTO task_snapshots (task_id, task) VALUES (%s, %s::jsonb)",
                    (task["id"], payload),
                )
            conn.commit()

    def iter_tasks(self) -> Iterable[dict]:
        self.ensure_ready()
        with self._connect() as conn:
            with conn.cursor() as cur:
                cur.execute("SELECT task FROM tasks ORDER BY updated_at DESC NULLS LAST LIMIT 500")
                for row in cur.fetchall():
                    yield row[0]


class MirroredTaskStore:
    """Primary task store with best-effort mirror writes.

    Reads and uploads stay on the primary store. Mirror failures are logged but
    never block the current production flow.
    """

    def __init__(self, primary, mirror):
        self.primary = primary
        self.mirror = mirror

    def upload_dir(self, task_id: str) -> Path:
        return self.primary.upload_dir(task_id)

    def read_task(self, task_id: str) -> dict | None:
        return self.primary.read_task(task_id)

    def save_task(self, task: dict) -> None:
        self.primary.save_task(task)
        try:
            self.mirror.save_task(copy.deepcopy(task))
        except Exception as exc:
            print(f"[Storage] mirror save failed for task {task.get('id', '')}: {exc}", flush=True)

    def ensure_ready(self) -> None:
        self.primary.ensure_ready()
        try:
            self.mirror.ensure_ready()
        except Exception as exc:
            print(f"[Storage] mirror ensure failed: {exc}", flush=True)

    def iter_tasks(self) -> Iterable[dict]:
        return self.primary.iter_tasks()


def make_task_store():
    store_name = os.environ.get("TASK_STORE", "json").strip().lower()
    mirror_name = os.environ.get("TASK_STORE_MIRROR", "").strip().lower()
    if store_name == "postgres":
        return PostgresTaskStore()
    primary = JsonTaskStore()
    if mirror_name == "postgres":
        return MirroredTaskStore(primary, PostgresTaskStore())
    return primary


task_store = make_task_store()


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
